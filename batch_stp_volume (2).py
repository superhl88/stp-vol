import argparse
import concurrent.futures
import csv
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    from OCP.BRepGProp import BRepGProp
    from OCP.GProp import GProp_GProps
    from OCP.IFSelect import IFSelect_RetDone
    from OCP.STEPControl import STEPControl_Reader
    from OCP.TopoDS import TopoDS_Shape

    HAS_OCP = True
except Exception as e:
    HAS_OCP = False
    OCP_IMPORT_ERR = str(e)

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False


STEP_EXTS = {".stp", ".step"}


def _clean_user_path(raw: str) -> str:
    """清理用户输入的路径（去首尾空格、引号，处理 'E:' 这种盘符写法）"""
    s = raw.strip().strip('"').strip("'")
    if len(s) == 2 and s[0].isalpha() and s[1] == ":":
        s = s + "\\"
    return s


def prompt_interactive_paths() -> Tuple[str, str]:
    """交互式引导选择扫描目标（文件夹/磁盘）与输出目录"""
    print("=" * 64)
    print("STEP 批量体积计算工具（支持文件夹/磁盘扫描 + 自定义输出目录）")
    print("=" * 64)

    while True:
        print("\n请选择扫描目标类型：")
        print("  1. 扫描指定文件夹")
        print("  2. 扫描整个磁盘（如 E: 或 E:\\ 或 Z:）")
        choice = input("请输入选项 [1/2，默认 1]: ").strip() or "1"

        if choice == "1":
            raw_src = input("请输入要搜索的文件夹路径（例如 E:\\图纸 或 Z:\\B-工程管理\\航空项目）: ")
        elif choice == "2":
            raw_src = input("请输入要扫描的磁盘盘符或根目录（例如 E: 或 E:\\）: ")
        else:
            print("输入无效，请输入 1 或 2。")
            continue

        src_path = _clean_user_path(raw_src)
        if not src_path:
            print("路径不能为空，请重新输入。")
            continue

        p = Path(src_path)
        if not p.exists():
            print(f"路径不存在或当前不可访问: {src_path}")
            continue
        if not p.is_dir():
            print(f"该路径不是文件夹/磁盘根目录: {src_path}")
            continue

        scan_root = str(p.resolve()) if len(src_path) > 3 else str(p)
        break

    default_out = str(Path.cwd() / "output")
    while True:
        raw_out = input(f"\n请输入输出文件存放位置 [直接回车默认: {default_out}]: ")
        out_str = _clean_user_path(raw_out) if raw_out.strip() else default_out
        try:
            out_dir = Path(out_str)
            out_dir.mkdir(parents=True, exist_ok=True)
            break
        except Exception as e:
            print(f"无法创建或访问输出目录 '{out_str}': {e}")

    print("-" * 64)
    print(f"扫描目标 : {scan_root}")
    print(f"输出目录 : {out_dir}")
    print("-" * 64)
    return scan_root, str(out_dir)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="批量计算 STEP (.stp/.step) 文件体积")
    parser.add_argument("--root", type=str, default=None, help="要扫描的文件夹或磁盘路径")
    parser.add_argument("--output-dir", type=str, default=None, help="结果与断点文件存放目录")
    parser.add_argument("--workers", type=int, default=min(8, max(2, (os.cpu_count() or 4))), help="并发工作进程数")
    parser.add_argument("--timeout", type=float, default=45.0, help="单个文件处理超时时间（秒）")
    parser.add_argument("--chunk-size", type=int, default=10000, help="每个结果文件的最大记录数")
    parser.add_argument("--no-interactive", action="store_true", help="禁用交互输入，未传 --root 时使用当前目录")
    return parser.parse_args()


def discover_step_files(root_dir: str) -> List[str]:
    """递归扫描目录下所有 .stp / .step 文件，跳过常见系统/隐藏/输出目录"""
    skip_dirs = {
        "$recycle.bin",
        "system volume information",
        ".git",
        "__pycache__",
        ".venv",
        "venv",
        "build",
        "dist",
    }
    found: List[str] = []
    root_resolved = str(Path(root_dir).resolve()) if len(root_dir) > 3 else root_dir

    for current_root, dirnames, filenames in os.walk(root_resolved, topdown=True, followlinks=False):
        dirnames[:] = [
            d for d in dirnames
            if d.lower() not in skip_dirs and not d.startswith(".")
        ]
        for fname in filenames:
            ext = Path(fname).suffix.lower()
            if ext in STEP_EXTS:
                found.append(os.path.join(current_root, fname))

    found.sort()
    return found


def calc_step_volume_worker(file_path: str) -> Dict[str, Any]:
    """子进程执行单个 STEP 文件体积计算（mm^3 -> cm^3）"""
    start_t = time.perf_counter()
    record: Dict[str, Any] = {
        "file_path": file_path,
        "file_name": os.path.basename(file_path),
        "file_size_kb": round(os.path.getsize(file_path) / 1024.0, 2) if os.path.exists(file_path) else 0.0,
        "status": "failed",
        "volume_mm3": None,
        "volume_cm3": None,
        "solid_count": 0,
        "elapsed_sec": 0.0,
        "error": "",
    }

    if not HAS_OCP:
        record["error"] = f"OCP 未加载: {OCP_IMPORT_ERR}"
        record["elapsed_sec"] = round(time.perf_counter() - start_t, 3)
        return record

    try:
        reader = STEPControl_Reader()
        status = reader.ReadFile(file_path)
        if status != IFSelect_RetDone:
            record["error"] = f"STEP ReadFile 失败 (status={status})"
            record["elapsed_sec"] = round(time.perf_counter() - start_t, 3)
            return record

        reader.TransferRoots()
        shape: TopoDS_Shape = reader.OneShape()
        if shape.IsNull():
            record["error"] = "解析结果为空 Shape (IsNull)"
            record["elapsed_sec"] = round(time.perf_counter() - start_t, 3)
            return record

        props = GProp_GProps()
        BRepGProp.VolumeProperties_s(shape, props)
        vol_mm3 = float(props.Mass())
        if vol_mm3 < 0:
            vol_mm3 = abs(vol_mm3)

        record["status"] = "success"
        record["volume_mm3"] = round(vol_mm3, 4)
        record["volume_cm3"] = round(vol_mm3 / 1000.0, 6)
        record["elapsed_sec"] = round(time.perf_counter() - start_t, 3)
        return record

    except Exception as e:
        record["error"] = f"{type(e).__name__}: {e}"
        record["elapsed_sec"] = round(time.perf_counter() - start_t, 3)
        return record


def load_checkpoint(checkpoint_path: Path) -> Dict[str, Dict[str, Any]]:
    """读取历史断点记录"""
    done: Dict[str, Dict[str, Any]] = {}
    if not checkpoint_path.exists():
        return done

    with checkpoint_path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                item = json.loads(line)
                fp = item.get("file_path")
                if fp:
                    done[fp] = item
            except json.JSONDecodeError:
                continue
    return done


def append_checkpoint(checkpoint_path: Path, item: Dict[str, Any]) -> None:
    """追加写入单条断点记录"""
    with checkpoint_path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(item, ensure_ascii=False) + "\n")
        f.flush()


def export_excel_chunks(records: List[Dict[str, Any]], out_dir: Path, chunk_size: int) -> List[Path]:
    """按 chunk_size 分卷导出 Excel"""
    created_files: List[Path] = []
    headers = [
        "序号",
        "文件名",
        "完整路径",
        "文件大小(KB)",
        "状态",
        "体积(mm³)",
        "体积(cm³)",
        "耗时(秒)",
        "错误信息",
    ]

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    err_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    total = len(records)
    num_parts = max(1, (total + chunk_size - 1) // chunk_size)

    for part_idx in range(num_parts):
        start_i = part_idx * chunk_size
        end_i = min(total, (part_idx + 1) * chunk_size)
        chunk = records[start_i:end_i]

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"STEP体积_{part_idx + 1}"
        ws.append(headers)

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for offset, rec in enumerate(chunk, start=1):
            global_idx = start_i + offset
            row_vals = [
                global_idx,
                rec.get("file_name", ""),
                rec.get("file_path", ""),
                rec.get("file_size_kb", 0.0),
                rec.get("status", ""),
                rec.get("volume_mm3", ""),
                rec.get("volume_cm3", ""),
                rec.get("elapsed_sec", 0.0),
                rec.get("error", ""),
            ]
            ws.append(row_vals)
            current_row = offset + 1
            status_cell = ws.cell(row=current_row, column=5)
            if rec.get("status") == "success":
                status_cell.fill = ok_fill
            else:
                status_cell.fill = err_fill

        col_widths = [8, 32, 65, 15, 12, 18, 18, 12, 45]
        for idx, width in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = width

        out_file = out_dir / f"step_volumes_part{part_idx + 1:02d}_of_{num_parts:02d}.xlsx"
        wb.save(out_file)
        created_files.append(out_file)

    return created_files


def export_csv_chunks(records: List[Dict[str, Any]], out_dir: Path, chunk_size: int) -> List[Path]:
    """当未安装 openpyxl 时自动降级导出 CSV"""
    created_files: List[Path] = []
    headers = [
        "index",
        "file_name",
        "file_path",
        "file_size_kb",
        "status",
        "volume_mm3",
        "volume_cm3",
        "elapsed_sec",
        "error",
    ]
    total = len(records)
    num_parts = max(1, (total + chunk_size - 1) // chunk_size)

    for part_idx in range(num_parts):
        start_i = part_idx * chunk_size
        end_i = min(total, (part_idx + 1) * chunk_size)
        chunk = records[start_i:end_i]

        out_file = out_dir / f"step_volumes_part{part_idx + 1:02d}_of_{num_parts:02d}.csv"
        with out_file.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            for offset, rec in enumerate(chunk, start=1):
                writer.writerow([
                    start_i + offset,
                    rec.get("file_name", ""),
                    rec.get("file_path", ""),
                    rec.get("file_size_kb", 0.0),
                    rec.get("status", ""),
                    rec.get("volume_mm3", ""),
                    rec.get("volume_cm3", ""),
                    rec.get("elapsed_sec", 0.0),
                    rec.get("error", ""),
                ])
        created_files.append(out_file)

    return created_files


def create_executor(workers: int) -> concurrent.futures.ProcessPoolExecutor:
    return concurrent.futures.ProcessPoolExecutor(max_workers=workers)


def main() -> None:
    args = parse_args()

    if args.root and args.output_dir:
        scan_root = _clean_user_path(args.root)
        out_dir_str = _clean_user_path(args.output_dir)
        out_dir = Path(out_dir_str)
        out_dir.mkdir(parents=True, exist_ok=True)
    elif args.no_interactive:
        scan_root = _clean_user_path(args.root or os.getcwd())
        out_dir = Path(_clean_user_path(args.output_dir or (Path.cwd() / "output")))
        out_dir.mkdir(parents=True, exist_ok=True)
    else:
        scan_root, out_dir_str = prompt_interactive_paths()
        out_dir = Path(out_dir_str)

    checkpoint_path = out_dir / "checkpoint_progress.jsonl"
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 正在扫描 STEP 文件: {scan_root} ...")
    all_files = discover_step_files(scan_root)
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 共发现 {len(all_files)} 个 STEP 文件。")

    done_map = load_checkpoint(checkpoint_path)
    pending_files = [fp for fp in all_files if fp not in done_map]
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 已完成 {len(done_map)} 个，待处理 {len(pending_files)} 个。")

    if pending_files:
        workers = max(1, args.workers)
        executor = create_executor(workers)
        processed_count = 0
        total_pending = len(pending_files)

        try:
            idx = 0
            while idx < total_pending:
                batch = pending_files[idx : idx + workers]
                future_to_fp = {executor.submit(calc_step_volume_worker, fp): fp for fp in batch}

                pool_broken = False
                for fut in concurrent.futures.as_completed(future_to_fp):
                    fp = future_to_fp[fut]
                    try:
                        rec = fut.result(timeout=args.timeout)
                    except concurrent.futures.TimeoutError:
                        rec = {
                            "file_path": fp,
                            "file_name": os.path.basename(fp),
                            "file_size_kb": round(os.path.getsize(fp) / 1024.0, 2) if os.path.exists(fp) else 0.0,
                            "status": "timeout",
                            "volume_mm3": None,
                            "volume_cm3": None,
                            "solid_count": 0,
                            "elapsed_sec": args.timeout,
                            "error": f"处理超时 (>{args.timeout}s)",
                        }
                    except concurrent.futures.BrokenProcessPool as e:
                        rec = {
                            "file_path": fp,
                            "file_name": os.path.basename(fp),
                            "file_size_kb": round(os.path.getsize(fp) / 1024.0, 2) if os.path.exists(fp) else 0.0,
                            "status": "crashed",
                            "volume_mm3": None,
                            "volume_cm3": None,
                            "solid_count": 0,
                            "elapsed_sec": 0.0,
                            "error": f"子进程崩溃 (BrokenProcessPool): {e}",
                        }
                        pool_broken = True
                    except Exception as e:
                        rec = {
                            "file_path": fp,
                            "file_name": os.path.basename(fp),
                            "file_size_kb": round(os.path.getsize(fp) / 1024.0, 2) if os.path.exists(fp) else 0.0,
                            "status": "failed",
                            "volume_mm3": None,
                            "volume_cm3": None,
                            "solid_count": 0,
                            "elapsed_sec": 0.0,
                            "error": f"{type(e).__name__}: {e}",
                        }

                    done_map[fp] = rec
                    append_checkpoint(checkpoint_path, rec)
                    processed_count += 1
                    if processed_count % 50 == 0 or processed_count == total_pending:
                        print(
                            f"[{datetime.now().strftime('%H:%M:%S')}] "
                            f"进度: {processed_count}/{total_pending} | 最新: {rec['file_name']} ({rec['status']})"
                        )

                if pool_broken:
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] 检测到子进程崩溃，正在重建进程池...")
                    executor.shutdown(wait=False, cancel_futures=True)
                    executor = create_executor(workers)

                idx += workers
        finally:
            executor.shutdown(wait=False, cancel_futures=True)

    ordered_records = [done_map[fp] for fp in all_files if fp in done_map]
    if HAS_OPENPYXL:
        exported = export_excel_chunks(ordered_records, out_dir, args.chunk_size)
    else:
        print("提示: 当前环境未安装 openpyxl，已自动导出为 CSV。")
        exported = export_csv_chunks(ordered_records, out_dir, args.chunk_size)

    success_cnt = sum(1 for r in ordered_records if r.get("status") == "success")
    fail_cnt = len(ordered_records) - success_cnt
    print("=" * 64)
    print(f"处理完成！总计: {len(ordered_records)} | 成功: {success_cnt} | 异常/超时: {fail_cnt}")
    print(f"断点文件: {checkpoint_path}")
    for fpath in exported:
        print(f"结果文件: {fpath}")
    print("=" * 64)


if __name__ == "__main__":
    # Windows 多进程打包（PyInstaller）必需
    try:
        import multiprocessing

        multiprocessing.freeze_support()
    except Exception:
        pass
    main()