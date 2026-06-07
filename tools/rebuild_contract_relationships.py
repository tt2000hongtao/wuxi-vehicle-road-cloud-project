#!/usr/bin/env python3
"""Refresh contract document index, then rebuild contract relationships.

This is the safe one-command entrypoint for the contract-management prototype:

1. Scan the configured contract source directories and regenerate
   ``prototype/data/document-assets.json``.
2. Rebuild ``prototype/data/contract-relationships.json`` from the refreshed
   document index and the contract-flow workbook.

The scanner is read-only against source contract folders. It only writes output
JSON/CSV files inside this workspace.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path


WORKSPACE = Path(__file__).resolve().parents[1]
PYTHON_BIN = os.environ.get("PYTHON_BIN") or sys.executable

DEFAULT_CONTRACT_SOURCES = [
    Path("/Users/tt2000/Documents/天安智联/0000无锡市车路云一体化项目/0.合同/盖章版合同 WORD 版"),
    Path("/Users/tt2000/Documents/天安智联/0000无锡市车路云一体化项目/0.合同/盖章版合同"),
]
DEFAULT_WORKBOOK = Path(
    "/Users/tt2000/Documents/天安智联/0000无锡市车路云一体化项目/0.合同/合同管理/无锡车路云一体化采购价格测算表202500920.xlsx"
)
DEFAULT_DOCUMENT_ASSETS = WORKSPACE / "prototype/data/document-assets.json"
DEFAULT_DOCUMENT_ASSETS_CSV = WORKSPACE / "document_assets/import_batches/latest-contract-document-assets.csv"
DEFAULT_RELATIONSHIPS = WORKSPACE / "prototype/data/contract-relationships.json"
DEFAULT_DERIVED_RELATIONSHIPS = WORKSPACE / "document_assets/derived/contract-relationships.json"


def run_command(args: list[str], cwd: Path) -> dict:
    completed = subprocess.run(args, cwd=str(cwd), check=True, text=True, capture_output=True)
    stdout = completed.stdout.strip()
    if not stdout:
        return {}
    try:
        return json.loads(stdout)
    except json.JSONDecodeError:
        return {"stdout": stdout}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="扫描合同目录并重建合同关系。")
    parser.add_argument(
        "--source",
        action="append",
        default=[],
        help="合同源目录或文件，可重复传入；不传则扫描默认 Word 合同目录和盖章合同目录。",
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK, help="合同流 Excel 文件。")
    parser.add_argument("--document-assets", type=Path, default=DEFAULT_DOCUMENT_ASSETS, help="刷新后的 document-assets.json 输出路径。")
    parser.add_argument("--document-assets-csv", type=Path, default=DEFAULT_DOCUMENT_ASSETS_CSV, help="刷新后的 CSV 清单输出路径。")
    parser.add_argument("--relationships", type=Path, default=DEFAULT_RELATIONSHIPS, help="合同关系 JSON 输出路径。")
    parser.add_argument("--derived-relationships", type=Path, default=DEFAULT_DERIVED_RELATIONSHIPS, help="派生合同关系 JSON 输出路径。")
    parser.add_argument("--project-code", default="wuxi-cv2x", help="文档资产项目编码。")
    parser.add_argument("--max-files", type=int, default=0, help="最多扫描文件数；0 表示不限制。")
    parser.add_argument("--skip-scan", action="store_true", help="跳过目录扫描，仅基于现有 document-assets.json 重建合同关系。")
    parser.add_argument("--exclude-file", type=Path, default=None, help="合同系统移除清单；匹配 contractId/sourcePath/sha256 的文件不进入重建结果。")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    sources = [Path(item).expanduser() for item in args.source] if args.source else DEFAULT_CONTRACT_SOURCES

    summary: dict[str, object] = {
        "sources": [str(source) for source in sources],
        "documentAssets": str(args.document_assets),
        "relationships": str(args.relationships),
        "derivedRelationships": str(args.derived_relationships),
    }

    if not args.skip_scan:
        scan_command = [
            PYTHON_BIN,
            str(WORKSPACE / "tools/scan_document_assets.py"),
            "--project-code",
            args.project_code,
            "--output",
            str(args.document_assets),
            "--csv-output",
            str(args.document_assets_csv),
        ]
        if args.max_files:
            scan_command.extend(["--max-files", str(args.max_files)])
        for source in sources:
            scan_command.extend(["--source", str(source)])
        scan_result = run_command(scan_command, WORKSPACE)
        summary["scan"] = scan_result.get("summary", scan_result)
    else:
        summary["scan"] = "skipped"

    analyze_command = [
        PYTHON_BIN,
        str(WORKSPACE / "tools/analyze_contract_relationships.py"),
        "--workbook",
        str(args.workbook),
        "--document-assets",
        str(args.document_assets),
        "--out",
        str(args.relationships),
        "--derived",
        str(args.derived_relationships),
    ]
    if args.exclude_file:
        analyze_command.extend(["--exclude-file", str(args.exclude_file)])
    analyze_result = run_command(analyze_command, WORKSPACE)
    summary["relationshipsSummary"] = analyze_result

    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
