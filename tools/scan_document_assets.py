#!/usr/bin/env python3
"""只读扫描项目文档资产目录，生成文档元数据清单。

用法示例：
  python3 tools/scan_document_assets.py \
    --source /path/to/contracts \
    --source /path/to/site-archives \
    --output document_assets/import_batches/scan-report.json

脚本只读取源文件并计算元数据/sha256，不复制、不删除、不修改源文件。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import mimetypes
import os
import re
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DOCUMENT_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp",
    ".xml", ".json", ".svg", ".txt", ".md", ".csv",
}

CATEGORY_RULES = [
    ("contract", re.compile(r"合同|协议|付款|发票|采购|销售|总集|分包")),
    ("meeting", re.compile(r"会议|纪要|周例会|专题会|会签")),
    ("change", re.compile(r"变更|签证|洽商|调整|补充")),
    ("plan", re.compile(r"计划|进度|排期|里程碑|交付")),
    ("warehouse", re.compile(r"入库|出库|到货|领料|送货|物料|设备清单")),
    ("construction", re.compile(r"安装|施工|验收|竣工|接线|开箱|照片|一档|一路一档|一点一档")),
    ("map_asset", re.compile(r"MAP|RSI|地图|信号机|map_xml|map_json|rsi_xml|rsi_json|SVG", re.IGNORECASE)),
    ("ops", re.compile(r"运维|工单|巡检|故障|整改|离线|异常")),
    ("management", re.compile(r"汇报|审批|说明|风险|管理")),
]

OBJECT_RULES = [
    ("site", re.compile(r"NodeID[:：_\- ]?(\d{3,8})|node[:：_\- ]?(\d{3,8})", re.IGNORECASE)),
    ("contract", re.compile(r"合同(?:编号)?[:：_\- ]?([A-Za-z0-9\-]{4,})")),
    ("map_asset", re.compile(r"(?:map|rsi)_?(\d{3,8})[_\-]", re.IGNORECASE)),
]


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        while True:
            chunk = file.read(chunk_size)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def guess_category(path: Path) -> str:
    text = str(path)
    for category, pattern in CATEGORY_RULES:
        if pattern.search(text):
            return category
    return "unclassified"


def guess_relation(path: Path) -> dict[str, Any] | None:
    text = str(path)
    for object_type, pattern in OBJECT_RULES:
        match = pattern.search(text)
        if match:
            object_id = next((group for group in match.groups() if group), match.group(0))
            return {
                "objectType": object_type,
                "objectId": object_id,
                "matchSource": "filename_or_path_rule",
                "confidence": 0.6,
            }
    return None


def storage_key_for(project_code: str, category: str, file_hash: str, path: Path) -> str:
    year = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y")
    safe_name = path.name.replace("/", "_")
    return f"raw/{project_code}/{category}/{year}/{file_hash[:2]}/{file_hash}/{safe_name}"


def scan_sources(sources: list[Path], project_code: str, max_files: int | None = None) -> dict[str, Any]:
    records: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []
    seen_hashes: dict[str, str] = {}

    for source in sources:
        if not source.exists():
            errors.append({"path": str(source), "error": "source_not_found"})
            continue
        paths = [source] if source.is_file() else source.rglob("*")
        for path in paths:
            if max_files and len(records) >= max_files:
                break
            if not path.is_file():
                continue
            ext = path.suffix.lower()
            if ext not in DOCUMENT_EXTENSIONS:
                continue
            try:
                stat = path.stat()
                file_hash = sha256_file(path)
                category = guess_category(path)
                relation = guess_relation(path)
                duplicate_of = seen_hashes.get(file_hash)
                if not duplicate_of:
                    seen_hashes[file_hash] = str(path)
                mime_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
                records.append({
                    "fileName": path.name,
                    "sourcePath": str(path),
                    "fileExt": ext.lstrip("."),
                    "mimeType": mime_type,
                    "fileSizeBytes": stat.st_size,
                    "sha256": file_hash,
                    "documentCategory": category,
                    "storageProvider": "local",
                    "storageKey": storage_key_for(project_code, category, file_hash, path),
                    "lastModifiedAt": datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat(),
                    "duplicateOf": duplicate_of or "",
                    "relation": relation,
                    "reviewStatus": "pending_review",
                    "parseStatus": "pending",
                    "qualityStatus": "warning" if duplicate_of or not relation else "normal",
                })
            except Exception as exc:  # noqa: BLE001
                errors.append({"path": str(path), "error": str(exc)})

    category_counts = Counter(record["documentCategory"] for record in records)
    ext_counts = Counter(record["fileExt"] for record in records)
    duplicates = [record for record in records if record["duplicateOf"]]
    unlinked = [record for record in records if not record["relation"]]
    total_size = sum(record["fileSizeBytes"] for record in records)

    return {
        "batchNo": f"SCAN-{datetime.now().strftime('%Y%m%d-%H%M%S')}",
        "projectCode": project_code,
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "sources": [str(source) for source in sources],
        "summary": {
            "fileCount": len(records),
            "totalSizeBytes": total_size,
            "duplicateFileCount": len(duplicates),
            "unlinkedFileCount": len(unlinked),
            "categoryCounts": dict(category_counts),
            "extensionCounts": dict(ext_counts),
            "errorCount": len(errors),
        },
        "records": records,
        "errors": errors,
    }


def write_csv(report: dict[str, Any], output: Path) -> None:
    fields = [
        "fileName", "sourcePath", "fileExt", "mimeType", "fileSizeBytes", "sha256",
        "documentCategory", "storageProvider", "storageKey", "lastModifiedAt",
        "duplicateOf", "reviewStatus", "parseStatus", "qualityStatus",
    ]
    with output.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        for record in report["records"]:
            row = {field: record.get(field, "") for field in fields}
            writer.writerow(row)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="只读扫描文档资产目录，生成元数据报告。")
    parser.add_argument("--source", action="append", required=True, help="待扫描源目录或文件，可重复传入。")
    parser.add_argument("--project-code", default="wuxi-cv2x", help="项目编码。")
    parser.add_argument("--output", default="document_assets/import_batches/latest-document-scan.json", help="JSON 报告输出路径。")
    parser.add_argument("--csv-output", default="", help="可选 CSV 清单输出路径。")
    parser.add_argument("--max-files", type=int, default=0, help="最多扫描文件数；0 表示不限制。")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    sources = [Path(source).expanduser() for source in args.source]
    report = scan_sources(sources, args.project_code, args.max_files or None)
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if args.csv_output:
        csv_output = Path(args.csv_output)
        csv_output.parent.mkdir(parents=True, exist_ok=True)
        write_csv(report, csv_output)
    print(json.dumps({"output": str(output), "summary": report["summary"]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
