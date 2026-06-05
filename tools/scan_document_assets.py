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

CONTRACT_SUBCATEGORY_RULES = [
    ("front_sales_contract", re.compile(r"前项天安销售|前项|前向|销售|1\.前向协议汇总")),
    ("back_procurement_contract", re.compile(r"后项天安采购|后向|采购盖章|分包协议汇总|0\.分包协议汇总")),
    ("payment_first_delivery", re.compile(r"第一笔到货款|第一次进度款|第一笔")),
    ("payment_second_delivery", re.compile(r"第二笔到货款|第二次进度款|第二笔")),
    ("payment_third_delivery", re.compile(r"第三笔到货款|第三次进度款|第四次请款|工程款|进度款|请款|支付申请")),
    ("workload_list", re.compile(r"移动分项工作量清单|工作量清单|工程量汇总|分项明细")),
    ("sales_procurement_analysis", re.compile(r"销采|收支台账|采购价格测算|颜色匹配|设备级|分项明细级")),
]

KNOWN_CONTRACT_PARTIES = [
    "上海帆一", "中通服网盈", "中城工联", "合创智行", "工业安装", "四维图新",
    "航天海特", "海康智联", "中电鸿", "公安部", "浙江海康", "车城智联",
    "无锡电信", "信通院", "金中天", "万集", "志合", "博世", "华通", "隆顺",
    "海特", "帆一", "浪潮", "中通服",
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
    if "合同管理" in text:
        return "contract"
    for category, pattern in CATEGORY_RULES:
        if pattern.search(text):
            return category
    return "unclassified"


def guess_contract_subcategory(path: Path) -> str:
    text = str(path)
    for subcategory, pattern in CONTRACT_SUBCATEGORY_RULES:
        if pattern.search(text):
            return subcategory
    return "contract_other"


def slug_for(value: str, max_length: int = 80) -> str:
    slug = re.sub(r"\.[^.]+$", "", value)
    slug = re.sub(r"[\s/\\:：*?\"<>|()（）【】\[\]，,]+", "-", slug)
    slug = re.sub(r"-+", "-", slug).strip("-")
    return slug[:max_length] or "document"


def item_no_from_name(name: str) -> str | None:
    normalized = name.strip().replace("——", "-").replace("--", "-")
    match = re.match(r"^(\d+(?:(?:\.|-)\d+){0,3})", normalized)
    return match.group(1) if match else None


def strip_item_no(name: str) -> str:
    normalized = name.strip().replace("——", "-").replace("--", "-")
    return re.sub(r"^\d+(?:(?:\.|-)\d+){0,3}\s*", "", normalized).strip()


def extract_contract_date(text: str) -> str:
    match = re.search(r"(20\d{2})[-年](\d{1,2})[-月](\d{1,2})日?", text)
    if not match:
        return ""
    year, month, day = match.groups()
    return f"{year}-{int(month):02d}-{int(day):02d}"


def extract_contract_amount(text: str) -> dict[str, Any]:
    match = re.search(r"(\d+(?:\.\d+)?)\s*(亿元|亿|万元|万|元)(每月)?", text)
    if not match:
        monthly = re.search(r"(\d+(?:\.\d+)?)\s*每月", text)
        if not monthly:
            return {"amountText": "", "amountCny": None, "amountCycle": ""}
        return {
            "amountText": monthly.group(0),
            "amountCny": float(monthly.group(1)),
            "amountCycle": "monthly",
        }
    value = float(match.group(1))
    unit = match.group(2)
    amount_cny = value
    if unit in {"亿元", "亿"}:
        amount_cny = value * 100000000
    elif unit in {"万元", "万"}:
        amount_cny = value * 10000
    return {
        "amountText": match.group(0),
        "amountCny": round(amount_cny, 2),
        "amountCycle": "monthly" if match.group(3) else "",
    }


def extract_contract_parties(stem: str) -> dict[str, str]:
    text = strip_item_no(stem)
    text = re.sub(r"20\d{2}[-年]\d{1,2}[-月]\d{1,2}日?", "", text)
    text = re.sub(r"\d+(?:\.\d+)?\s*(?:亿元|亿|万元|万|元)(?:每月)?", "", text)
    text = re.sub(r"合同附件\d*|补充协议|补充修改协议|作废|条款有误|用印|盖章|更新", "", text)
    text = text.strip(" -_　")
    pair = re.search(r"([\u4e00-\u9fa5A-Za-z][\u4e00-\u9fa5A-Za-z0-9]*)\s*-\s*([\u4e00-\u9fa5A-Za-z][\u4e00-\u9fa5A-Za-z0-9]*)", text)
    if pair:
        party_a, party_b = pair.group(1), pair.group(2)
        return {"partyA": party_a, "partyB": party_b, "counterparty": party_b if "天安" in party_a else party_a}
    tianan_buy = re.search(r"天安采([\u4e00-\u9fa5A-Za-z0-9]+)", text)
    if tianan_buy:
        return {"partyA": "天安", "partyB": tianan_buy.group(1), "counterparty": tianan_buy.group(1)}
    for party in KNOWN_CONTRACT_PARTIES:
        if party in text:
            if "天安" in text:
                return {"partyA": "天安", "partyB": party, "counterparty": party}
            return {"partyA": "", "partyB": party, "counterparty": party}
    words = re.findall(r"[\u4e00-\u9fa5A-Za-z0-9]{2,}", text)
    if "天安" in text and words:
        counterparty = next((word for word in words if "天安" not in word and "车路云" not in word), "")
        return {"partyA": "天安", "partyB": counterparty, "counterparty": counterparty}
    return {"partyA": "", "partyB": "", "counterparty": ""}


def extract_contract_fields(path: Path, subcategory: str) -> dict[str, Any]:
    stem = path.stem
    normalized_stem = strip_item_no(stem)
    amount = extract_contract_amount(normalized_stem)
    parties = extract_contract_parties(stem)
    flags = []
    if re.search(r"补充|补充协议|补充修改", stem):
        flags.append("supplement")
    if re.search(r"附件", stem):
        flags.append("attachment")
    if re.search(r"作废|条款有误|税率有变", stem):
        flags.append("voided")
    if re.search(r"发票", str(path.parent)):
        flags.append("invoice")
    if re.search(r"送货单|收货单", str(path.parent) + stem):
        flags.append("delivery_receipt")
    return {
        "contractDirection": "front_sales" if subcategory == "front_sales_contract" else "back_procurement" if subcategory == "back_procurement_contract" else "",
        "contractItemNo": item_no_from_name(path.name) or "",
        "contractName": stem,
        "signDate": extract_contract_date(stem),
        "amountText": amount["amountText"],
        "amountCny": amount["amountCny"],
        "amountCycle": amount["amountCycle"],
        "partyA": parties["partyA"],
        "partyB": parties["partyB"],
        "counterparty": parties["counterparty"],
        "flags": flags,
        "documentStatusHint": "voided" if "voided" in flags else "effective_candidate",
    }


def guess_contract_relation(path: Path, subcategory: str) -> dict[str, Any]:
    name = path.name
    item_no = item_no_from_name(name)
    stem_slug = slug_for(name)
    folder_slug = slug_for(path.parent.name)
    if subcategory == "front_sales_contract":
        object_type = "contract"
        relation_type = "front_sales_contract"
        object_id = item_no or stem_slug
    elif subcategory == "back_procurement_contract":
        object_type = "contract"
        relation_type = "back_procurement_contract"
        object_id = item_no or stem_slug
    elif subcategory.startswith("payment_"):
        object_type = "contract_flow_line"
        relation_type = "payment_support"
        object_id = item_no or folder_slug
    elif subcategory == "workload_list":
        object_type = "contract_item"
        relation_type = "workload_list"
        object_id = item_no or stem_slug
    elif subcategory == "sales_procurement_analysis":
        object_type = "contract_flow_line"
        relation_type = "sales_procurement_analysis"
        object_id = stem_slug
    else:
        object_type = "contract"
        relation_type = "contract_document"
        object_id = stem_slug
    return {
        "objectType": object_type,
        "objectId": object_id,
        "objectDisplayName": path.stem,
        "relationType": relation_type,
        "matchSource": "contract_path_rule",
        "confidence": 0.55 if subcategory == "contract_other" else 0.72,
    }


def guess_relation(path: Path, category: str, subcategory: str) -> dict[str, Any] | None:
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
    if category == "contract":
        return guess_contract_relation(path, subcategory)
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
                subcategory = guess_contract_subcategory(path) if category == "contract" else ""
                extracted_fields = extract_contract_fields(path, subcategory) if category == "contract" else {}
                relation = guess_relation(path, category, subcategory)
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
                    "documentSubcategory": subcategory,
                    "storageProvider": "local",
                    "storageKey": storage_key_for(project_code, category, file_hash, path),
                    "lastModifiedAt": datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat(),
                    "duplicateOf": duplicate_of or "",
                    "relation": relation,
                    "relationObjectType": relation.get("objectType", "") if relation else "",
                    "relationObjectId": relation.get("objectId", "") if relation else "",
                    "relationType": relation.get("relationType", "") if relation else "",
                    "relationConfidence": relation.get("confidence", "") if relation else "",
                    "extractedFields": extracted_fields,
                    "contractDirection": extracted_fields.get("contractDirection", ""),
                    "contractItemNo": extracted_fields.get("contractItemNo", ""),
                    "contractName": extracted_fields.get("contractName", ""),
                    "signDate": extracted_fields.get("signDate", ""),
                    "amountText": extracted_fields.get("amountText", ""),
                    "amountCny": extracted_fields.get("amountCny", ""),
                    "partyA": extracted_fields.get("partyA", ""),
                    "partyB": extracted_fields.get("partyB", ""),
                    "counterparty": extracted_fields.get("counterparty", ""),
                    "documentStatusHint": extracted_fields.get("documentStatusHint", ""),
                    "fieldFlags": ",".join(extracted_fields.get("flags", [])),
                    "reviewStatus": "pending_review",
                    "parseStatus": "pending",
                    "qualityStatus": "warning" if duplicate_of or not relation else "normal",
                })
            except Exception as exc:  # noqa: BLE001
                errors.append({"path": str(path), "error": str(exc)})

    category_counts = Counter(record["documentCategory"] for record in records)
    subcategory_counts = Counter(record["documentSubcategory"] for record in records if record.get("documentSubcategory"))
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
            "subcategoryCounts": dict(subcategory_counts),
            "extensionCounts": dict(ext_counts),
            "errorCount": len(errors),
        },
        "records": records,
        "errors": errors,
    }


def write_csv(report: dict[str, Any], output: Path) -> None:
    fields = [
        "fileName", "sourcePath", "fileExt", "mimeType", "fileSizeBytes", "sha256",
        "documentCategory", "documentSubcategory", "storageProvider", "storageKey", "lastModifiedAt",
        "relationObjectType", "relationObjectId", "relationType", "relationConfidence",
        "contractDirection", "contractItemNo", "contractName", "signDate", "amountText",
        "amountCny", "partyA", "partyB", "counterparty", "documentStatusHint", "fieldFlags",
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
