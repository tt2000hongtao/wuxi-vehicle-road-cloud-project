#!/usr/bin/env python3
"""Build front/back contract relationship candidates from workbook + Word contracts.

This script corrects an important modeling point:

- Excel detail sheets are NOT contracts. They are equipment/software/service
  list pools.
- Contract boundaries must come from signed front-sales/back-procurement
  contract documents.
- The tail section of workbook sheet `合同流` is used as the macro cash-flow
  skeleton.
- Front contracts and back contracts form many-to-many relationships through
  Tianan, and must be confirmed at item/amount/tax evidence level later.

The output is intentionally candidate-only. It must not be used for payment or
gross-margin calculation without human confirmation.
"""

from __future__ import annotations

import argparse
import json
import math
import re
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook


DEFAULT_WORKBOOK = Path(
    "/Users/tt2000/Documents/天安智联/0000无锡市车路云一体化项目/0.合同/合同管理/无锡车路云一体化采购价格测算表202500920.xlsx"
)
DEFAULT_DOCUMENT_ASSETS = Path("prototype/data/document-assets.json")
DEFAULT_OUT = Path("prototype/data/contract-relationships.json")
DEFAULT_DERIVED = Path("document_assets/derived/contract-relationships.json")

PARTY_ALIASES = {
    "江苏天安智联科技股份有限公司": "天安",
    "天安智联": "天安",
    "天安": "天安",
    "中国移动通信集团江苏有限公司无锡分公司": "移动",
    "中国移动": "移动",
    "江苏移动": "移动",
    "移动": "移动",
    "浪潮": "浪潮",
    "航天大为": "大为",
    "大为": "大为",
    "华通": "华通",
    "隆顺": "隆顺",
    "中信科智联": "中信科",
    "中信科": "中信科",
    "联通华盛": "联通华盛",
    "联通": "联通",
    "吉利": "吉利",
    "大华": "大华",
    "海康智联": "海康智联",
    "海康": "海康智联",
    "工业安装": "工业安装",
    "中邮建": "中邮建",
    "上研": "上研",
    "上电科": "上电科",
    "中通服网盈": "中通服网盈",
    "尚行": "尚行",
    "帆一": "尚行",
    "上汽帆一": "尚行",
    "合创": "合创",
    "万集": "万集",
    "车城": "车城",
    "四维图新": "四维图新",
    "信通院": "信通院",
    "交科所": "交科所",
    "金中天": "金中天",
    "金晓": "金晓",
    "中科创达": "中科创达",
    "华三": "华三",
    "华为": "华为",
    "回车": "回车",
}

DEVICE_KEYWORDS = [
    "RSU", "信号机", "信号灯", "服务器", "GPU", "雷达", "激光雷达", "毫米波",
    "摄像", "相机", "边缘云", "云控", "高精地图", "地图", "数据质量", "C-V2X",
    "网关", "密码机", "公交", "车载", "OBU", "光纤", "物联网卡", "机柜",
    "施工", "安装", "运维", "软件", "平台", "算法", "三维实景",
]

LEGACY_MACRO_FLOW_KEY = "移动-" + "吉利" + "-天安（大华+" + "希迪" + "+RSU)"
MACRO_FLOW_NAME_OVERRIDES = {
    LEGACY_MACRO_FLOW_KEY: "移动-四维图新-天安（大华+海康智联+RSU)",
}

PAYMENT_STAGE_TEMPLATES = [
    {
        "stageCode": "advance",
        "stageName": "预付款",
        "stageOrder": 10,
        "triggerCondition": "合同签署、预付款条款满足",
        "ownerAuditRequirement": "通常不要求设备审计，但需绑定后续设备级资金池",
        "upstreamRequirement": "背靠背条款下需上游预付款到账",
        "releaseRule": "按后向合同预付款条款释放；不得越过候选关系确认闸门",
        "partialReleaseAllowed": True,
    },
    {
        "stageCode": "delivery",
        "stageName": "到货款",
        "stageOrder": 20,
        "triggerCondition": "到货、入库、发票或送货签收证据齐备",
        "ownerAuditRequirement": "可按到货批次预审，最终受业主设备级审计确认金额约束",
        "upstreamRequirement": "对应设备到货阶段回款或设备级资金池可用",
        "releaseRule": "按到货数量和后向单价折算部分释放",
        "partialReleaseAllowed": True,
    },
    {
        "stageCode": "installation",
        "stageName": "安装款",
        "stageOrder": 30,
        "triggerCondition": "安装签证、安装数量、NodeID/点位绑定完成",
        "ownerAuditRequirement": "按安装点位进入业主审计或预审",
        "upstreamRequirement": "对应设备安装阶段回款或资金池可用",
        "releaseRule": "按实际安装数量折算释放，未绑定点位不得释放",
        "partialReleaseAllowed": True,
    },
    {
        "stageCode": "acceptance",
        "stageName": "验收款",
        "stageOrder": 40,
        "triggerCondition": "验收单、验收数量和问题闭环完成",
        "ownerAuditRequirement": "验收资料进入业主设备级审计",
        "upstreamRequirement": "对应验收阶段回款或资金池可用",
        "releaseRule": "按验收通过数量和审计前置规则释放",
        "partialReleaseAllowed": True,
    },
    {
        "stageCode": "owner_audit",
        "stageName": "审计款",
        "stageOrder": 50,
        "triggerCondition": "业主审计确认数量和金额",
        "ownerAuditRequirement": "必须取得正式设备级审计结果",
        "upstreamRequirement": "审计款到账或明确可归集到设备级资金池",
        "releaseRule": "以审计确认金额为基数，扣除核减后释放",
        "partialReleaseAllowed": True,
    },
    {
        "stageCode": "warranty",
        "stageName": "质保金",
        "stageOrder": 60,
        "triggerCondition": "质保期届满、无遗留问题或问题已闭环",
        "ownerAuditRequirement": "受最终审计确认金额和质保扣款约束",
        "upstreamRequirement": "质保/尾款到账或可用",
        "releaseRule": "扣除质量问题、审计核减和已付款后释放",
        "partialReleaseAllowed": True,
    },
]

DEVICE_CASHFLOW_SCHEMA = {
    "minimumUnit": "front_contract_item -> back_contract_item -> item_pool_row -> NodeID/site -> owner_audit -> payment_stage",
    "mustHaveFields": [
        "macroFlowId",
        "frontContractId",
        "backContractId",
        "itemName",
        "model",
        "quantity",
        "frontTaxRate",
        "backTaxRate",
        "frontAmountTaxIncluded",
        "backAmountTaxIncluded",
        "nodeId",
        "ownerAuditStatus",
        "ownerAuditConfirmedAmount",
        "upstreamReceiptReleasedAmount",
        "downstreamPaymentReleasedAmount",
        "overpaymentRiskStatus",
    ],
    "gateRule": "front/back candidates are financialUseAllowed=false until confirmed; confirmed links still require device-item split and owner-audit evidence before payment release.",
}


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return re.sub(r"\s+", " ", text)


def stable_id(prefix: str, *parts: Any) -> str:
    import hashlib

    raw = "|".join(clean(p) for p in parts)
    return f"{prefix}_{hashlib.sha1(raw.encode('utf-8')).hexdigest()[:16]}"


def money(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if math.isnan(value):
            return None
        return float(value)
    s = clean(value).replace(",", "").replace("￥", "")
    try:
        return float(s)
    except Exception:
        return None


def normalize_party(text: Any) -> str:
    s = clean(text)
    if not s:
        return ""
    normalized_token = s.strip("()（） ")
    if normalized_token.upper() in {"RSU", "OBU", "GPU", "C-V2X"}:
        return ""
    for key, val in PARTY_ALIASES.items():
        if key in s:
            return val
    # Ignore obvious notes/numbers.
    if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", s) or len(s) > 40:
        return ""
    return s


def parties_in_text(text: str) -> list[str]:
    found = []
    for key, val in PARTY_ALIASES.items():
        if key in text and val not in found:
            found.append(val)
    return found


def split_chain(text: Any) -> list[str]:
    s = clean(text)
    if not s:
        return []
    s = re.sub(r"[【】\\[\\]（）()]", "-", s)
    parts = re.split(r"[-—→>+/、,，]+", s)
    out = []
    for part in parts:
        p = normalize_party(part)
        if p and p not in out:
            out.append(p)
    return out


def normalize_macro_flow_name(name: str) -> str:
    return MACRO_FLOW_NAME_OVERRIDES.get(clean(name), clean(name))


def extract_docx_text(path: Path) -> str:
    doc = Document(str(path))
    parts: list[str] = []
    parts.extend(p.text for p in doc.paragraphs if p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if any(cells):
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def extract_amounts(text: str, limit: int = 20) -> list[float]:
    amounts = []
    for m in re.finditer(r"(?<![A-Za-z0-9])(\d{1,3}(?:,\d{3})+|\d+)(?:\.(\d+))?\s*(亿元|亿|万元|万|元)?", text):
        raw = m.group(0)
        n = float((m.group(1) + ("." + m.group(2) if m.group(2) else "")).replace(",", ""))
        unit = m.group(3) or ""
        if unit in ("亿元", "亿"):
            n *= 100000000
        elif unit in ("万元", "万"):
            n *= 10000
        elif not unit and n < 10000:
            continue
        if 1000 <= n <= 500000000 and n not in amounts:
            amounts.append(round(n, 2))
        if len(amounts) >= limit:
            break
    return amounts


def extract_tax_rates(text: str) -> list[str]:
    tax_context = "\n".join(line for line in text.splitlines() if any(k in line for k in ["税率", "税点", "含税", "增值税", "发票"]))
    out = []
    for val in re.findall(r"(\d{1,2}(?:\.\d+)?)\s*%", tax_context):
        n = float(val)
        if 0 < n <= 20:
            s = f"{val}%"
            if s not in out:
                out.append(s)
    return out


def keyword_hits(text: str) -> list[str]:
    return [kw for kw in DEVICE_KEYWORDS if kw.lower() in text.lower()]


@dataclass
class MacroFlow:
    id: str
    rowNumber: int
    packageName: str
    packageAmount: float | None
    chain: list[str]
    tiananAmount: float | None
    downstreamPartiesAfterTianan: list[str]
    amounts: list[float]
    notes: str
    flowType: str


def extract_macro_flows(workbook: Path) -> list[MacroFlow]:
    wb = load_workbook(workbook, read_only=True, data_only=True)
    ws = wb["合同流"]
    flows: list[MacroFlow] = []
    current: dict[str, Any] | None = None
    for rn in range(138, ws.max_row + 1):
        package_name = clean(ws.cell(rn, 5).value)
        package_amount = money(ws.cell(rn, 6).value)
        row_values = [ws.cell(rn, c).value for c in range(1, 19)]
        row_text = " ".join(clean(v) for v in row_values if v is not None)
        if package_name and package_name != "天安合计":
            if current:
                flows.append(build_macro_flow(current))
            current = {
                "rowNumber": rn,
                "packageName": package_name,
                "packageAmount": package_amount,
                "rows": [row_values],
                "notes": [],
            }
        elif current:
            current["rows"].append(row_values)
        if current and row_text:
            current["notes"].append(row_text)
    if current:
        flows.append(build_macro_flow(current))
    return flows


def build_macro_flow(raw: dict[str, Any]) -> MacroFlow:
    parties = []
    amounts = []
    notes = " ".join(raw.get("notes", []))
    package_name = normalize_macro_flow_name(raw["packageName"])
    notes = notes.replace(raw["packageName"], package_name)
    for row in raw["rows"]:
        for c in [4, 6, 8, 10, 12, 14, 16]:
            if c < len(row):
                p = normalize_party(row[c])
                if p and p not in parties:
                    parties.append(p)
        for c in [5, 7, 9, 11, 13, 15, 17]:
            if c < len(row):
                n = money(row[c])
                if n is not None:
                    amounts.append(round(n, 2))
    chain = split_chain(package_name)
    for p in parties:
        if p not in chain:
            chain.append(p)
    after_tianan = []
    if "天安" in chain:
        after_tianan = chain[chain.index("天安") + 1 :]
    flow_type = "tianan_direct" if chain and chain[0] == "天安" else "mobile_to_tianan" if "移动" in chain and "天安" in chain else "other"
    tianan_amount = raw["packageAmount"] if "天安" in chain else None
    return MacroFlow(
        id=stable_id("macro_flow", raw["rowNumber"], raw["packageName"], raw["packageAmount"]),
        rowNumber=raw["rowNumber"],
        packageName=package_name,
        packageAmount=raw["packageAmount"],
        chain=chain,
        tiananAmount=tianan_amount,
        downstreamPartiesAfterTianan=after_tianan,
        amounts=amounts,
        notes=notes[:1200],
        flowType=flow_type,
    )


@dataclass
class ContractDoc:
    id: str
    fileName: str
    sourcePath: str
    direction: str
    counterparty: str
    parties: list[str]
    amounts: list[float]
    taxRates: list[str]
    keywords: list[str]
    textLength: int
    backToBackLikely: bool
    evidenceLines: list[str]


def extract_contract_docs(document_assets: Path) -> list[ContractDoc]:
    data = json.loads(document_assets.read_text(encoding="utf-8"))
    docs = []
    for record in data.get("records", []):
        path = Path(record["sourcePath"])
        ext = path.suffix.lower()
        text = ""
        if ext == ".docx":
            try:
                text = extract_docx_text(path)
            except Exception:
                text = ""
        elif ext == ".doc":
            # Older .doc can be parsed by the existing script; keep filename-level
            # evidence here if conversion is unavailable.
            text = ""
        fields = record.get("extractedFields", {})
        filename = record.get("fileName", path.name)
        base_text = f"{filename}\n{fields.get('contractName','')}\n{fields.get('partyA','')}\n{fields.get('partyB','')}\n{fields.get('counterparty','')}\n{text}"
        parties = parties_in_text(base_text)
        for key in [fields.get("partyA"), fields.get("partyB"), fields.get("counterparty")]:
            p = normalize_party(key)
            if p and p not in parties:
                parties.append(p)
        amount_candidates = []
        if fields.get("amountCny"):
            amount_candidates.append(float(fields["amountCny"]))
        for amount in extract_amounts(text):
            if amount not in amount_candidates:
                amount_candidates.append(amount)
        taxes = extract_tax_rates(text)
        kws = keyword_hits(base_text)
        lines = []
        for line in text.splitlines():
            compact = clean(line)
            if compact and (any(k in compact for k in kws[:8]) or any(t in compact for t in taxes)):
                lines.append(compact[:260])
            if len(lines) >= 8:
                break
        docs.append(
            ContractDoc(
                id=stable_id("contract_doc", filename, record.get("sourcePath")),
                fileName=filename,
                sourcePath=record.get("sourcePath", ""),
                direction=fields.get("contractDirection") or record.get("documentSubcategory", ""),
                counterparty=normalize_party(fields.get("counterparty", "")),
                parties=parties,
                amounts=amount_candidates[:20],
                taxRates=taxes,
                keywords=kws,
                textLength=len(text),
                backToBackLikely=any(k in text for k in ["甲方客户", "业主", "同等比例", "背靠背", "回款"]),
                evidenceLines=lines,
            )
        )
    return docs


def amount_close(a: float, b: float) -> bool:
    if not a or not b:
        return False
    tolerance = max(5000, min(a, b) * 0.03)
    return abs(a - b) <= tolerance


def match_contract_to_flows(doc: ContractDoc, flows: list[MacroFlow]) -> list[dict[str, Any]]:
    matches = []
    doc_parties = set(doc.parties)
    doc_keywords = set(doc.keywords)
    for flow in flows:
        score = 0
        reasons = []
        overlap = doc_parties & set(flow.chain + flow.downstreamPartiesAfterTianan)
        if overlap:
            score += len(overlap) * 3
            reasons.append(f"主体重合:{','.join(sorted(overlap))}")
        if doc.direction == "front_sales" and "天安" in doc_parties and flow.flowType in ("mobile_to_tianan", "tianan_direct"):
            score += 2
            reasons.append("前向合同涉及天安")
        if doc.direction == "back_procurement" and "天安" in doc_parties and flow.downstreamPartiesAfterTianan:
            score += 2
            reasons.append("后向合同涉及天安下游")
        amount_hits = []
        for a in doc.amounts[:10]:
            for b in [flow.packageAmount or 0, *(flow.amounts[:20])]:
                if amount_close(a, b):
                    amount_hits.append((a, b))
                    break
        if amount_hits:
            score += min(6, len(amount_hits) * 2)
            reasons.append("金额接近:" + ",".join(f"{int(a)}≈{int(b)}" for a, b in amount_hits[:3]))
        keyword_overlap = [kw for kw in doc_keywords if kw in flow.packageName or kw in flow.notes]
        if keyword_overlap:
            score += min(4, len(keyword_overlap))
            reasons.append("设备/服务关键词:" + ",".join(keyword_overlap[:5]))
        if score >= 3:
            matches.append(
                {
                    "macroFlowId": flow.id,
                    "macroFlowPackage": flow.packageName,
                    "macroFlowRow": flow.rowNumber,
                    "score": score,
                    "reasons": reasons,
                    "confidence": "high" if score >= 10 else "medium" if score >= 6 else "low",
                }
            )
    return sorted(matches, key=lambda x: (-x["score"], x["macroFlowRow"]))[:8]


def build_front_back_candidates(docs: list[ContractDoc], flows: list[MacroFlow]) -> list[dict[str, Any]]:
    front_docs = [d for d in docs if d.direction == "front_sales"]
    back_docs = [d for d in docs if d.direction == "back_procurement"]
    flow_matches = {d.id: match_contract_to_flows(d, flows) for d in docs}
    candidates = []
    for front in front_docs:
        front_flow_ids = {m["macroFlowId"] for m in flow_matches.get(front.id, [])}
        for back in back_docs:
            back_flow_ids = {m["macroFlowId"] for m in flow_matches.get(back.id, [])}
            shared_flow = sorted(front_flow_ids & back_flow_ids)
            party_overlap = sorted((set(front.parties) & set(back.parties)) - {"天安"})
            keyword_overlap = sorted(set(front.keywords) & set(back.keywords))
            tax_overlap = sorted(set(front.taxRates) & set(back.taxRates))
            score = 0
            reasons = []
            if shared_flow:
                score += 7
                reasons.append("命中同一宏观合同流")
            if "天安" in front.parties and "天安" in back.parties:
                score += 3
                reasons.append("均通过天安")
            if party_overlap:
                score += min(4, len(party_overlap) * 2)
                reasons.append("共同相对方/下游主体:" + ",".join(party_overlap[:4]))
            if keyword_overlap:
                score += min(5, len(keyword_overlap))
                reasons.append("设备/服务关键词重合:" + ",".join(keyword_overlap[:6]))
            if tax_overlap:
                score += 1
                reasons.append("税率重合:" + ",".join(tax_overlap))
            # Avoid flooding with weak pairs.
            if score < 8:
                continue
            candidates.append(
                {
                    "id": stable_id("front_back_candidate", front.id, back.id, ",".join(shared_flow)),
                    "frontContractId": front.id,
                    "frontContractName": front.fileName,
                    "backContractId": back.id,
                    "backContractName": back.fileName,
                    "sharedMacroFlowIds": shared_flow,
                    "partyOverlap": party_overlap,
                    "keywordOverlap": keyword_overlap[:12],
                    "taxRateOverlap": tax_overlap,
                    "score": score,
                    "confidence": "high" if score >= 14 else "medium",
                    "status": "candidate",
                    "financialUseAllowed": False,
                    "reasons": reasons,
                    "importantNote": "这是前后向合同多对多候选关系，不是设备级最终确认。必须继续比对合同附件清单、金额、税率、数量和履约资料。",
                }
            )
    return sorted(candidates, key=lambda x: (-x["score"], x["frontContractName"], x["backContractName"]))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--document-assets", type=Path, default=DEFAULT_DOCUMENT_ASSETS)
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT)
    parser.add_argument("--derived", type=Path, default=DEFAULT_DERIVED)
    args = parser.parse_args()

    macro_flows = extract_macro_flows(args.workbook)
    docs = extract_contract_docs(args.document_assets)
    doc_flow_matches = [
        {"contractId": d.id, "contractName": d.fileName, "direction": d.direction, "matches": match_contract_to_flows(d, macro_flows)}
        for d in docs
    ]
    candidates = build_front_back_candidates(docs, macro_flows)
    payload = {
        "schemaVersion": 1,
        "source": {
            "workbook": str(args.workbook),
            "documentAssets": str(args.document_assets),
        },
        "summary": {
            "macroFlowCount": len(macro_flows),
            "contractDocumentCount": len(docs),
            "frontContractCount": sum(1 for d in docs if d.direction == "front_sales"),
            "backContractCount": sum(1 for d in docs if d.direction == "back_procurement"),
            "frontBackCandidateCount": len(candidates),
            "taxRateCounts": dict(Counter(t for d in docs for t in d.taxRates)),
            "note": "Excel detail sheets are item pools, not contract boundaries. Front/back contract relationships are many-to-many through Tianan.",
        },
        "macroFlows": [asdict(f) for f in macro_flows],
        "contracts": [asdict(d) for d in docs],
        "contractToMacroFlowMatches": doc_flow_matches,
        "frontBackRelationshipCandidates": candidates,
        "deviceCashflowSchema": DEVICE_CASHFLOW_SCHEMA,
        "ownerAuditTrackingModel": {
            "minimumAuditUnit": "cashflow_unit_id + node_id/site_id + item_name + audit_stage",
            "statuses": ["not_submitted", "submitted", "under_review", "confirmed", "partially_confirmed", "deducted", "evidence_required"],
            "rule": "Upstream receipt and downstream payment release must use owner audit confirmed quantity/amount, not only contract-level amount.",
        },
        "paymentStageTemplates": PAYMENT_STAGE_TEMPLATES,
    }
    for path in [args.out, args.derived]:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(payload["summary"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
