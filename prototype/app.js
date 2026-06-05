const state = {
  panel: "overview",
  view: "table",
  query: "",
  district: "全部区域",
  perception: "全部感知类型",
  adaptive: "全部自适应",
  variableLane: "全部可变车道",
  selectedSite: null,
  importType: null,
  importTask: null,
  importHistory: [],
  opsDate: "2026-05-18",
  opsCalendarMode: "month",
  opsCalendarCursor: "2026-05",
  visualTheme: "macos",
  mapAssetColumnFilters: {
    siteType: [],
    district: [],
    completeness: [],
    mapXml: [],
    mapJson: [],
    signalExcel: [],
    svg: [],
  },
  siteColumnFilters: {
    district: [],
    type: [],
    vendor: [],
    perception: [],
    adaptive: [],
    variableLane: [],
  },
  activeMapAssetFilter: null,
  activeSiteFilter: null,
  selectedMapAssetId: null,
  documentAssets: {
    loading: true,
    error: "",
    reportPath: "",
    batchNo: "",
    generatedAt: "",
    summary: null,
    records: [],
  },
  contractReview: {
    decisions: {},
    filters: {
      query: "",
      decision: "all",
      confidence: "all",
      taxRate: "all",
    },
    selectedCandidateId: "",
  },
  contractRelationships: {
    loading: true,
    error: "",
    summary: null,
    macroFlows: [],
    contracts: [],
    contractToMacroFlowMatches: [],
    frontBackRelationshipCandidates: [],
    deviceCashflowSchema: null,
    ownerAuditTrackingModel: null,
    paymentStageTemplates: [],
    selectedMacroFlowId: "",
  },
};

const data = window.PROTOTYPE_DATA;
const sites = data.sites;
const sitesByNodeId = new Map(sites.map((site) => [String(site.nodeId || "").trim(), site]));
const roadsideStatusState = {
  currentDate: window.ROADSIDE_STATUS_DATA?.importDate || "2026-05-18",
  currentRows: normalizeRoadsideStatusRows(window.ROADSIDE_STATUS_DATA?.rows || []),
  archives: [],
};
const mapAssetState = {
  loading: true,
  error: "",
  root: "",
  generatedAt: "",
  summary: null,
  intersections: [],
  previewDrag: null,
};
const ROADSIDE_STATUS_STORAGE_KEY = "wuxi-roadside-status-state-v1";
const VISUAL_THEME_STORAGE_KEY = "wuxi-project-visual-theme-v1";
const CONTRACT_REVIEW_STORAGE_KEY = "wuxi-contract-review-decisions-v1";
const visualThemes = ["macos", "command", "trajectory"];
const persistenceState = {
  backendAvailable: false,
  lastSavedAt: null,
};
const unmatchedItems = buildUnmatchedItems();
const documentCategoryMeta = {
  contract: ["合同类", "总集、销售、采购、补充协议、付款申请", "contract"],
  meeting: ["会议类", "会议纪要、周例会、专题会、纪要确认单", "meeting"],
  change: ["变更类", "设计、工程、需求、价格、签证变更", "change"],
  plan: ["计划类", "项目计划、施工计划、进度计划、交付计划", "plan"],
  warehouse: ["物资类", "出入库单、到货单、领料单、设备清单", "warehouse"],
  construction: ["施工类", "安装签证、施工记录、照片、接线表、验收单", "construction"],
  map_asset: ["地图类", "MAP、RSI、信号机配置、SVG、质检报告", "map"],
  ops: ["运维类", "工单、巡检、故障、整改闭环资料", "ops"],
  management: ["管理类", "汇报材料、审批单、内部说明、风险清单", "management"],
  unclassified: ["未分类", "需要人工补充分类型和业务对象关联", "unclassified"],
};
const importErrors = [
  { sheet: "点位管理表", row: 148, level: "warning", code: "DISTRICT_MERGED", message: "经开区已归并为新吴区，原值进入 original_district 审计字段。" },
  { sheet: "点位管理表", row: 233, level: "error", code: "NODE_ID_EMPTY", message: "NodeID 为空，进入待补全治理，不计入正式点位数。" },
  { sheet: "设备管理表/Sheet1", row: 781, level: "warning", code: "DEVICE_UNMATCHED", message: "NodeID 未匹配到点位，进入未匹配池。" },
  { sheet: "设备管理表/未匹配", row: 46, level: "warning", code: "ASSET_TYPE_PENDING", message: "需人工标记为点位、仓库、机房、车端或服务项。" },
];
const amapInstances = new WeakMap();
const importTemplates = {
  sites: {
    title: "导入点位管理表",
    purpose: "用于导入点位管理表，生成或更新标准点位库。",
    template: "/Users/tt2000/Documents/天安智联/AI/项目管理工具包/Project Data/点位管理表-无锡车路云.xlsx",
    result: "识别 NodeID、点位名称、行政区域、点位类型、信号机厂商、感知点位类型、自适应点位、可变车道与 GCJ-02 展示坐标。",
  },
  devices: {
    title: "导入设备管理表",
    purpose: "用于导入设备安装位置表，形成设备安装台账并按 NodeID 关联到点位。该表不是路侧设备运行状态表。",
    template: "/Users/tt2000/Documents/天安智联/AI/项目管理工具包/Project Data/设备安装位置表.xlsx",
    result: "生成设备安装台账；未匹配记录进入 2447 条未匹配池，支持治理到点位、仓库、机房或服务项。",
  },
  roadsideStatus: {
    title: "导入路侧设备运行状态表",
    purpose: "用于导入每日路侧设备运行状态快照，来源为运维状态数据文件，不读取设备安装位置表。当日导入后，原当前状态表进入历史档案库，新文件成为当前路侧设备状态表。",
    template: "/Users/tt2000/Documents/天安智联/0000无锡市车路云一体化项目/11.运维/设备数据0518(1).xlsx",
    result: "生成运维统计、路侧设备状态详表、异常日历、离线设备点位地图和当日异常导出清单。",
  },
};
const importFieldMaps = {
  sites: [
    ["NodeID", "node_id", "唯一业务主键"],
    ["点位名称", "site_name", "由原表点位名称字段识别"],
    ["行政区域", "district", "锡山/新吴/惠山/梁溪/滨湖/经开"],
    ["点位类型", "site_type", "十/T/路段/匝道汇入/匝道汇出/V2P/环岛/Y"],
    ["信号机厂商", "signal_vendor", "标准化为厂商字典"],
    ["感知点位类型", "perception_site_type", "信号机厂商后展示"],
    ["是否自适应点位", "is_adaptive_site", "是/否标准化"],
    ["是否可变车道路口", "is_variable_lane_site", "由可变车道数量/方向推导"],
    ["GCJ-02 坐标", "longitude_gcj02 / latitude_gcj02", "业务页面和地图展示坐标"],
  ],
  devices: [
    ["NodeID", "node_id", "匹配点位"],
    ["点位名称", "site_name", "辅助匹配"],
    ["安装位置", "install_position", "设备落位"],
    ["供应商单位", "supplier_name", "后项合同板块"],
    ["物料编码", "material_code", "设备识别"],
    ["物料名称", "material_name", "设备类型归类"],
    ["后项合同名称", "contract_name", "合同穿透"],
    ["品牌型号", "brand_model", "规格型号"],
    ["送货/入库/领料/安装数量", "quantity_fields", "履约进度"],
  ],
  roadsideStatus: [
    ["设备编号", "device_id", "路侧设备唯一编号"],
    ["产品名称", "product_name", "设备产品名称"],
    ["设备类型名称", "device_type_name", "按设备类型统计"],
    ["设备位置", "device_position", "点位或道路位置"],
    ["经度/纬度", "longitude_gcj02 / latitude_gcj02", "离线设备地图落点"],
    ["路口ID", "intersection_id", "点位关联辅助键"],
    ["区域", "district", "行政区域"],
    ["状态", "status", "在线/离线/异常"],
    ["启用状态", "enabled_status", "仅启用设备纳入统计"],
  ],
};
const districtClassMap = {
  锡山: "district-xishan",
  新吴: "district-xinwu",
  惠山: "district-huishan",
  梁溪: "district-liangxi",
  滨湖: "district-binhu",
  经开: "district-xinwu",
};

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));
const pageHeaders = {
  overview: ["项目总览", "从总览页直接导入点位库和设备安装位置表"],
  sites: ["点位管理", "表格、卡片和地图三种视图共用筛选条件"],
  devices: ["设备管理", "按设备类型逐个展示，支持点位、安装位置、供应商、合同和履约状态穿透"],
  mapAssets: ["地图资产", "按路口聚合 MAP、RSI、信号机 Excel 和 SVG 预览，支持完整性校核与文件穿透"],
  imports: ["导入中心", "点位管理表、设备管理表、路侧设备状态表、合同和资料导入的预览、校验、确认与报告"],
  coordinates: ["坐标异常", "业务页面以 GCJ-02 为准；CGCS2000 仅作为原始数据审计字段保存"],
  contracts: ["合同管理", "基于 Excel 宏观合同流和 Word 前后向合同文本，重构经天安形成的多对多资金关系"],
  warehouse: ["出入库管理", "一期原型展示送货、入库、领料和安装数量差异"],
  documents: ["文档资产中心", "元数据、版本、业务对象关联、解析质检与云迁移兼容存储"],
  ops: ["运维管理", "每日路侧设备运行状态快照、异常统计、离线设备地图和历史日历"],
};

function buildUnmatchedItems() {
  const suppliers = ["无锡市工业设备安装有限公司", "江苏移动", "中信科", "车联网集团", "万集科技", "华通"];
  const materials = ["智能抱杆箱", "RSU 路侧单元", "边缘计算节点", "施工服务", "机房交换机", "软件开发服务", "视频感知设备"];
  const contracts = ["感知设备后向合同", "网络通信后向合同", "机房资源合同", "软件服务合同", "施工安装合同"];
  return Array.from({ length: 18 }, (_, index) => {
    const material = materials[index % materials.length];
    const isService = /服务|施工|软件/.test(material);
    const target = isService ? "服务项" : index % 5 === 0 ? "机房" : index % 3 === 0 ? "仓库" : "点位";
    return {
      id: `UM-${String(index + 1).padStart(4, "0")}`,
      status: index < 4 ? "待治理" : index < 10 ? "待确认" : "已建议",
      assetType: isService ? "软件/服务类" : index % 4 === 0 ? "机房设备" : "点位设备",
      material,
      supplier: suppliers[index % suppliers.length],
      contract: contracts[index % contracts.length],
      target,
    };
  });
}


function formatBytes(value) {
  const bytes = Number(value || 0);
  if (bytes >= 1024 ** 4) return `${(bytes / 1024 ** 4).toFixed(2)} TB`;
  if (bytes >= 1024 ** 3) return `${(bytes / 1024 ** 3).toFixed(2)} GB`;
  if (bytes >= 1024 ** 2) return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
  if (bytes >= 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${bytes} B`;
}

function documentCategoryLabel(category) {
  return documentCategoryMeta[category]?.[0] || category || "未分类";
}

function documentSubcategoryLabel(subcategory) {
  const labels = {
    front_sales_contract: "前向销售合同",
    back_procurement_contract: "后向采购合同",
    payment_first_delivery: "第一笔到货款",
    payment_second_delivery: "第二笔到货款",
    payment_third_delivery: "第三笔到货款/进度款",
    workload_list: "工作量清单",
    sales_procurement_analysis: "销采分析台账",
    contract_other: "其他合同资料",
  };
  return labels[subcategory] || subcategory || "";
}

function documentCategoryDisplay(record) {
  const category = documentCategoryLabel(record.documentCategory);
  const subcategory = documentSubcategoryLabel(record.documentSubcategory);
  return subcategory ? `${category} / ${subcategory}` : category;
}

function contractInsightDisplay(record) {
  const fields = record.extractedFields || {};
  const itemNo = fields.contractItemNo ? `${fields.contractItemNo} · ` : "";
  const parties = fields.partyA && fields.partyB ? `${fields.partyA} → ${fields.partyB}` : fields.counterparty || fields.contractName || "-";
  return `${itemNo}${parties}`;
}

function fieldFlagsLabel(record) {
  const fields = record.extractedFields || {};
  const labels = {
    supplement: "补充协议",
    attachment: "附件",
    voided: "作废线索",
    invoice: "发票",
    delivery_receipt: "送/收货单",
  };
  const flags = fields.flags || [];
  return flags.length ? flags.map((flag) => labels[flag] || flag).join("、") : "-";
}

function contractDirectionLabel(record) {
  const direction = record.extractedFields?.contractDirection || "";
  if (direction === "front_sales") return "前向销售";
  if (direction === "back_procurement") return "后向采购";
  return documentSubcategoryLabel(record.documentSubcategory) || "-";
}

function statusLabel(value) {
  const labels = {
    draft: "草稿",
    effective: "有效",
    archived: "已归档",
    voided: "已作废",
    pending_review: "待审核",
    pending: "待解析",
    success: "成功",
    failed: "失败",
    skipped: "跳过",
    normal: "正常",
    warning: "待治理",
    error: "异常",
  };
  return labels[value] || value || "-";
}

async function loadDocumentAssets() {
  state.documentAssets.loading = true;
  state.documentAssets.error = "";
  renderDocuments();
  try {
    let response = await fetch("/api/document-assets", { cache: "no-store" });
    if (!response.ok) {
      response = await fetch("./data/document-assets.json", { cache: "no-store" });
    }
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    state.documentAssets = {
      loading: false,
      error: "",
      reportPath: payload.reportPath || "",
      batchNo: payload.batchNo || "",
      generatedAt: payload.generatedAt || "",
      summary: payload.summary || {},
      records: payload.records || [],
    };
  } catch (error) {
    state.documentAssets = {
      ...state.documentAssets,
      loading: false,
      error: error.message || "文档资产报告读取失败",
      summary: null,
      records: [],
    };
  }
  renderDocuments();
}

function formatNumber(value) {
  return new Intl.NumberFormat("zh-CN").format(value || 0);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatMoney(value) {
  if (!value) return "0";
  if (value >= 100000000) return `${(value / 100000000).toFixed(2)} 亿`;
  if (value >= 10000) return `${(value / 10000).toFixed(1)} 万`;
  return formatNumber(value);
}

function formatMetricValue(value) {
  if (typeof value === "number") return formatNumber(value);
  if (typeof value === "string") return value;
  return formatNumber(value || 0);
}

function loadContractReviewState() {
  try {
    const storage = globalThis.localStorage;
    const raw = storage?.getItem(CONTRACT_REVIEW_STORAGE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    state.contractReview.decisions = parsed?.decisions || {};
  } catch (error) {
    state.contractReview.decisions = {};
  }
}

function saveContractReviewState() {
  try {
    const storage = globalThis.localStorage;
    storage?.setItem(
      CONTRACT_REVIEW_STORAGE_KEY,
      JSON.stringify({
        updatedAt: new Date().toISOString(),
        decisions: state.contractReview.decisions,
      }),
    );
  } catch (error) {
    // In restricted preview browsers localStorage can be unavailable; keep the
    // in-memory decision state so the workbench remains usable.
  }
}

function contractCandidateDecision(candidateId) {
  return state.contractReview.decisions[candidateId] || { status: "candidate" };
}

function setContractCandidateDecision(candidateId, status) {
  if (!candidateId) return;
  if (status === "candidate") {
    delete state.contractReview.decisions[candidateId];
  } else {
    state.contractReview.decisions[candidateId] = {
      status,
      reviewer: "当前用户",
      reviewedAt: new Date().toISOString(),
      evidence: status === "confirmed" ? "原型人工确认：合同链、金额、税率、分项名称待正式系统复核" : "原型人工否决：候选关系不进入正式资金计算",
    };
  }
  state.contractReview.selectedCandidateId = candidateId;
  saveContractReviewState();
  renderContractSystemViews();
  renderMetrics();
  renderContractStrip();
}

function contractReviewStats() {
  const candidates = state.contractRelationships.frontBackRelationshipCandidates || [];
  return candidates.reduce(
    (acc, candidate) => {
      const status = contractCandidateDecision(candidate.id).status || "candidate";
      if (status === "confirmed") acc.confirmed += 1;
      else if (status === "rejected") acc.rejected += 1;
      else acc.pending += 1;
      return acc;
    },
    { pending: 0, confirmed: 0, rejected: 0 },
  );
}

async function loadContractRelationships() {
  state.contractRelationships.loading = true;
  state.contractRelationships.error = "";
  renderContractSystemViews();
  try {
    const response = await fetch("./data/contract-relationships.json", { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    state.contractRelationships = {
      loading: false,
      error: "",
      summary: payload.summary || {},
      macroFlows: payload.macroFlows || [],
      contracts: payload.contracts || [],
      contractToMacroFlowMatches: payload.contractToMacroFlowMatches || [],
      frontBackRelationshipCandidates: payload.frontBackRelationshipCandidates || [],
      deviceCashflowSchema: payload.deviceCashflowSchema || null,
      ownerAuditTrackingModel: payload.ownerAuditTrackingModel || null,
      paymentStageTemplates: payload.paymentStageTemplates || [],
      selectedMacroFlowId: state.contractRelationships.selectedMacroFlowId || payload.macroFlows?.[0]?.id || "",
    };
  } catch (error) {
    state.contractRelationships = {
      ...state.contractRelationships,
      loading: false,
      error: error.message || "合同关系重构数据读取失败",
      summary: null,
      macroFlows: [],
      contracts: [],
      contractToMacroFlowMatches: [],
      frontBackRelationshipCandidates: [],
      deviceCashflowSchema: null,
      ownerAuditTrackingModel: null,
      paymentStageTemplates: [],
      selectedMacroFlowId: "",
    };
  }
  renderContractSystemViews();
  renderMetrics();
  renderContractStrip();
}

function districtClass(site) {
  return districtClassMap[site.district] || "district-other";
}

function statusClass(site) {
  return ["已安装", "已调试", "已验收", "运维中"].includes(site.status) ? "done" : "warn";
}

function perceptionType(site) {
  const value = (site.perception || "").trim();
  if (!value || value === "未标注" || value === "否" || value === "0") return "无";
  if (value.includes("雷") && value.includes("视")) return "雷视融合";
  if (value.includes("雷达")) return "毫米波雷达";
  if (value.includes("RSU")) return "RSU";
  if (value.includes("视频")) return "视频";
  return value === "是" ? "感知" : value;
}

function isVariableLane(site) {
  return Number(site.variableLaneCount || 0) > 0 || Boolean((site.variableLaneDirection || "").trim());
}

function deviceType(device) {
  const name = `${device.name || ""}${device.model || ""}`;
  if (/信号机|信号/.test(name)) return "信号系统";
  if (/摄像|视频|雷达|RSU|感知/.test(name)) return "感知设备";
  if (/抱杆箱|机柜|机箱/.test(name)) return "箱柜";
  if (/光模块|交换机|网络|光纤|物联网/.test(name)) return "网络通信";
  if (/施工费|服务|测评|开发|运维/.test(name)) return "服务项";
  return "其他设备";
}

function deviceRecordCount(site) {
  return site.devices?.length || site.deviceRecordCount || 0;
}

function deviceTypeCount(site) {
  return site.deviceTypeCount || new Set((site.devices || []).map(deviceType)).size || site.deviceCount || 0;
}

function districtBadge(site) {
  return `<span class="district-badge ${districtClass(site)}">${site.district}</span>`;
}

function issueText(site) {
  return site.issueCount > 0 ? `${site.issueCount} 项` : "-";
}


const siteFilterLabels = {
  district: "区域",
  type: "类型",
  vendor: "信号机厂商",
  perception: "感知点位类型",
  adaptive: "自适应",
  variableLane: "可变车道",
};

function siteFilterValue(site, key) {
  if (key === "district") return site.district || "未标注";
  if (key === "type") return site.type || "未标注";
  if (key === "vendor") return site.vendor || "未标注";
  if (key === "perception") return perceptionType(site);
  if (key === "adaptive") return site.adaptive ? "是" : "否";
  if (key === "variableLane") return isVariableLane(site) ? "是" : "否";
  return "";
}

function siteFilterValues(key) {
  if (key === "adaptive" || key === "variableLane") return ["是", "否"];
  return Array.from(new Set(sites.map((site) => siteFilterValue(site, key))))
    .filter(Boolean)
    .sort((a, b) => String(a).localeCompare(String(b), "zh-CN"));
}

function siteMultiFilterPass(value, selectedValues) {
  if (!selectedValues?.length) return true;
  return selectedValues.includes(value);
}

function renderSiteSummaryCards() {
  const target = $("#siteSummaryCards");
  if (!target) return;
  const cards = [
    ["路口点位数", sites.length],
    ["R1-点位数", sites.filter((site) => /^R1/i.test(perceptionType(site))).length],
    ["R2点位数", sites.filter((site) => /^R2/i.test(perceptionType(site))).length],
    ["礼让行人点位数", sites.filter((site) => /礼让行人/.test(perceptionType(site)) || /礼让行人/.test(site.type || "")).length],
    ["匝道汇入点位数", sites.filter((site) => site.type === "匝道汇入" || /匝道汇入/.test(perceptionType(site))).length],
    ["R3点位数", sites.filter((site) => /^R3/i.test(perceptionType(site))).length],
  ];
  target.innerHTML = cards
    .map(([label, value], index) => `<div class="site-summary-card tone-${index + 1}"><span>${label}</span><strong>${formatNumber(value)}</strong></div>`)
    .join("");
}

function renderSiteColumnFilters() {
  $$(`[data-site-filter-open]`).forEach((button) => {
    const key = button.dataset.siteFilterOpen;
    const count = state.siteColumnFilters[key]?.length || 0;
    button.classList.toggle("active", count > 0);
    button.textContent = count ? `筛${count}` : "筛";
  });
  renderSiteFilterPopover();
}

function renderSiteFilterPopover() {
  const popover = $("#siteFilterPopover");
  if (!popover) return;
  const key = state.activeSiteFilter;
  if (!key) {
    popover.hidden = true;
    popover.innerHTML = "";
    return;
  }
  const trigger = $(`[data-site-filter-open="${key}"]`);
  if (!trigger) {
    popover.hidden = true;
    return;
  }
  const selected = state.siteColumnFilters[key] || [];
  const values = siteFilterValues(key);
  const columnRect = trigger.closest("th")?.getBoundingClientRect() || trigger.getBoundingClientRect();
  popover.hidden = false;
  const width = Math.round(columnRect.width);
  popover.style.width = `${width}px`;
  popover.style.left = `${Math.min(window.innerWidth - width - 12, Math.max(12, columnRect.left))}px`;
  popover.style.top = `${columnRect.bottom + 6}px`;
  popover.innerHTML = `
    <div class="site-filter-popover-head">
      <strong>${siteFilterLabels[key]}</strong>
      <button type="button" data-site-filter-clear="${key}">清空</button>
    </div>
    <div class="site-filter-options">
      ${values
        .map(
          (value) => `
            <label>
              <input type="checkbox" data-site-filter-option="${key}" value="${value}" ${selected.includes(value) ? "checked" : ""} />
              <span>${value}</span>
            </label>
          `,
        )
        .join("")}
    </div>
  `;
}

function renderSiteColumnHandles() {
  const table = $("#siteTable");
  if (!table) return;
  table.querySelectorAll(".column-resizer").forEach((handle) => handle.remove());
  Array.from(table.querySelectorAll("th")).forEach((th, index, headers) => {
    if (index === headers.length - 1) return;
    const handle = document.createElement("button");
    handle.type = "button";
    handle.className = "column-resizer site-column-resizer";
    handle.setAttribute("aria-label", `调整 ${th.textContent.trim()} 列宽`);
    handle.dataset.columnIndex = String(index);
    handle.dataset.resizeTable = "site";
    th.appendChild(handle);
  });
}

function applySiteViewLayout() {
  const tableWrap = $("#tableView .table-wrap");
  const table = $("#siteTable");
  if (tableWrap && table) {
    const header = table.querySelector("thead");
    const firstRow = table.querySelector("tbody tr");
    const headerHeight = Math.ceil(header?.getBoundingClientRect().height || 48);
    const rowHeight = Math.ceil(firstRow?.getBoundingClientRect().height || 68);
    tableWrap.style.minHeight = "0px";
    tableWrap.style.height = "auto";
    tableWrap.style.maxHeight = "none";
    const top = tableWrap.getBoundingClientRect().top;
    const height = Math.max(headerHeight + 72, Math.floor(window.innerHeight - top - 24));
    tableWrap.style.height = `${height}px`;
    tableWrap.style.maxHeight = `${height}px`;
  }

  const mapLayout = $("#mapView .map-layout");
  if (mapLayout) {
    mapLayout.style.height = "auto";
    const top = mapLayout.getBoundingClientRect().top;
    const bottomPadding = 24;
    const height = Math.max(420, Math.floor(window.innerHeight - top - bottomPadding));
    mapLayout.style.height = `${height}px`;
  }
}

function applySiteViewLayoutSettled() {
  applySiteViewLayout();
  requestAnimationFrame(() => {
    applySiteViewLayout();
    setTimeout(applySiteViewLayout, 80);
  });
}

function startSiteColumnResize(event, handle) {
  event.preventDefault();
  event.stopPropagation();
  const table = $("#siteTable");
  const th = handle.closest("th");
  if (!table || !th) return;
  const columnIndex = Number(handle.dataset.columnIndex);
  const startX = event.clientX;
  const startWidth = th.getBoundingClientRect().width;
  const cells = Array.from(table.querySelectorAll(`tr > *:nth-child(${columnIndex + 1})`));
  const onMove = (moveEvent) => {
    const width = Math.max(columnIndex === 0 ? 54 : 76, startWidth + moveEvent.clientX - startX);
    cells.forEach((cell) => {
      cell.style.width = `${width}px`;
      cell.style.minWidth = `${width}px`;
      cell.style.maxWidth = `${width}px`;
    });
    if (state.activeSiteFilter) renderSiteFilterPopover();
  };
  const onUp = () => {
    document.removeEventListener("mousemove", onMove);
    document.removeEventListener("mouseup", onUp);
    document.body.classList.remove("resizing-column");
  };
  document.body.classList.add("resizing-column");
  document.addEventListener("mousemove", onMove);
  document.addEventListener("mouseup", onUp);
}

function importPanelElement() {
  if (state.panel === "ops" && $("#opsImportPanel")) return $("#opsImportPanel");
  return state.panel === "imports" && $("#importCenterPanel") ? $("#importCenterPanel") : $("#importPanel");
}

function closeImportPanel(panel = importPanelElement()) {
  if (!panel) return;
  panel.innerHTML = "";
  panel.classList.remove("open");
}

function filteredSites() {
  const q = state.query.trim().toLowerCase();
  const filters = state.siteColumnFilters;
  return sites.filter((site) => {
    const text = `${site.nodeId} ${site.name} ${site.vendor} ${site.district} ${site.type} ${perceptionType(site)}`.toLowerCase();
    const districtOk = siteMultiFilterPass(siteFilterValue(site, "district"), filters.district);
    const typeOk = siteMultiFilterPass(siteFilterValue(site, "type"), filters.type);
    const vendorOk = siteMultiFilterPass(siteFilterValue(site, "vendor"), filters.vendor);
    const perceptionOk = siteMultiFilterPass(siteFilterValue(site, "perception"), filters.perception);
    const adaptiveOk = siteMultiFilterPass(siteFilterValue(site, "adaptive"), filters.adaptive);
    const variableOk = siteMultiFilterPass(siteFilterValue(site, "variableLane"), filters.variableLane);
    return districtOk && typeOk && vendorOk && perceptionOk && adaptiveOk && variableOk && (!q || text.includes(q));
  }).sort((a, b) => (b.issueCount || 0) - (a.issueCount || 0));
}

function allDevices() {
  return sites.flatMap((site) =>
    site.devices.map((device) => ({
      ...device,
      nodeId: site.nodeId,
      siteName: site.name,
      district: site.district,
      type: deviceType(device),
    })),
  );
}

function setPanel(panelId) {
  state.panel = panelId;
  $$(".panel").forEach((panel) => panel.classList.toggle("active", panel.id === panelId));
  $$(".nav-item").forEach((item) => item.classList.toggle("active", item.dataset.panel === panelId));
  renderPageHeader();
  if (panelId === "overview") requestAnimationFrame(renderOverviewMap);
  if (panelId === "sites") requestAnimationFrame(applySiteViewLayoutSettled);
  if (panelId === "sites" && state.view === "map") requestAnimationFrame(() => { applySiteViewLayoutSettled(); renderMap($("#siteMap"), filteredSites()); });
  if (panelId === "mapAssets") renderMapAssets();
  if (panelId === "contracts") renderContracts();
  if (panelId === "documents") renderDocuments();
  if (panelId === "coordinates") requestAnimationFrame(renderCoordinateIssues);
  if (panelId === "ops") requestAnimationFrame(renderOps);
}

function renderPageHeader() {
  const [title, subtitle] = pageHeaders[state.panel] || pageHeaders.overview;
  $("#pageTitle").textContent = title;
  $("#pageSubtitle").textContent = subtitle;
  $(".page-heading")?.classList.toggle("sites-heading", state.panel === "sites");
}

function categoryLabel(category) {
  const labels = {
    map_xml: "MAP XML",
    map_json: "MAP JSON",
    map_svg: "SVG 预览",
    rsi_xml: "RSI XML",
    rsi_json: "RSI JSON",
    signal_excel: "信号机 Excel",
  };
  return labels[category] || category || "其他";
}

function mapAssetFileLabel(file) {
  if (file.category === "map_svg" && file.folder === "map_xml") return "MAP XML TO SVG";
  if (file.category === "map_svg" && file.folder === "map_json") return "MAP JSON TO SVG";
  if (file.category === "signal_excel") return "信号机 EXCEL";
  return categoryLabel(file.category);
}

function mapAssetFileRank(file) {
  if (file.category === "map_xml") return 1;
  if (file.category === "map_svg" && file.folder === "map_xml") return 2;
  if (file.category === "map_json") return 3;
  if (file.category === "map_svg" && file.folder === "map_json") return 4;
  if (file.category === "signal_excel") return 5;
  if (file.category === "rsi_xml") return 6;
  if (file.category === "rsi_json") return 7;
  return 20;
}

function orderedMapAssetFiles(asset) {
  return [...(asset.files || [])].sort((a, b) => mapAssetFileRank(a) - mapAssetFileRank(b) || String(a.fileName).localeCompare(String(b.fileName), "zh-CN"));
}

function mapAssetFileByPath(asset, relativePath) {
  return (asset.files || []).find((file) => file.relativePath === relativePath);
}

function formatAssetVersion(value) {
  const text = toText(value);
  if (!text) return "-";
  const compact = text.match(/^(20\d{2})(\d{2})(\d{2})T?(\d{2})?(\d{2})?/);
  if (compact) {
    const [, year, month, day, hour, minute] = compact;
    return `${year}-${month}-${day}${hour ? ` ${hour}:${minute || "00"}` : ""}`;
  }
  if (text.includes("T")) return text.slice(0, 16).replace("T", " ");
  return text;
}

function mapAssetSite(asset) {
  return sitesByNodeId.get(String(asset?.nodeId || "").trim()) || null;
}

function mapAssetDisplayName(asset) {
  return mapAssetSite(asset)?.name || asset.name || "-";
}

function mapAssetDisplayType(asset) {
  return mapAssetSite(asset)?.type || "未匹配";
}

async function loadMapAssets(force = false) {
  mapAssetState.loading = true;
  mapAssetState.error = "";
  renderMapAssets();
  try {
    const response = await fetch(`/api/map-assets${force ? `?t=${Date.now()}` : ""}`, { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    mapAssetState.root = payload.root || "";
    mapAssetState.generatedAt = payload.generatedAt || "";
    mapAssetState.summary = payload.summary || null;
    mapAssetState.intersections = Array.isArray(payload.intersections) ? payload.intersections : [];
    if (!state.selectedMapAssetId && mapAssetState.intersections.length) {
      const incomplete = mapAssetState.intersections.find((item) => item.status !== "完整");
      state.selectedMapAssetId = (incomplete || mapAssetState.intersections[0]).id;
    }
  } catch (error) {
    console.error("Map asset load failed.", error);
    mapAssetState.error = error.message || "地图资产目录读取失败";
  } finally {
    mapAssetState.loading = false;
    renderMapAssets();
  }
}

function mapAssetCompletenessText(asset) {
  return asset.missing?.length ? `缺${asset.missing.join("、")}` : "文件齐套";
}

const mapAssetFilterLabels = {
  siteType: "点位类型",
  district: "区域",
  completeness: "完整性",
  mapXml: "MAP-XML",
  mapJson: "MAP-JSON",
  signalExcel: "MAP-Excel",
  svg: "SVG",
};

function mapAssetFilterValues(key) {
  if (key === "siteType") return Array.from(new Set(mapAssetState.intersections.map(mapAssetDisplayType))).filter(Boolean).sort((a, b) => String(a).localeCompare(String(b), "zh-CN"));
  if (key === "district") return Array.from(new Set(mapAssetState.intersections.map((asset) => asset.district))).filter(Boolean).sort((a, b) => String(a).localeCompare(String(b), "zh-CN"));
  if (key === "completeness") return Array.from(new Set(mapAssetState.intersections.map(mapAssetCompletenessText))).filter(Boolean).sort((a, b) => String(a).localeCompare(String(b), "zh-CN"));
  return ["有", "无"];
}

function mapAssetMultiFilterPass(value, selectedValues) {
  if (!selectedValues?.length) return true;
  return selectedValues.includes(value);
}

function mapAssetCountFilterPass(count, selectedValues) {
  if (!selectedValues?.length) return true;
  if (selectedValues.includes("有") && count > 0) return true;
  if (selectedValues.includes("无") && count === 0) return true;
  return false;
}

function filteredMapAssets() {
  const q = state.query.trim().toLowerCase();
  const filters = state.mapAssetColumnFilters;
  return mapAssetState.intersections.filter((asset) => {
    const completeness = mapAssetCompletenessText(asset);
    const text = `${asset.nodeId} ${mapAssetDisplayName(asset)} ${mapAssetDisplayType(asset)} ${asset.name} ${asset.district} ${asset.status} ${asset.missing?.join(" ")}`.toLowerCase();
    const siteTypeOk = mapAssetMultiFilterPass(mapAssetDisplayType(asset), filters.siteType);
    const districtOk = mapAssetMultiFilterPass(asset.district, filters.district);
    const completenessOk = mapAssetMultiFilterPass(completeness, filters.completeness);
    const mapXmlOk = mapAssetCountFilterPass(asset.counts.map_xml || 0, filters.mapXml);
    const mapJsonOk = mapAssetCountFilterPass(asset.counts.map_json || 0, filters.mapJson);
    const signalExcelOk = mapAssetCountFilterPass(asset.counts.signal_excel || 0, filters.signalExcel);
    const svgOk = mapAssetCountFilterPass(asset.counts.map_svg || 0, filters.svg);
    return siteTypeOk && districtOk && completenessOk && mapXmlOk && mapJsonOk && signalExcelOk && svgOk && (!q || text.includes(q));
  });
}

function renderMapAssetColumnFilters() {
  $$("[data-map-asset-filter-open]").forEach((button) => {
    const key = button.dataset.mapAssetFilterOpen;
    const count = state.mapAssetColumnFilters[key]?.length || 0;
    button.classList.toggle("active", count > 0);
    button.textContent = count ? `筛${count}` : "筛";
  });
  renderMapAssetFilterPopover();
}

function renderMapAssetFilterPopover() {
  const popover = $("#mapAssetFilterPopover");
  if (!popover) return;
  const key = state.activeMapAssetFilter;
  if (!key) {
    popover.hidden = true;
    popover.innerHTML = "";
    return;
  }
  const trigger = $(`[data-map-asset-filter-open="${key}"]`);
  if (!trigger) {
    popover.hidden = true;
    return;
  }
  const selected = state.mapAssetColumnFilters[key] || [];
  const values = mapAssetFilterValues(key);
  const columnRect = trigger.closest("th")?.getBoundingClientRect() || trigger.getBoundingClientRect();
  popover.hidden = false;
  const width = Math.max(140, Math.round(columnRect.width));
  popover.style.width = `${width}px`;
  popover.style.left = `${Math.min(window.innerWidth - width - 12, Math.max(12, columnRect.left))}px`;
  popover.style.top = `${columnRect.bottom + 6}px`;
  popover.innerHTML = `
    <div class="map-asset-filter-popover-head">
      <strong>${mapAssetFilterLabels[key]}</strong>
      <button type="button" data-map-asset-filter-clear="${key}">清空</button>
    </div>
    <div class="map-asset-filter-options">
      ${values
        .map(
          (value) => `
            <label>
              <input type="checkbox" data-map-asset-filter-option="${key}" value="${value}" ${selected.includes(value) ? "checked" : ""} />
              <span>${value}</span>
            </label>
          `,
        )
        .join("")}
    </div>
  `;
}

function renderMapAssetMetrics() {
  const target = $("#mapAssetMetrics");
  if (!target) return;
  const summary = mapAssetState.summary;
  if (!summary) {
    target.innerHTML = "";
    return;
  }
  const cards = [
    ["路口资产", summary.intersectionTotal, "按区域 + NodeID 聚合"],
    ["源文件", summary.fileTotal, "XML / JSON / XLSX / SVG"],
    ["完整资产", summary.completeTotal, "MAP + RSI + 信号机 + SVG"],
    ["待补全", summary.incompleteTotal, "需继续校核或补文件"],
  ];
  target.innerHTML = cards
    .map(
      ([title, value, note]) => `
        <div class="metric">
          <span>${title}</span>
          <strong>${formatNumber(value)}</strong>
          <small>${note}</small>
        </div>
      `,
    )
    .join("");
}

function renderMapAssetRows() {
  const target = $("#mapAssetRows");
  if (!target) return;
  const rows = filteredMapAssets();
  $("#mapAssetCount").textContent = mapAssetState.loading ? "扫描中" : `${formatNumber(rows.length)} / ${formatNumber(mapAssetState.intersections.length)} 个路口`;
  target.innerHTML = rows
    .map((asset, index) => {
      const selected = asset.id === state.selectedMapAssetId;
      const completenessText = mapAssetCompletenessText(asset);
      const completenessClass = asset.missing?.length ? "warn" : "done";
      return `
        <tr class="${selected ? "selected-row" : ""}" data-open-map-asset="${asset.id}">
          <td class="map-asset-index">${index + 1}</td>
          <td><strong>${asset.nodeId}</strong></td>
          <td><strong>${mapAssetDisplayName(asset)}</strong></td>
          <td>${mapAssetDisplayType(asset)}</td>
          <td>${asset.district}</td>
          <td><span class="status-pill ${completenessClass}">${completenessText}</span></td>
          <td>${asset.counts.map_xml || 0}</td>
          <td>${asset.counts.map_json || 0}</td>
          <td>${asset.counts.signal_excel || 0}</td>
          <td>${asset.counts.map_svg || 0} SVG</td>
          <td>${formatAssetVersion(asset.latestVersion)}</td>
        </tr>
      `;
    })
    .join("");
  renderMapAssetColumnHandles();
}

function renderMapAssetColumnHandles() {
  const table = $(".map-asset-table table");
  if (!table) return;
  table.querySelectorAll(".column-resizer").forEach((handle) => handle.remove());
  Array.from(table.querySelectorAll("th")).forEach((th, index, headers) => {
    if (index === headers.length - 1) return;
    const handle = document.createElement("button");
    handle.type = "button";
    handle.className = "column-resizer";
    handle.setAttribute("aria-label", `调整 ${th.textContent.trim()} 列宽`);
    handle.dataset.columnIndex = String(index);
    th.appendChild(handle);
  });
}

function startMapAssetColumnResize(event, handle) {
  event.preventDefault();
  event.stopPropagation();
  const table = $(".map-asset-table table");
  const th = handle.closest("th");
  if (!table || !th) return;
  const columnIndex = Number(handle.dataset.columnIndex);
  const startX = event.clientX;
  const startWidth = th.getBoundingClientRect().width;
  const cells = Array.from(table.querySelectorAll(`tr > *:nth-child(${columnIndex + 1})`));
  const onMove = (moveEvent) => {
    const width = Math.max(72, startWidth + moveEvent.clientX - startX);
    cells.forEach((cell) => {
      cell.style.width = `${width}px`;
      cell.style.minWidth = `${width}px`;
      cell.style.maxWidth = `${width}px`;
    });
  };
  const onUp = () => {
    document.removeEventListener("mousemove", onMove);
    document.removeEventListener("mouseup", onUp);
    document.body.classList.remove("resizing-column");
  };
  document.body.classList.add("resizing-column");
  document.addEventListener("mousemove", onMove);
  document.addEventListener("mouseup", onUp);
}

function getMapAssetPreviewTransform(preview) {
  return {
    scale: Number(preview.dataset.scale || 1),
    x: Number(preview.dataset.x || 0),
    y: Number(preview.dataset.y || 0),
  };
}

function setMapAssetPreviewTransform(preview, transform) {
  const img = preview.querySelector("img");
  if (!img) return;
  const scale = Math.min(8, Math.max(0.5, transform.scale));
  const x = Number.isFinite(transform.x) ? transform.x : 0;
  const y = Number.isFinite(transform.y) ? transform.y : 0;
  preview.dataset.scale = String(scale);
  preview.dataset.x = String(x);
  preview.dataset.y = String(y);
  img.style.transform = `translate(${x}px, ${y}px) scale(${scale})`;
}

function handleMapAssetPreviewWheel(event, preview) {
  const img = preview.querySelector("img");
  if (!img) return;
  event.preventDefault();
  const current = getMapAssetPreviewTransform(preview);
  const nextScale = Math.min(8, Math.max(0.5, current.scale * (event.deltaY < 0 ? 1.12 : 0.88)));
  const rect = preview.getBoundingClientRect();
  const centerX = rect.left + rect.width / 2;
  const centerY = rect.top + rect.height / 2;
  const pointerX = event.clientX - centerX;
  const pointerY = event.clientY - centerY;
  const ratio = nextScale / current.scale;
  setMapAssetPreviewTransform(preview, {
    scale: nextScale,
    x: pointerX - (pointerX - current.x) * ratio,
    y: pointerY - (pointerY - current.y) * ratio,
  });
}

function startMapAssetPreviewDrag(event, preview) {
  if (!preview.querySelector("img")) return;
  event.preventDefault();
  const current = getMapAssetPreviewTransform(preview);
  mapAssetState.previewDrag = {
    preview,
    startX: event.clientX,
    startY: event.clientY,
    x: current.x,
    y: current.y,
  };
  preview.classList.add("dragging");
  document.body.classList.add("dragging-map-preview");
}

function moveMapAssetPreviewDrag(event) {
  const drag = mapAssetState.previewDrag;
  if (!drag) return;
  setMapAssetPreviewTransform(drag.preview, {
    scale: getMapAssetPreviewTransform(drag.preview).scale,
    x: drag.x + event.clientX - drag.startX,
    y: drag.y + event.clientY - drag.startY,
  });
}

function stopMapAssetPreviewDrag() {
  const drag = mapAssetState.previewDrag;
  if (!drag) return;
  drag.preview.classList.remove("dragging");
  document.body.classList.remove("dragging-map-preview");
  mapAssetState.previewDrag = null;
}

function renderMapAssetDetail() {
  const target = $("#mapAssetDetail");
  if (!target) return;
  const asset = mapAssetState.intersections.find((item) => item.id === state.selectedMapAssetId) || filteredMapAssets()[0];
  if (!asset) {
    target.innerHTML = `<div class="empty-detail">${mapAssetState.loading ? "正在扫描地图资产目录" : "暂无匹配地图资产"}</div>`;
    return;
  }
  state.selectedMapAssetId = asset.id;
  const mapXmlPreviewUrl = asset.mapXmlSvgPath ? `/api/map-assets/file?path=${encodeURIComponent(asset.mapXmlSvgPath)}` : "";
  const mapJsonPreviewUrl = asset.mapJsonSvgPath ? `/api/map-assets/file?path=${encodeURIComponent(asset.mapJsonSvgPath)}` : "";
  const mapXmlSvgFile = mapAssetFileByPath(asset, asset.mapXmlSvgPath);
  const mapJsonSvgFile = mapAssetFileByPath(asset, asset.mapJsonSvgPath);
  const displayName = mapAssetDisplayName(asset);
  target.innerHTML = `
    <div class="map-asset-detail-head">
      <span>${asset.district} / NodeID ${asset.nodeId}</span>
      <h3>${displayName}</h3>
      <strong class="status-pill ${asset.status === "完整" ? "done" : "warn"}">${asset.missing?.length ? `缺${asset.missing.join("、")}` : "文件齐套"}</strong>
      <a class="primary-btn map-asset-export" href="/api/map-assets/export?id=${encodeURIComponent(asset.id)}">导出当前点位地图资产</a>
    </div>
    <div class="map-asset-preview-stack">
      <section class="map-asset-preview-block">
        <h4>map_xml 目录 SVG</h4>
        <p>${mapXmlSvgFile ? `由 ${mapXmlSvgFile.fileName} 生成` : "未找到 map_xml 目录 SVG 文件"}</p>
        <div class="map-asset-preview" data-map-asset-preview>
          ${
            mapXmlPreviewUrl
              ? `<img src="${mapXmlPreviewUrl}" alt="${displayName} map_xml 目录 SVG 预览" />`
              : `<div class="empty-detail">map_xml 目录没有 SVG 预览文件</div>`
          }
        </div>
      </section>
      <section class="map-asset-preview-block">
        <h4>map_json 目录 SVG</h4>
        <p>${mapJsonSvgFile ? `由 ${mapJsonSvgFile.fileName} 生成` : "未找到 map_json 目录 SVG 文件"}</p>
        <div class="map-asset-preview" data-map-asset-preview>
          ${
            mapJsonPreviewUrl
              ? `<img src="${mapJsonPreviewUrl}" alt="${displayName} map_json 目录 SVG 预览" />`
              : `<div class="empty-detail">map_json 目录没有 SVG 预览文件</div>`
          }
        </div>
      </section>
    </div>
    <div class="map-asset-file-list">
      <div class="map-asset-file-title">点位文件</div>
      ${orderedMapAssetFiles(asset)
        .map(
          (file) => `
            <a href="/api/map-assets/file?path=${encodeURIComponent(file.relativePath)}" target="_blank" rel="noreferrer">
              <strong>${mapAssetFileLabel(file)}</strong>
              <span>${file.fileName}</span>
              <small>${formatAssetVersion(file.version || file.modifiedAt)} / ${(file.size / 1024).toFixed(1)} KB</small>
            </a>
          `,
        )
        .join("")}
    </div>
  `;
}

function exportMapAssetList(assets, fileName, sheetName, emptyMessage) {
  if (!assets.length) {
    alert(emptyMessage);
    return;
  }
  const rows = [
    ["序号", "NodeID", "点位名称", "点位类型", "区域", "缺失项", "MAP-XML", "MAP-JSON", "MAP-Excel", "SVG", "RSI-XML", "RSI-JSON", "源文件数", "最新版本"],
    ...assets.map((asset, index) => [
      index + 1,
      asset.nodeId,
      mapAssetDisplayName(asset),
      mapAssetDisplayType(asset),
      asset.district,
      asset.missing?.length ? asset.missing.join("、") : "文件齐套",
      asset.counts.map_xml || 0,
      asset.counts.map_json || 0,
      asset.counts.signal_excel || 0,
      asset.counts.map_svg || 0,
      asset.counts.rsi_xml || 0,
      asset.counts.rsi_json || 0,
      asset.fileTotal || asset.files?.length || 0,
      formatAssetVersion(asset.latestVersion),
    ]),
  ];
  writeWorkbookFile(fileName, sheetName, rows, [
    { wch: 8 },
    { wch: 12 },
    { wch: 34 },
    { wch: 14 },
    { wch: 12 },
    { wch: 24 },
    { wch: 10 },
    { wch: 10 },
    { wch: 11 },
    { wch: 8 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 18 },
  ]);
}

function exportAllMapAssets() {
  exportMapAssetList(mapAssetState.intersections, "全部地图资产点位列表.xlsx", "全部地图资产", "当前地图资产索引为空，请先扫描目录。");
}

function exportIncompleteMapAssets() {
  exportMapAssetList(
    mapAssetState.intersections.filter((asset) => asset.missing?.length),
    "不完整地图资产点位列表.xlsx",
    "不完整地图资产",
    "当前地图资产索引中没有完整性缺失的点位。",
  );
}

function renderMapAssets() {
  const status = $("#mapAssetStatus");
  if (!status) return;
  status.classList.toggle("error", Boolean(mapAssetState.error));
  if (mapAssetState.error) {
    status.textContent = `地图资产读取失败：${mapAssetState.error}`;
  } else if (mapAssetState.loading) {
    status.textContent = "正在扫描地图资产目录，首次扫描会读取 XML、JSON、XLSX、SVG 文件索引";
  } else {
    status.textContent = `数据目录：${mapAssetState.root} · 索引时间：${formatAssetVersion(mapAssetState.generatedAt)}`;
  }
  renderMapAssetColumnFilters();
  renderMapAssetMetrics();
  renderMapAssetRows();
  renderMapAssetDetail();
}

function loadVisualTheme() {
  try {
    const saved = localStorage.getItem(VISUAL_THEME_STORAGE_KEY);
    return visualThemes.includes(saved) ? saved : "macos";
  } catch {
    return "macos";
  }
}

function saveVisualTheme(theme) {
  try {
    localStorage.setItem(VISUAL_THEME_STORAGE_KEY, theme);
  } catch {
    // Storage can be unavailable in restricted browser contexts; the live toggle still works.
  }
}

function applyVisualTheme(theme = state.visualTheme) {
  state.visualTheme = visualThemes.includes(theme) ? theme : "macos";
  document.body.dataset.visualTheme = state.visualTheme;
  const toggle = $("#visualThemeToggle");
  if (!toggle) return;
  const themeLabels = {
    macos: "大屏风格",
    command: "轨迹风格",
    trajectory: "MacOS 风格",
  };
  const themeTitles = {
    macos: "切换为原来的大屏渲染风格",
    command: "切换为轨迹实时指标风格",
    trajectory: "切换为跟随系统外观的 MacOS 风格",
  };
  toggle.textContent = themeLabels[state.visualTheme];
  toggle.setAttribute("aria-pressed", String(state.visualTheme !== "macos"));
  toggle.title = themeTitles[state.visualTheme];
}

function toggleVisualTheme() {
  const currentIndex = visualThemes.indexOf(state.visualTheme);
  applyVisualTheme(visualThemes[(currentIndex + 1) % visualThemes.length]);
  saveVisualTheme(state.visualTheme);
}

function renderMetrics() {
  $("#metricSiteTotal").textContent = formatNumber(data.stats.siteTotal);
  $("#metricDeviceRows").textContent = formatNumber(data.stats.deviceRows);
  $("#metricUnmatchedRows").textContent = formatNumber(data.stats.unmatchedRows);
  const cr = state.contractRelationships;
  const summary = cr.summary || {};
  const contractMetric = $("#metricContractRows");
  if (contractMetric) {
    if (cr.loading) contractMetric.textContent = "读取中";
    else if (cr.error) contractMetric.textContent = "需重建";
    else contractMetric.textContent = `前向 ${formatNumber(summary.frontContractCount || 0)} / 后向 ${formatNumber(summary.backContractCount || 0)}`;
  }
  const contractMetricHint = $('.metric-link[data-panel-link="contracts"] small');
  if (contractMetricHint) {
    const candidates = summary.frontBackCandidateCount ?? cr.frontBackRelationshipCandidates.length;
    contractMetricHint.textContent = cr.loading ? "读取合同关系" : `${formatNumber(candidates)} 条候选待确认`;
  }
  const unmatchedHint = $("#overviewUnmatchedCard small");
  if (unmatchedHint) unmatchedHint.textContent = `${formatNumber(data.stats.unmatchedRows)} 条待治理`;
}

function renderStatusProgress() {
  const statusOrder = ["已规划", "已勘察", "待施工", "施工中", "已安装", "已调试", "已验收", "运维中"];
  const counts = statusOrder.map((status) => ({
    status,
    count: sites.filter((site) => site.status === status).length,
  }));
  const max = Math.max(...counts.map((item) => item.count), 1);
  $("#statusProgress").innerHTML = counts
    .filter((item) => item.count > 0)
    .map(
      (item) => `
        <div class="progress-item">
          <span>${item.status}</span>
          <div class="bar"><i style="width:${Math.max(8, (item.count / max) * 100)}%"></i></div>
          <strong>${item.count}</strong>
        </div>
      `,
    )
    .join("");
}

function renderContractStrip() {
  const holder = $("#contractStrip");
  if (!holder) return;
  const cr = state.contractRelationships;
  const summary = cr.summary || {};
  const reviewStats = contractReviewStats();
  if (cr.loading) {
    holder.innerHTML = `
      <div class="contract-row contract-row-muted"><div><b>正在读取合同关系重构数据</b><span>读取 Excel 宏观合同流、Word 合同解析和前后向候选关系</span></div><span class="status-pill warn">读取中</span></div>
      <div class="contract-row contract-row-muted"><div><b>设备级资金流闸门</b><span>未确认的候选关系不会进入付款或毛利计算</span></div><span class="status-pill warn">锁定</span></div>`;
    return;
  }
  if (cr.error) {
    holder.innerHTML = `<div class="contract-row contract-row-muted"><div><b>合同关系数据读取失败</b><span>${escapeHtml(cr.error)}</span></div><span class="status-pill error">失败</span></div>`;
    return;
  }
  const rows = [
    ["Word 合同边界", `共 ${formatNumber(summary.contractDocumentCount || 0)} 份，前向 ${formatNumber(summary.frontContractCount || 0)} / 后向 ${formatNumber(summary.backContractCount || 0)}`, "文本边界", "done"],
    ["前后向多对多候选", `${formatNumber(summary.frontBackCandidateCount ?? cr.frontBackRelationshipCandidates.length)} 条候选；${formatNumber(reviewStats.pending)} 条待确认`, "待确认", reviewStats.pending ? "warn" : "done"],
    ["资金计算闸门", "候选关系确认后才能拆到设备级资金流单元；未确认关系保持锁定", "锁定", "warn"],
    ["审计与阶段付款", "业主设备级审计结果决定上游回款，下游按预付/到货/安装/验收/审计/质保阶段释放", "设备级", "done"],
  ];
  holder.innerHTML = rows
    .map(
      ([title, desc, status, cls]) => `
        <div class="contract-row">
          <div>
            <b>${escapeHtml(title)}</b>
            <span>${escapeHtml(desc)}</span>
          </div>
          <span class="status-pill ${cls}">${escapeHtml(status)}</span>
        </div>
      `,
    )
    .join("");
}

function showImportPanel(type) {
  const config = importTemplates[type];
  if (!config) return;
  const panel = importPanelElement();
  state.importType = type;
  state.importTask = null;
  panel.innerHTML = `
    <div class="import-card">
      <div>
        <span>导入入口</span>
        <strong>${config.title}</strong>
        <p>${config.purpose}</p>
      </div>
      <dl>
        <dt>模板格式</dt>
        <dd>${config.template}</dd>
        <dt>导入结果</dt>
        <dd>${config.result}</dd>
      </dl>
      <div class="import-steps">
        <span class="active">上传文件</span>
        <span>字段识别</span>
        <span>预校验</span>
        <span>预览</span>
        <span>确认导入</span>
        <span>导入报告</span>
      </div>
      <div class="import-footer">
        <button class="primary-btn" data-choose-import="${type}">选择 Excel 文件</button>
        <button class="ghost-btn" data-simulate-import="${type}">使用模板样例模拟导入</button>
      </div>
    </div>
  `;
  panel.classList.add("open");
}

function expectedFileName(type) {
  if (type === "sites") return "点位管理表-无锡车路云.xlsx";
  if (type === "roadsideStatus") return "设备数据0518(1).xlsx";
  return "设备安装位置表.xlsx";
}

function importInputForType(type) {
  if (type === "sites") return $("#siteImportInput");
  if (type === "roadsideStatus") return $("#roadsideStatusImportInput");
  return $("#deviceImportInput");
}

function fileLooksLikeTemplate(type, fileName) {
  const normalized = fileName.replace(/\s+/g, "");
  if (type === "sites") return normalized.includes("点位管理表") && normalized.endsWith(".xlsx");
  if (type === "roadsideStatus") return normalized.includes("设备数据") && normalized.endsWith(".xlsx");
  return normalized.includes("设备安装位置") && normalized.endsWith(".xlsx");
}

function rowHasAny(row, keys) {
  return keys.some((key) => Object.prototype.hasOwnProperty.call(row, key) && toText(row[key]));
}

function isRoadsideStatusSourceRow(row) {
  const hasIdentity = rowHasAny(row, ["设备编号", "产品名称", "设备类型名称", "设备类型编码"]);
  const hasOpsStatus = rowHasAny(row, ["状态", "在线状态", "启用状态", "是否启用"]);
  const hasLocation = rowHasAny(row, ["经度", "纬度", "路口ID", "设备位置", "关联路口"]);
  return hasIdentity && hasOpsStatus && hasLocation;
}

function toText(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

function toNumber(value) {
  const text = toText(value).replace(/,/g, "");
  const number = Number.parseFloat(text);
  return Number.isFinite(number) ? number : 0;
}

function toBool(value) {
  const text = toText(value);
  return ["是", "有", "1", "true", "TRUE", "√", "Y", "y"].includes(text);
}

function pick(row, candidates) {
  const keys = Object.keys(row);
  for (const candidate of candidates) {
    if (Object.prototype.hasOwnProperty.call(row, candidate)) return row[candidate];
  }
  for (const candidate of candidates) {
    const key = keys.find((item) => item.replace(/\s+/g, "").includes(candidate.replace(/\s+/g, "")));
    if (key) return row[key];
  }
  return "";
}

function normalizeDistrict(value) {
  const text = toText(value);
  if (text.includes("锡山")) return "锡山";
  if (text.includes("新吴") || text.includes("经开")) return "新吴";
  if (text.includes("惠山")) return "惠山";
  if (text.includes("梁溪")) return "梁溪";
  if (text.includes("滨湖")) return "滨湖";
  return text || "未标注";
}

function transformLat(x, y) {
  let ret = -100 + 2 * x + 3 * y + 0.2 * y * y + 0.1 * x * y + 0.2 * Math.sqrt(Math.abs(x));
  ret += ((20 * Math.sin(6 * x * Math.PI) + 20 * Math.sin(2 * x * Math.PI)) * 2) / 3;
  ret += ((20 * Math.sin(y * Math.PI) + 40 * Math.sin((y / 3) * Math.PI)) * 2) / 3;
  ret += ((160 * Math.sin((y / 12) * Math.PI) + 320 * Math.sin((y * Math.PI) / 30)) * 2) / 3;
  return ret;
}

function transformLng(x, y) {
  let ret = 300 + x + 2 * y + 0.1 * x * x + 0.1 * x * y + 0.1 * Math.sqrt(Math.abs(x));
  ret += ((20 * Math.sin(6 * x * Math.PI) + 20 * Math.sin(2 * x * Math.PI)) * 2) / 3;
  ret += ((20 * Math.sin(x * Math.PI) + 40 * Math.sin((x / 3) * Math.PI)) * 2) / 3;
  ret += ((150 * Math.sin((x / 12) * Math.PI) + 300 * Math.sin((x / 30) * Math.PI)) * 2) / 3;
  return ret;
}

function outOfChina(lng, lat) {
  return lng < 72.004 || lng > 137.8347 || lat < 0.8293 || lat > 55.8271;
}

function wgs84ToGcj02(lng, lat) {
  if (!lng || !lat || outOfChina(lng, lat)) return { lng, lat, valid: false };
  const a = 6378245;
  const ee = 0.006693421622965943;
  let dLat = transformLat(lng - 105, lat - 35);
  let dLng = transformLng(lng - 105, lat - 35);
  const radLat = (lat / 180) * Math.PI;
  let magic = Math.sin(radLat);
  magic = 1 - ee * magic * magic;
  const sqrtMagic = Math.sqrt(magic);
  dLat = (dLat * 180) / (((a * (1 - ee)) / (magic * sqrtMagic)) * Math.PI);
  dLng = (dLng * 180) / ((a / sqrtMagic) * Math.cos(radLat) * Math.PI);
  return { lng: lng + dLng, lat: lat + dLat, valid: true };
}

function normalizeRoadsideStatusDevice(device) {
  if (!device) return device;
  if (device.coordinateConvertMethod) return device;
  const rawLng = toNumber(device.lngCgcs ?? device.rawLng ?? device.lng);
  const rawLat = toNumber(device.latCgcs ?? device.rawLat ?? device.lat);
  const converted = wgs84ToGcj02(rawLng, rawLat);
  return {
    ...device,
    lngCgcs: rawLng,
    latCgcs: rawLat,
    lng: converted.lng || rawLng,
    lat: converted.lat || rawLat,
    lngGcj: converted.lng || rawLng,
    latGcj: converted.lat || rawLat,
    coordinateConvertMethod: "cgcs2000_as_wgs84_to_gcj02",
    coordinateConvertStatus: converted.valid ? "success" : "failed",
  };
}

function normalizeRoadsideStatusRows(rows) {
  return (rows || []).map(normalizeRoadsideStatusDevice);
}

function applySavedRoadsideStatusState(saved) {
  if (!saved || !Array.isArray(saved.currentRows) || !Array.isArray(saved.archives)) return false;
  roadsideStatusState.currentDate = saved.currentDate || roadsideStatusState.currentDate;
  roadsideStatusState.currentRows = normalizeRoadsideStatusRows(saved.currentRows);
  roadsideStatusState.archives = saved.archives
    .filter((item) => item?.importDate && Array.isArray(item.rows))
    .map((item) => ({
      ...item,
      rows: normalizeRoadsideStatusRows(item.rows),
    }));
  state.opsDate = roadsideStatusState.currentDate;
  state.opsCalendarCursor = state.opsDate.slice(0, 7);
  return true;
}

function roadsideStatusStatePayload() {
  return {
    currentDate: roadsideStatusState.currentDate,
    currentRows: roadsideStatusState.currentRows,
    archives: roadsideStatusState.archives,
  };
}

function setPersistenceStatus(backendAvailable, lastSavedAt = persistenceState.lastSavedAt) {
  persistenceState.backendAvailable = backendAvailable;
  persistenceState.lastSavedAt = lastSavedAt;
  renderPersistenceStatus();
}

function renderPersistenceStatus() {
  const target = $("#opsPersistenceStatus");
  if (!target) return;
  const connected = persistenceState.backendAvailable;
  target.classList.toggle("connected", connected);
  target.classList.toggle("disconnected", !connected);
  target.textContent = connected
    ? `本地文件持久化已连接${persistenceState.lastSavedAt ? ` · ${persistenceState.lastSavedAt}` : ""}`
    : "本地文件持久化未连接，当前只能暂存到浏览器缓存";
}

async function loadRoadsideStatusState() {
  try {
    const response = await fetch("/api/roadside-status-state", { cache: "no-store" });
    if (response.ok) {
      const saved = await response.json();
      if (applySavedRoadsideStatusState(saved)) {
        setPersistenceStatus(true, saved.updatedAt ? new Date(saved.updatedAt).toLocaleString("zh-CN", { hour12: false }) : null);
        return;
      }
    }
  } catch (error) {
    console.info("Roadside status file backend unavailable, using localStorage fallback.");
  }
  setPersistenceStatus(false);
  try {
    const raw = window.localStorage?.getItem(ROADSIDE_STATUS_STORAGE_KEY);
    if (!raw) return;
    applySavedRoadsideStatusState(JSON.parse(raw));
  } catch (error) {
    console.warn("Roadside status localStorage restore failed.", error);
  }
}

async function saveRoadsideStatusState() {
  const payload = roadsideStatusStatePayload();
  let backendSaved = false;
  try {
    const response = await fetch("/api/roadside-status-state", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const result = await response.json();
    backendSaved = true;
    setPersistenceStatus(true, result.updatedAt ? new Date(result.updatedAt).toLocaleString("zh-CN", { hour12: false }) : new Date().toLocaleString("zh-CN", { hour12: false }));
  } catch (error) {
    console.info("Roadside status file backend unavailable, saved to localStorage only.");
    setPersistenceStatus(false);
  }
  try {
    window.localStorage?.setItem(ROADSIDE_STATUS_STORAGE_KEY, JSON.stringify(payload));
  } catch (error) {
    console.warn("Roadside status localStorage save failed.", error);
  }
  return backendSaved;
}

function readWorkbookFile(file) {
  if (!window.XLSX) {
    return Promise.reject(new Error("缺少 XLSX 解析库，请确认 prototype/vendor/xlsx.full.min.js 已加载。"));
  }
  return file.arrayBuffer().then((buffer) => window.XLSX.read(buffer, { type: "array", cellDates: true }));
}

function rowsFromSheet(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return window.XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
}

function firstNonEmptySheet(workbook) {
  return workbook.SheetNames.find((name) => rowsFromSheet(workbook, name).length) || workbook.SheetNames[0];
}

function parseSiteRow(row, index) {
  const rawLng = toNumber(pick(row, ["GCJ-02 经度", "高德展示经度", "原始规划经度", "CGCS2000 经度", "经度"]));
  const rawLat = toNumber(pick(row, ["GCJ-02 纬度", "高德展示纬度", "原始规划纬度", "CGCS2000 纬度", "纬度"]));
  const hasGcj = toText(pick(row, ["GCJ-02 经度", "高德展示经度"])) && toText(pick(row, ["GCJ-02 纬度", "高德展示纬度"]));
  const converted = hasGcj ? { lng: rawLng, lat: rawLat, valid: Boolean(rawLng && rawLat) } : wgs84ToGcj02(rawLng, rawLat);
  const variableLaneCount = toNumber(pick(row, ["可变车道数量"]));
  const variableLaneDirection = toText(pick(row, ["可变车道方向"]));
  const nodeId = toText(pick(row, ["NodeID", "node_id", "节点编号"]));
  const name = toText(pick(row, ["点位名称", "路口名称", "名称"]));
  const site = {
    serialNo: toText(pick(row, ["序号"])) || String(index + 1),
    nodeId,
    crossId: toText(pick(row, ["CrossID", "cross_id"])),
    name: name || `未命名点位 ${index + 1}`,
    englishName: toText(pick(row, ["点位英文名称", "路口英文名称", "NodeName", "英文"])),
    district: normalizeDistrict(pick(row, ["行政区域", "区域", "区县"])),
    originalDistrict: toText(pick(row, ["行政区域", "区域", "区县"])),
    type: toText(pick(row, ["点位类型", "路口类型", "类型"])) || "未标注",
    vendor: toText(pick(row, ["信号机厂商", "信号厂商", "厂商"])) || "待确认",
    perception: toText(pick(row, ["感知点位类型", "感知点位", "感知类型"])) || "未标注",
    adaptive: toBool(pick(row, ["是否自适应路口", "是否为自适应路口", "是否自适应点位", "自适应路口", "自适应点位"])),
    scope: toBool(pick(row, ["是否计入本期信号系统改造范围", "本期改造范围"])),
    variableLaneCount,
    variableLaneDirection,
    lngCgcs: rawLng,
    latCgcs: rawLat,
    lngGcj: converted.lng || 120.31,
    latGcj: converted.lat || 31.49,
    coordinateConvertMethod: hasGcj ? "source_gcj02" : "cgcs2000_as_wgs84_to_gcj02",
    status: "已规划",
    deviceRecordCount: 0,
    deviceTypeCount: 0,
    installedQty: 0,
    issueCount: converted.valid ? 0 : 1,
    archiveCompleteness: 30,
    devices: [],
  };
  return site;
}

function parseDeviceRow(row) {
  return {
    nodeId: toText(pick(row, ["NodeID", "node_id", "节点编号"])),
    siteName: toText(pick(row, ["点位名称", "路口名称"])),
    position: toText(pick(row, ["安装位置", "部署位置"])),
    supplier: toText(pick(row, ["供应商单位", "供应商", "厂商"])),
    materialCode: toText(pick(row, ["物料编码", "物料编号", "编码"])),
    name: toText(pick(row, ["物料名称", "设备名称", "名称"])),
    contract: toText(pick(row, ["后项合同名称", "后向合同名称", "合同名称", "合同"])),
    model: toText(pick(row, ["品牌型号", "规格型号", "型号"])),
    unit: toText(pick(row, ["单位"])),
    contractQty: toNumber(pick(row, ["后项合同数量", "合同数量", "数量"])),
    deliveredQty: toNumber(pick(row, ["已送货", "送货数量", "送货单数量"])),
    inboundQty: toNumber(pick(row, ["入库单数量", "已入库", "入库数量"])),
    pickedQty: toNumber(pick(row, ["本期领料数量", "领料数量", "已领料"])),
    installedQty: toNumber(pick(row, ["已安装", "安装数量"])),
    warehouseQty: toNumber(pick(row, ["在仓库", "仓库数量"])),
    backAmount: toNumber(pick(row, ["后项金额", "金额", "合同金额"])),
    issue: toText(pick(row, ["开箱验收问题", "开箱验收单问题", "问题", "异常"])),
  };
}

function parseRoadsideStatusRow(row, index) {
  const rawLng = toNumber(pick(row, ["经度", "CGCS2000 经度", "原始经度"]));
  const rawLat = toNumber(pick(row, ["纬度", "CGCS2000 纬度", "原始纬度"]));
  const hasGcj = toText(pick(row, ["GCJ-02 经度", "高德展示经度"])) && toText(pick(row, ["GCJ-02 纬度", "高德展示纬度"]));
  const gcjLng = toNumber(pick(row, ["GCJ-02 经度", "高德展示经度"]));
  const gcjLat = toNumber(pick(row, ["GCJ-02 纬度", "高德展示纬度"]));
  const converted = hasGcj ? { lng: gcjLng, lat: gcjLat, valid: Boolean(gcjLng && gcjLat) } : wgs84ToGcj02(rawLng, rawLat);
  return {
    deviceId: toText(pick(row, ["设备编号", "device_id"])) || `RS-${index + 1}`,
    productName: toText(pick(row, ["产品名称", "设备名称"])),
    vendorCode: toText(pick(row, ["厂商编码"])),
    vendorName: toText(pick(row, ["厂商名称", "供应商", "厂商"])),
    deviceTypeCode: toText(pick(row, ["设备类型编码"])),
    deviceTypeName: toText(pick(row, ["设备类型名称", "设备类型"])),
    nodeType: toText(pick(row, ["节点类型"])),
    serialNo: toText(pick(row, ["设备序列号", "序列号"])),
    devicePosition: toText(pick(row, ["设备位置", "位置"])),
    lngCgcs: rawLng,
    latCgcs: rawLat,
    lng: converted.lng || rawLng,
    lat: converted.lat || rawLat,
    lngGcj: converted.lng || rawLng,
    latGcj: converted.lat || rawLat,
    coordinateConvertMethod: hasGcj ? "source_gcj02" : "cgcs2000_as_wgs84_to_gcj02",
    coordinateConvertStatus: converted.valid ? "success" : "failed",
    intersectionId: toText(pick(row, ["路口ID", "节点编号"])),
    intersectionType: toText(pick(row, ["路口类型"])),
    area: normalizeDistrict(pick(row, ["区域"])),
    relatedIntersection: toText(pick(row, ["关联路口"])),
    ipAddress: toText(pick(row, ["IP地址", "IP"])),
    purpose: toText(pick(row, ["设备用途"])),
    ownerOrg: toText(pick(row, ["归属单位"])),
    opsOrg: toText(pick(row, ["运维单位"])),
    status: toText(pick(row, ["状态", "在线状态"])) || "未知",
    installPosition: "",
    policeDeviceId: toText(pick(row, ["设备编号(交警)", "交警设备编号"])),
    enabledStatus: toText(pick(row, ["启用状态", "是否启用"])) || "未标注",
    sourceRow: index + 2,
  };
}

async function createImportTaskFromWorkbook(type, file) {
  const workbook = await readWorkbookFile(file);
  if (type === "sites") return createSiteImportTask(file, workbook);
  if (type === "devices") return createDeviceImportTask(file, workbook);
  return createRoadsideStatusImportTask(file, workbook);
}

function createSiteImportTask(file, workbook) {
  const sheetName = firstNonEmptySheet(workbook);
  const rows = rowsFromSheet(workbook, sheetName);
  const importedSites = rows
    .map(parseSiteRow)
    .filter((site) => site.nodeId || site.name);
  const nodeIds = importedSites.map((site) => site.nodeId).filter(Boolean);
  const duplicateNodeIdCount = nodeIds.length - new Set(nodeIds).size;
  const coordinateErrorCount = importedSites.filter((site) => site.issueCount > 0).length;
  const dictionaryErrorCount = importedSites.filter((site) => ["未标注", "待确认"].includes(site.type) || site.vendor === "待确认").length;
  return {
    id: `sites-${Date.now()}`,
    type: "sites",
    fileName: file.name,
    fileSize: file.size,
    sourceSheet: sheetName,
    templateMatched: fileLooksLikeTemplate("sites", file.name),
    totalRows: rows.length,
    validRows: importedSites.length,
    invalidRows: rows.length - importedSites.length,
    duplicateNodeIdCount,
    coordinateErrorCount,
    dictionaryErrorCount,
    sampleRows: importedSites.slice(0, 8),
    importedSites,
    confirmed: false,
  };
}

function createDeviceImportTask(file, workbook) {
  const mainSheet = workbook.SheetNames.find((name) => name === "Sheet1") || firstNonEmptySheet(workbook);
  const unmatchedSheet = workbook.SheetNames.find((name) => name.includes("未匹配"));
  const mainRows = rowsFromSheet(workbook, mainSheet);
  const unmatchedSheetRows = unmatchedSheet ? rowsFromSheet(workbook, unmatchedSheet) : [];
  const importedDevices = mainRows
    .map(parseDeviceRow)
    .filter((device) => device.nodeId || device.name || device.materialCode);
  const currentNodeIds = new Set(sites.map((site) => site.nodeId));
  const matchedRows = importedDevices.filter((device) => currentNodeIds.has(device.nodeId)).length;
  const matchFailedRows = importedDevices.length - matchedRows;
  return {
    id: `devices-${Date.now()}`,
    type: "devices",
    fileName: file.name,
    fileSize: file.size,
    sourceSheet: mainSheet,
    templateMatched: fileLooksLikeTemplate("devices", file.name),
    totalRows: mainRows.length + unmatchedSheetRows.length,
    validRows: importedDevices.length,
    invalidRows: mainRows.length - importedDevices.length,
    sheet1TotalRows: mainRows.length,
    matchedRows,
    matchFailedRows,
    unmatchedSheetRows: unmatchedSheetRows.length + matchFailedRows,
    unpackingIssueRows: importedDevices.filter((device) => device.issue).length,
    quantityErrorRows: importedDevices.filter((device) => {
      const qty = device.contractQty || device.deliveredQty || device.inboundQty || device.installedQty;
      return qty && (device.installedQty > qty || device.inboundQty > qty);
    }).length,
    sampleRows: importedDevices.slice(0, 8),
    importedDevices,
    importedUnmatchedRows: unmatchedSheetRows,
    confirmed: false,
  };
}

function createRoadsideStatusImportTask(file, workbook) {
  const sheetName = workbook.SheetNames.find((name) => name === "Sheet1") || firstNonEmptySheet(workbook);
  const rows = rowsFromSheet(workbook, sheetName);
  const schemaMatchedRows = rows.filter(isRoadsideStatusSourceRow).length;
  const importedRows = rows
    .filter(isRoadsideStatusSourceRow)
    .map(parseRoadsideStatusRow)
    .filter((device) => device.deviceId || device.productName || device.deviceTypeName);
  const enabledRows = importedRows.filter((device) => isRoadsideEnabled(device));
  const offlineRows = enabledRows.filter((device) => device.status === "离线").length;
  const abnormalRows = enabledRows.filter((device) => isRoadsideAbnormal(device)).length;
  const dateMatch = file.name.match(/(20\\d{2})[-_年.]?(\\d{1,2})[-_月.]?(\\d{1,2})/) || file.name.match(/(\\d{2})(\\d{2})/);
  const importDate = dateMatch && dateMatch.length >= 4
    ? `${dateMatch[1]}-${String(dateMatch[2]).padStart(2, "0")}-${String(dateMatch[3]).padStart(2, "0")}`
    : new Date().toISOString().slice(0, 10);
  return {
    id: `roadside-${Date.now()}`,
    type: "roadsideStatus",
    fileName: file.name,
    fileSize: file.size,
    sourceSheet: sheetName,
    templateMatched: fileLooksLikeTemplate("roadsideStatus", file.name) && schemaMatchedRows > 0,
    totalRows: rows.length,
    validRows: importedRows.length,
    invalidRows: rows.length - importedRows.length,
    schemaMatchedRows,
    enabledRows: enabledRows.length,
    disabledRows: importedRows.length - enabledRows.length,
    offlineRows,
    abnormalRows,
    sampleRows: importedRows.slice(0, 8),
    importedRoadsideRows: importedRows,
    importDate,
    confirmed: false,
  };
}

function createImportTask(type, file) {
  const isSites = type === "sites";
  const isRoadside = type === "roadsideStatus";
  const sampleRows = isSites ? sites.slice(0, 8) : isRoadside ? roadsideStatusState.currentRows.slice(0, 8) : allDevices().slice(0, 8);
  const enabledRows = roadsideStatusState.currentRows.filter(isRoadsideEnabled);
  const abnormalRows = enabledRows.filter(isRoadsideAbnormal).length;
  const offlineRows = enabledRows.filter((device) => device.status === "离线").length;
  const invalidRows = isSites ? 1 : 0;
  const matchedRows = allDevices().length;
  return {
    id: `${type}-${Date.now()}`,
    type,
    fileName: file?.name || expectedFileName(type),
    fileSize: file?.size || 0,
    templateMatched: file ? fileLooksLikeTemplate(type, file.name) : true,
    totalRows: isSites ? data.stats.siteTotal + invalidRows : isRoadside ? roadsideStatusState.currentRows.length : data.stats.deviceRows + data.stats.unmatchedRows,
    validRows: isSites ? data.stats.siteTotal : isRoadside ? roadsideStatusState.currentRows.length : data.stats.deviceRows,
    invalidRows,
    duplicateNodeIdCount: isSites ? 0 : null,
    coordinateErrorCount: isSites ? 1 : null,
    dictionaryErrorCount: isSites ? 3 : null,
    sheet1TotalRows: isSites ? null : data.stats.deviceRows,
    matchedRows: isSites ? null : matchedRows,
    unmatchedSheetRows: isSites ? null : data.stats.unmatchedRows,
    unpackingIssueRows: isSites ? null : 42,
    quantityErrorRows: isSites ? null : 18,
    enabledRows: isRoadside ? enabledRows.length : null,
    disabledRows: isRoadside ? roadsideStatusState.currentRows.length - enabledRows.length : null,
    offlineRows: isRoadside ? offlineRows : null,
    abnormalRows: isRoadside ? abnormalRows : null,
    importedRoadsideRows: isRoadside ? roadsideStatusState.currentRows : null,
    importDate: isRoadside ? roadsideStatusState.currentDate : null,
    sampleRows,
    confirmed: false,
  };
}

function renderImportError(type, message) {
  const config = importTemplates[type];
  const panel = importPanelElement();
  panel.innerHTML = `
    <div class="import-card">
      <div>
        <span>导入失败</span>
        <strong>${config.title}</strong>
        <p>${message}</p>
      </div>
      <div class="import-footer">
        <button class="primary-btn" data-choose-import="${type}">重新选择文件</button>
        <button class="ghost-btn" data-simulate-import="${type}">使用模板样例模拟导入</button>
      </div>
    </div>
  `;
  panel.classList.add("open");
}

function renderImportTask(task, stage = "preview") {
  const config = importTemplates[task.type];
  const fields = importFieldMaps[task.type];
  const isSites = task.type === "sites";
  const isRoadside = task.type === "roadsideStatus";
  const stages = ["上传文件", "字段识别", "预校验", "预览", "确认导入", "导入报告"];
  const activeIndex = stage === "report" ? 5 : stage === "confirm" ? 4 : 3;
  const panel = importPanelElement();
  panel.innerHTML = `
    <div class="import-card import-card-full">
      <div class="import-title-row">
        <div>
          <span>导入任务 ${task.id}</span>
          <strong>${config.title}</strong>
          <p>${task.fileName} ${task.fileSize ? ` / ${(task.fileSize / 1024).toFixed(1)} KB` : " / 模板样例"}</p>
        </div>
        <span class="status-pill ${task.templateMatched ? "done" : "warn"}">${task.templateMatched ? "模板匹配" : "模板疑似不匹配"}</span>
      </div>
      <div class="import-steps">
        ${stages.map((name, index) => `<span class="${index <= activeIndex ? "active" : ""}">${name}</span>`).join("")}
      </div>
      <div class="import-summary-grid">
        ${
          isSites
            ? `
              <div><span>总行数</span><strong>${formatNumber(task.totalRows)}</strong></div>
              <div><span>有效行</span><strong>${formatNumber(task.validRows)}</strong></div>
              <div><span>失败行</span><strong>${formatNumber(task.invalidRows)}</strong></div>
              <div><span>坐标异常</span><strong>${formatNumber(task.coordinateErrorCount)}</strong></div>
              <div><span>字典异常</span><strong>${formatNumber(task.dictionaryErrorCount)}</strong></div>
              <div><span>重复 NodeID</span><strong>${formatNumber(task.duplicateNodeIdCount)}</strong></div>
            `
            : isRoadside
              ? `
              <div><span>总行数</span><strong>${formatNumber(task.totalRows)}</strong></div>
              <div><span>当前表行数</span><strong>${formatNumber(task.validRows)}</strong></div>
              <div><span>已启用</span><strong>${formatNumber(task.enabledRows)}</strong></div>
              <div><span>未启用</span><strong>${formatNumber(task.disabledRows)}</strong></div>
              <div><span>离线</span><strong>${formatNumber(task.offlineRows)}</strong></div>
              <div><span>异常</span><strong>${formatNumber(task.abnormalRows)}</strong></div>
            `
            : `
              <div><span>Sheet1 行数</span><strong>${formatNumber(task.sheet1TotalRows)}</strong></div>
              <div><span>匹配成功</span><strong>${formatNumber(task.matchedRows)}</strong></div>
              <div><span>未匹配池</span><strong>${formatNumber(task.unmatchedSheetRows)}</strong></div>
              <div><span>开箱问题</span><strong>${formatNumber(task.unpackingIssueRows)}</strong></div>
              <div><span>数量异常</span><strong>${formatNumber(task.quantityErrorRows)}</strong></div>
              <div><span>治理目标</span><strong>点位/仓库/机房/服务项</strong></div>
            `
        }
      </div>
      <div class="import-columns">
        <section>
          <h3>字段识别</h3>
          <div class="field-map">
            ${fields
              .map(
                ([source, target, note]) => `
                  <div>
                    <span>${source}</span>
                    <strong>${target}</strong>
                    <small>${note}</small>
                  </div>
                `,
              )
              .join("")}
          </div>
        </section>
        <section>
          <h3>预览数据</h3>
          <div class="preview-list">
            ${task.sampleRows
              .map((row) =>
                isSites
                  ? `<div><strong>${row.nodeId}</strong><span>${row.name} / ${row.district} / ${row.type} / ${perceptionType(row)}</span></div>`
                  : isRoadside
                    ? `<div><strong>${row.deviceId}</strong><span>${row.deviceTypeName || "-"} / ${row.devicePosition || "-"} / ${row.status || "-"} / ${row.enabledStatus || "-"}</span></div>`
                  : `<div><strong>${row.nodeId || "未匹配"}</strong><span>${row.name || "-"} / ${row.supplier || "-"} / ${row.contract || "-"}</span></div>`,
              )
              .join("")}
          </div>
        </section>
      </div>
      <div class="import-footer">
        ${
          stage === "report"
            ? `<button class="primary-btn" data-import-finish="${task.type}">${isSites ? "查看点位管理" : isRoadside ? "查看运维管理" : "查看设备管理"}</button>
               <button class="ghost-btn" data-choose-import="${task.type}">重新导入</button>`
            : `<button class="primary-btn" data-confirm-import="${task.type}">确认导入</button>
               <button class="ghost-btn" data-choose-import="${task.type}">更换文件</button>`
        }
      </div>
      ${stage === "report" ? renderImportReport(task) : ""}
    </div>
  `;
  panel.classList.add("open");
}

function renderImportReport(task) {
  const isSites = task.type === "sites";
  const isRoadside = task.type === "roadsideStatus";
  return `
    <div class="import-report">
      <h3>导入报告</h3>
      <p>${isSites ? "点位库已更新，地图坐标以 GCJ-02 为准，CGCS2000 已进入审计字段。" : isRoadside ? "原当前路侧设备状态表已进入历史档案库，新导入数据已成为当前表。" : "设备安装台账已更新，未匹配记录已进入治理池。"}</p>
      <ul>
        <li>${isSites ? `成功导入 ${formatNumber(task.validRows)} 个点位` : isRoadside ? `当前路侧设备状态表 ${formatNumber(task.validRows)} 条` : `成功匹配 ${formatNumber(task.matchedRows)} 条设备记录`}</li>
        <li>${isSites ? `坐标异常 ${formatNumber(task.coordinateErrorCount)} 条，进入坐标治理` : isRoadside ? `已启用 ${formatNumber(task.enabledRows)} 条，未启用 ${formatNumber(task.disabledRows)} 条` : `未匹配 ${formatNumber(task.unmatchedSheetRows)} 条，进入未匹配池`}</li>
        <li>${isSites ? `字典异常 ${formatNumber(task.dictionaryErrorCount)} 条，需人工复核` : isRoadside ? `离线 ${formatNumber(task.offlineRows)} 条，异常 ${formatNumber(task.abnormalRows)} 条` : `开箱验收问题 ${formatNumber(task.unpackingIssueRows)} 条，生成异常标识`}</li>
      </ul>
    </div>
  `;
}

function pushImportHistory(task) {
  const exists = state.importHistory.some((item) => item.id === task.id);
  if (exists) return;
  state.importHistory.unshift({
    id: task.id,
    type: task.type,
    fileName: task.fileName,
    validRows: task.validRows,
    invalidRows: task.invalidRows,
    matchedRows: task.matchedRows,
    unmatchedRows: task.unmatchedSheetRows,
    status: "已导入",
    createdAt: new Date().toLocaleString("zh-CN", { hour12: false }),
  });
  state.importHistory = state.importHistory.slice(0, 8);
}

async function startImport(type, file = null) {
  state.importType = type;
  try {
    state.importTask = file ? await createImportTaskFromWorkbook(type, file) : createImportTask(type, file);
    renderImportTask(state.importTask);
  } catch (error) {
    console.error("Import failed.", error);
    renderImportError(type, error.message || "文件解析失败，请检查 Excel 格式。");
  }
}

function applySiteImport(task) {
  if (!task.importedSites) return;
  sites.splice(0, sites.length, ...task.importedSites);
  data.stats.siteTotal = task.validRows;
  data.stats.prototypeSites = task.validRows;
}

function applyDeviceImport(task) {
  if (!task.importedDevices) return;
  const siteMap = new Map(sites.map((site) => [site.nodeId, site]));
  sites.forEach((site) => {
    site.devices = [];
    site.deviceRecordCount = 0;
    site.deviceTypeCount = 0;
    site.installedQty = 0;
  });
  task.importedDevices.forEach((device) => {
    const site = siteMap.get(device.nodeId);
    if (!site) return;
    site.devices.push(device);
  });
  sites.forEach((site) => {
    const typeCount = new Set(site.devices.map(deviceType)).size;
    site.deviceRecordCount = site.devices.length;
    site.deviceTypeCount = typeCount;
    site.installedQty = site.devices.reduce((sum, device) => sum + (device.installedQty || 0), 0);
    if (site.devices.some((device) => device.issue)) site.issueCount = Math.max(site.issueCount || 0, 1);
    site.archiveCompleteness = Math.min(95, 35 + site.devices.length * 5);
  });
  data.stats.deviceRows = task.matchedRows;
  data.stats.unmatchedRows = task.unmatchedSheetRows;
}

async function applyRoadsideStatusImport(task) {
  if (!task.importedRoadsideRows) return;
  const today = new Date().toISOString().slice(0, 10);
  if (roadsideStatusState.currentDate && roadsideStatusState.currentDate !== today && roadsideStatusState.currentRows.length) {
    roadsideStatusState.archives = roadsideStatusState.archives.filter((item) => item.importDate !== roadsideStatusState.currentDate);
    roadsideStatusState.archives.unshift({
      importDate: roadsideStatusState.currentDate,
      rows: roadsideStatusState.currentRows,
      archivedAt: new Date().toLocaleString("zh-CN", { hour12: false }),
    });
  }
  roadsideStatusState.archives = roadsideStatusState.archives.filter((item) => item.importDate !== today);
  roadsideStatusState.currentDate = today;
  roadsideStatusState.currentRows = task.importedRoadsideRows;
  state.opsDate = roadsideStatusState.currentDate;
  state.opsCalendarCursor = state.opsDate.slice(0, 7);
  const backendSaved = await saveRoadsideStatusState();
  if (!backendSaved) alert("本地文件持久化服务未连接。本次导入只保存到了浏览器缓存，请通过 http://127.0.0.1:4173 访问并确认本地服务已启动。");
}

async function confirmImport(type) {
  if (!state.importTask || state.importTask.type !== type) return;
  state.importTask.confirmed = true;
  if (type === "sites") {
    applySiteImport(state.importTask);
    renderFilters();
  }
  if (type === "devices") {
    applyDeviceImport(state.importTask);
    $("#unmatchedPanel").classList.add("open");
  }
  if (type === "roadsideStatus") {
    await applyRoadsideStatusImport(state.importTask);
  }
  pushImportHistory(state.importTask);
  renderMetrics();
  renderStatusProgress();
  renderSites();
  renderDevices();
  renderWarehouse();
  renderImportCenter();
  renderCoordinateIssues();
  renderOps();
  if (type === "roadsideStatus" && state.panel === "ops") {
    closeImportPanel($("#opsImportPanel"));
    state.importTask = null;
    state.importType = null;
    return;
  }
  renderImportTask(state.importTask, "report");
}

function mapBounds(items) {
  const lngs = items.map((item) => item.lngGcj);
  const lats = items.map((item) => item.latGcj);
  return {
    minLng: Math.min(...lngs),
    maxLng: Math.max(...lngs),
    minLat: Math.min(...lats),
    maxLat: Math.max(...lats),
  };
}

function markerPosition(site, bounds) {
  const lngRange = bounds.maxLng - bounds.minLng || 1;
  const latRange = bounds.maxLat - bounds.minLat || 1;
  return {
    left: 7 + ((site.lngGcj - bounds.minLng) / lngRange) * 86,
    top: 88 - ((site.latGcj - bounds.minLat) / latRange) * 76,
  };
}

function renderFallbackMap(container, items, options = {}) {
  const bounds = mapBounds(items.length ? items : sites);
  container.innerHTML = "";
  container.classList.remove("amap-live");
  items.forEach((site) => {
    const pos = markerPosition(site, bounds);
    const isCurrent = options.currentNodeId && site.nodeId === String(options.currentNodeId);
    const marker = document.createElement("div");
    marker.className = `marker ${districtClass(site)} ${isCurrent ? "current" : "neighbor"} ${site.issueCount ? "has-issue" : ""}`;
    marker.style.left = `${pos.left}%`;
    marker.style.top = `${pos.top}%`;
    marker.title = `${isCurrent ? "当前点位：" : "周边点位："}${site.nodeId} ${site.name}`;
    marker.innerHTML = `<button aria-label="${site.name}"></button><span>${isCurrent ? "当前" : ""}</span>`;
    marker.querySelector("button").addEventListener("click", () => openSite(site.nodeId));
    container.appendChild(marker);
  });
  if (options.label) {
    const label = document.createElement("div");
    label.className = "map-label";
    label.textContent = options.label;
    container.appendChild(label);
  }
}

function renderAmap(container, items, options = {}) {
  if (!window.AMap || !items.length) return false;

  const previousMap = amapInstances.get(container);
  if (previousMap) {
    previousMap.destroy();
    amapInstances.delete(container);
  }

  container.innerHTML = "";
  container.classList.add("amap-live");

  const centerSite = options.centerSite || items[0];
  const center = [centerSite.lngGcj, centerSite.latGcj];
  const map = new window.AMap.Map(container, {
    zoom: options.zoom || 13,
    center,
    viewMode: "2D",
    mapStyle: "amap://styles/normal",
  });

  amapInstances.set(container, map);
  map.setCenter(center);
  map.setZoom(options.zoom || (items.length <= 6 ? 14 : 11));

  const markers = items.map((site) => {
    const isCurrent = options.currentNodeId && site.nodeId === String(options.currentNodeId);
    const marker = new window.AMap.Marker({
      position: [site.lngGcj, site.latGcj],
      title: `${isCurrent ? "当前点位：" : "周边点位："}${site.nodeId} ${site.name}`,
      offset: new window.AMap.Pixel(isCurrent ? -13 : -7, isCurrent ? -13 : -7),
      zIndex: isCurrent ? 120 : 100,
      content: `<button class="amap-site-marker ${districtClass(site)} ${isCurrent ? "current" : "neighbor"} ${site.issueCount ? "has-issue" : ""}" title="${site.name}" data-node="${site.nodeId}">${isCurrent ? "当前" : ""}</button>`,
    });
    marker.on("click", () => openSite(site.nodeId));
    return marker;
  });

  map.add(markers);
  if (markers.length > 1) {
    map.setFitView(markers, false, [36, 36, 36, 36], options.maxZoom || 15);
  }
  return true;
}

function renderMap(container, items, options = {}) {
  try {
    if (renderAmap(container, items, options)) return;
  } catch (error) {
    console.warn("AMap render failed, fallback map enabled.", error);
  }
  renderFallbackMap(container, items, options);
}

function renderOverviewMap() {
  const overviewMap = $("#overviewMap");
  if (!overviewMap) return;
  renderMap(overviewMap, sites, {
    label: `全量点位预览：${formatNumber(sites.length)} 个点位`,
    zoom: 10,
    maxZoom: 12,
  });
}

function renderFilters() {
  renderSiteSummaryCards();
  renderSiteColumnFilters();
}

function renderSites() {
  const rows = filteredSites();
  $("#siteRows").innerHTML = rows
    .map(
      (site, index) => `
        <tr data-open-site="${site.nodeId}" tabindex="0" aria-label="打开 ${site.name} 编辑页">
          <td class="site-index-col">${index + 1}</td>
          <td><strong>${site.nodeId}</strong></td>
          <td><strong>${site.name}</strong><br /><small>${site.englishName || ""}</small></td>
          <td>${districtBadge(site)}</td>
          <td>${site.type}</td>
          <td>${site.vendor || "-"}</td>
          <td>${perceptionType(site)}</td>
          <td>${site.adaptive ? "是" : "否"}</td>
          <td>${isVariableLane(site) ? `是 / ${site.variableLaneCount || "-"}` : "否"}</td>
          <td>${site.lngGcj.toFixed(6)}<br /><small>${site.latGcj.toFixed(6)}</small></td>
          <td>${deviceTypeCount(site)} 类 / ${deviceRecordCount(site)} 条 / ${site.installedQty || 0} 已安装</td>
          <td><span class="issue-pill ${site.issueCount ? "warn" : ""}">${issueText(site)}</span></td>
          <td>
            <button class="link-btn" data-open-site="${site.nodeId}">编辑</button>
            <button class="link-btn danger" data-delete-site="${site.nodeId}">删除</button>
          </td>
        </tr>
      `,
    )
    .join("");
  renderSiteColumnHandles();
  renderSiteColumnFilters();
  applySiteViewLayoutSettled();

  $("#siteCards").innerHTML = rows
    .map(
      (site) => `
        <article class="site-card">
          <div class="site-card-head">
            <h3 class="site-card-title"><span class="site-card-node">${site.nodeId}</span><span class="site-card-name">${site.name}</span></h3>
            ${districtBadge(site)}
          </div>
          <p class="site-card-meta">
            <span>${site.type}</span>
            <span>${site.vendor || "-"}</span>
            <span>${perceptionType(site)}</span>
            <span>${site.adaptive ? "自适应" : "非自适应"}</span>
            <span>${isVariableLane(site) ? "可变车道" : "非可变车道"}</span>
          </p>
          <div class="card-foot">
            <span>${deviceTypeCount(site)} 类设备 · ${deviceRecordCount(site)} 条记录 · 异常 ${issueText(site)} · 档案 ${site.archiveCompleteness}%</span>
            <button class="link-btn" data-open-site="${site.nodeId}">编辑</button>
          </div>
        </article>
      `,
    )
    .join("");

  if (state.view === "map") {
    renderMap($("#siteMap"), rows);
  }
  $("#mapList").innerHTML = rows
    .map(
      (site) => `
        <button data-open-site="${site.nodeId}">
          <strong>${site.name}</strong><br />
          <span>${site.nodeId} / ${site.district} / ${perceptionType(site)} / 异常 ${issueText(site)}</span>
        </button>
      `,
    )
    .join("");
  bindSiteOpenButtons();
}

function bindSiteOpenButtons() {
  $$("[data-open-site]").forEach((element) => {
    element.onclick = (event) => {
      if (event.target.closest("[data-delete-site]")) return;
      event.preventDefault();
      event.stopPropagation();
      openSite(element.dataset.openSite);
    };
    element.onkeydown = (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        event.stopPropagation();
        openSite(element.dataset.openSite);
      }
    };
  });
}

function renderDevices() {
  const q = state.query.trim().toLowerCase();
  const rows = allDevices().filter((device) => {
    const text = `${device.nodeId} ${device.siteName} ${device.position} ${device.name} ${device.supplier} ${device.contract}`.toLowerCase();
    return !q || text.includes(q);
  });
  $("#deviceRows").innerHTML = rows
    .map(
      (device) => `
        <tr>
          <td>${device.siteName}<br /><small>${device.nodeId}</small></td>
          <td><span class="device-type">${device.type}</span></td>
          <td>${device.position || "-"}</td>
          <td>${device.name || "-"}<br /><small>${device.model || ""}</small></td>
          <td>${device.supplier || "-"}</td>
          <td>${device.contract || "-"}</td>
          <td>${device.deliveredQty || 0}</td>
          <td>${device.inboundQty || 0}</td>
          <td>${device.installedQty || 0}</td>
        </tr>
      `,
    )
    .join("");
  renderUnmatchedRows();
}

function renderContracts() {
  renderContractSystemViews();
}

function chainText(chain) {
  return (chain || []).filter(Boolean).join(" → ") || "-";
}

function contractDirectionText(direction) {
  if (direction === "front_sales") return "前向";
  if (direction === "back_procurement") return "后向";
  return "待判定";
}

function contractDirectionPillClass(direction) {
  return direction === "front_sales" ? "done" : direction === "back_procurement" ? "warn" : "";
}

function contractCandidateSearchText(candidate) {
  return [
    candidate.frontContractName,
    candidate.backContractName,
    ...(candidate.partyOverlap || []),
    ...(candidate.keywordOverlap || []),
    ...(candidate.taxRateOverlap || []),
    ...(candidate.reasons || []),
  ]
    .join(" ")
    .toLowerCase();
}

function filterContractCandidates(candidates) {
  const filters = state.contractReview.filters || {};
  const query = (filters.query || "").trim().toLowerCase();
  return candidates.filter((candidate) => {
    const decisionStatus = contractCandidateDecision(candidate.id).status || "candidate";
    if (filters.decision && filters.decision !== "all" && decisionStatus !== filters.decision) return false;
    if (filters.confidence && filters.confidence !== "all" && candidate.confidence !== filters.confidence) return false;
    if (filters.taxRate && filters.taxRate !== "all") {
      const rates = candidate.taxRateOverlap || [];
      if (filters.taxRate === "none") {
        if (rates.length) return false;
      } else if (!rates.includes(filters.taxRate)) {
        return false;
      }
    }
    if (query && !contractCandidateSearchText(candidate).includes(query)) return false;
    return true;
  });
}

function confirmedContractCandidates() {
  const cr = state.contractRelationships;
  return (cr.frontBackRelationshipCandidates || []).filter((candidate) => {
    if (contractCandidateDecision(candidate.id).status !== "confirmed") return false;
    if (cr.selectedMacroFlowId && !(candidate.sharedMacroFlowIds || []).includes(cr.selectedMacroFlowId)) return false;
    return true;
  });
}

function candidatePrimaryMacroFlow(candidate, macroById) {
  return macroById.get(candidate?.sharedMacroFlowIds?.[0]) || {};
}

function candidatePrimaryItemName(candidate) {
  return (candidate?.keywordOverlap || []).slice(0, 4).join("、") || "待抽取合同附件设备/服务项";
}

function buildDeviceCashflowPreviewUnits(macroById, limit = 40) {
  return confirmedContractCandidates().slice(0, limit).map((candidate, index) => {
    const flow = candidatePrimaryMacroFlow(candidate, macroById);
    const taxRates = candidate.taxRateOverlap || [];
    const taxRate = taxRates[0] || "待确认";
    const itemName = candidatePrimaryItemName(candidate);
    return {
      id: `preview-cashflow-${index + 1}`,
      candidateId: candidate.id,
      macroFlowName: flow.packageName || `共同宏观流 ${(candidate.sharedMacroFlowIds || []).length} 条`,
      frontContractName: candidate.frontContractName || "-",
      backContractName: candidate.backContractName || "-",
      itemName,
      nodeId: "待绑定",
      siteName: "待绑定点位",
      quantity: "待拆分",
      frontTaxRate: taxRate,
      backTaxRate: taxRate,
      ownerAuditStatus: "待拆分后申报",
      auditConfirmedAmount: "0",
      upstreamStatus: "未释放",
      downstreamStatus: "锁定",
      risk: "未形成设备级清单、NodeID 和业主审计结果",
    };
  });
}

function countConfirmedDevicePreviewUnits() {
  return confirmedContractCandidates().length;
}

function renderContractCandidateEvidence(candidate, macroById) {
  const panel = $("#contractCandidateEvidencePanel");
  if (!panel) return;
  if (!candidate) {
    panel.innerHTML = `
      <div class="contract-evidence-empty">
        <strong>未选择候选关系</strong>
        <span>点击候选关系中的“证据”按钮，查看主体、关键词、税率和宏观流命中依据。</span>
      </div>`;
    return;
  }
  const sharedFlows = (candidate.sharedMacroFlowIds || []).map((id) => macroById.get(id)?.packageName || id).slice(0, 6);
  const decision = contractCandidateDecision(candidate.id).status || "candidate";
  const decisionText = decision === "confirmed" ? "已确认" : decision === "rejected" ? "已否决" : "待确认";
  panel.innerHTML = `
    <div class="contract-evidence-title">
      <div>
        <strong>${escapeHtml(candidate.frontContractName || "-")}</strong>
        <span>经天安候选对应</span>
        <strong>${escapeHtml(candidate.backContractName || "-")}</strong>
      </div>
      <span class="status-pill ${decision === "confirmed" ? "done" : decision === "rejected" ? "error" : "warn"}">${escapeHtml(decisionText)}</span>
    </div>
    <div class="contract-evidence-grid">
      <span><b>共同宏观流</b>${escapeHtml(sharedFlows.join("；") || "未命中共同宏观流")}</span>
      <span><b>主体重合</b>${escapeHtml((candidate.partyOverlap || []).join("、") || "无")}</span>
      <span><b>税率重合</b>${escapeHtml((candidate.taxRateOverlap || []).join(" / ") || "待确认")}</span>
      <span><b>设备/服务关键词</b>${escapeHtml((candidate.keywordOverlap || []).slice(0, 12).join("、") || "待抽取")}</span>
      <span class="wide"><b>匹配依据</b>${escapeHtml((candidate.reasons || []).join("；") || "-")}</span>
      <span class="wide"><b>资金闸门</b>${escapeHtml(candidate.importantNote || "候选关系必须人工确认后才能进入设备级资金流计算。")}</span>
    </div>`;
}

function renderContractSystemViews() {
  const cr = state.contractRelationships;
  const status = $("#contractRelationshipStatus");
  const metrics = $("#contractSystemMetrics");
  const macroRows = $("#macroFlowRows");
  const docRows = $("#contractDocMatchRows");
  const candidateRows = $("#frontBackCandidateRows");
  const deviceCashflowRows = $("#deviceCashflowRows");
  const ownerAuditRows = $("#ownerAuditRows");
  const paymentStageRows = $("#paymentStageRows");
  if (!status && !metrics && !macroRows && !docRows && !candidateRows && !deviceCashflowRows && !ownerAuditRows && !paymentStageRows) return;

  if (status) {
    if (cr.loading) {
      status.textContent = "正在读取 Excel 合同流尾部和 Word 合同匹配结果...";
    } else if (cr.error) {
      status.textContent = `合同关系重构数据读取失败：${cr.error}`;
    } else {
      status.textContent = "已按“Excel 宏观合同流 + Word 前后向合同文本”重构候选关系；设备清单 sheet 不作为合同边界。";
    }
  }

  const summary = cr.summary || {};
  const reviewStats = contractReviewStats();
  const macroById = new Map(cr.macroFlows.map((flow) => [flow.id, flow]));
  if (metrics) {
    const confirmedPreviewUnits = countConfirmedDevicePreviewUnits();
    const cards = [
      ["宏观合同流", summary.macroFlowCount ?? cr.macroFlows.length, "Excel 合同流尾部"],
      ["Word 合同", summary.contractDocumentCount ?? cr.contracts.length, `前向 ${summary.frontContractCount || 0} / 后向 ${summary.backContractCount || 0}`],
      ["前后向候选", summary.frontBackCandidateCount ?? cr.frontBackRelationshipCandidates.length, "多对多，不是 sheet 关系"],
      ["待确认", reviewStats.pending, "不进入资金计算"],
      ["确认/否决", `${reviewStats.confirmed} / ${reviewStats.rejected}`, `${confirmedPreviewUnits} 条待设备级拆分`],
      ["税率线索", Object.entries(summary.taxRateCounts || {}).map(([k, v]) => `${k}:${v}`).join(" / ") || "待解析", "一级匹配字段"],
      ["审计付款闸门", cr.paymentStageTemplates?.length || 6, "预付/到货/安装/验收/审计/质保"],
    ];
    metrics.innerHTML = cards
      .map(
        ([label, value, note]) => `
          <article class="metric contract-system-metric">
            <span>${escapeHtml(label)}</span>
            <strong>${escapeHtml(formatMetricValue(value))}</strong>
            <small>${escapeHtml(note)}</small>
          </article>
        `,
      )
      .join("");
  }

  if (macroRows) {
    const flows = [...cr.macroFlows].sort((a, b) => (a.rowNumber || 0) - (b.rowNumber || 0));
    macroRows.innerHTML = flows.length
      ? flows
          .map(
            (flow) => `
              <tr class="${flow.id === cr.selectedMacroFlowId ? "selected-row" : ""}" data-select-macro-flow="${escapeHtml(flow.id)}">
                <td>${escapeHtml(flow.rowNumber || "-")}</td>
                <td><strong>${escapeHtml(flow.packageName || "-")}</strong><br /><small>金额 ${escapeHtml(formatMoney(flow.packageAmount || 0))}</small></td>
                <td>${escapeHtml(chainText(flow.chain))}</td>
                <td>${escapeHtml(formatMoney(flow.tiananAmount || flow.packageAmount || 0))}</td>
                <td><span class="status-pill ${flow.flowType === "mobile_to_tianan" ? "warn" : "done"}">${escapeHtml(flow.flowType === "mobile_to_tianan" ? "经天安回流/拆分" : flow.flowType === "tianan_direct" ? "天安直接拆分" : "其他")}</span></td>
              </tr>
            `,
          )
          .join("")
      : `<tr><td colspan="5">${cr.loading ? "正在读取宏观合同流..." : "暂无宏观合同流记录"}</td></tr>`;
  }

  if (docRows) {
    const matchByContract = new Map(cr.contractToMacroFlowMatches.map((item) => [item.contractId, item]));
    const docs = [...cr.contracts]
      .filter((doc) => {
        if (!cr.selectedMacroFlowId) return true;
        return (matchByContract.get(doc.id)?.matches || []).some((match) => match.macroFlowId === cr.selectedMacroFlowId);
      })
      .sort((a, b) => String(a.direction || "").localeCompare(String(b.direction || "")) || String(a.fileName || "").localeCompare(String(b.fileName || ""), "zh-CN"))
      .slice(0, 120);
    docRows.innerHTML = docs.length
      ? docs
          .map((doc) => {
            const match = matchByContract.get(doc.id);
            const top = match?.matches?.[0];
            return `
              <tr>
                <td><span class="status-pill ${contractDirectionPillClass(doc.direction)}">${escapeHtml(contractDirectionText(doc.direction))}</span></td>
                <td><strong>${escapeHtml(doc.fileName || "-")}</strong><br /><small>文本 ${escapeHtml(formatNumber(doc.textLength || 0))} 字</small></td>
                <td>${escapeHtml(chainText(doc.parties))}</td>
                <td>${escapeHtml((doc.taxRates || []).join(" / ") || "待确认")}</td>
                <td>${top ? `${escapeHtml(top.macroFlowPackage)}<br /><small>${escapeHtml((top.reasons || []).join("；"))}</small>` : "未匹配"}</td>
              </tr>
            `;
          })
          .join("")
      : `<tr><td colspan="5">${cr.loading ? "正在读取 Word 合同..." : "暂无合同匹配记录"}</td></tr>`;
  }

  let visibleCandidates = [];
  if (candidateRows) {
    const baseCandidates = [...cr.frontBackRelationshipCandidates]
      .filter((candidate) => !cr.selectedMacroFlowId || (candidate.sharedMacroFlowIds || []).includes(cr.selectedMacroFlowId));
    visibleCandidates = filterContractCandidates(baseCandidates)
      .sort((a, b) => {
        const statusRank = { candidate: 0, confirmed: 1, rejected: 2 };
        const aStatus = contractCandidateDecision(a.id).status || "candidate";
        const bStatus = contractCandidateDecision(b.id).status || "candidate";
        return (statusRank[aStatus] ?? 0) - (statusRank[bStatus] ?? 0) || (b.score || 0) - (a.score || 0);
      })
      .slice(0, 160);
    if (state.contractReview.selectedCandidateId && !visibleCandidates.some((candidate) => candidate.id === state.contractReview.selectedCandidateId)) {
      state.contractReview.selectedCandidateId = "";
    }
    if (!state.contractReview.selectedCandidateId && visibleCandidates.length) {
      state.contractReview.selectedCandidateId = visibleCandidates[0].id;
    }
    const selectedCandidate = visibleCandidates.find((candidate) => candidate.id === state.contractReview.selectedCandidateId);
    renderContractCandidateEvidence(selectedCandidate, macroById);
    candidateRows.innerHTML = visibleCandidates.length
      ? visibleCandidates
          .map((candidate) => {
            const decision = contractCandidateDecision(candidate.id);
            const decisionStatus = decision.status || "candidate";
            const confidenceClass = candidate.confidence === "high" ? "done" : candidate.confidence === "medium" ? "warn" : "error";
            const decisionClass = decisionStatus === "confirmed" ? "done" : decisionStatus === "rejected" ? "error" : "warn";
            const decisionText = decisionStatus === "confirmed" ? "已确认" : decisionStatus === "rejected" ? "已否决" : "待确认";
            const flowNames = (candidate.sharedMacroFlowIds || [])
              .map((id) => macroById.get(id)?.packageName)
              .filter(Boolean)
              .slice(0, 2);
            return `
              <tr class="${candidate.id === state.contractReview.selectedCandidateId ? "selected-row" : ""}">
                <td><span class="status-pill ${confidenceClass}">${escapeHtml(candidate.confidence || "candidate")}</span><br /><small>score ${escapeHtml(candidate.score ?? "-")}</small></td>
                <td><strong>${escapeHtml(candidate.frontContractName || "-")}</strong></td>
                <td><strong>${escapeHtml(candidate.backContractName || "-")}</strong></td>
                <td>${escapeHtml(flowNames.length ? flowNames.join("；") : `${(candidate.sharedMacroFlowIds || []).length} 条共同宏观流`)}<br /><small>${escapeHtml((candidate.taxRateOverlap || []).join(" / ") || "税率待确认")}</small></td>
                <td>${escapeHtml((candidate.reasons || []).join("；") || "-")}<br /><small>${escapeHtml((candidate.keywordOverlap || []).slice(0, 8).join("、"))}</small></td>
                <td>
                  <button class="link-btn evidence" data-contract-candidate-evidence="${escapeHtml(candidate.id)}">证据</button>
                  <span class="status-pill ${decisionClass}">${escapeHtml(decisionText)}</span>
                  <div class="contract-review-actions">
                    <button class="link-btn" data-contract-candidate-decision="confirmed" data-candidate-id="${escapeHtml(candidate.id)}">确认</button>
                    <button class="link-btn danger" data-contract-candidate-decision="rejected" data-candidate-id="${escapeHtml(candidate.id)}">否决</button>
                    <button class="link-btn muted" data-contract-candidate-decision="candidate" data-candidate-id="${escapeHtml(candidate.id)}">恢复</button>
                  </div>
                  ${decision.reviewedAt ? `<small>${escapeHtml(new Date(decision.reviewedAt).toLocaleString("zh-CN"))}</small>` : ""}
                </td>
              </tr>
            `;
          })
          .join("")
      : `<tr><td colspan="6">${cr.loading ? "正在生成前后向候选关系..." : "当前筛选条件下暂无前后向候选关系"}</td></tr>`;
    if (!visibleCandidates.length) renderContractCandidateEvidence(null, macroById);
  }

  if (deviceCashflowRows) {
    const previewUnits = buildDeviceCashflowPreviewUnits(macroById, 40);
    const schema = cr.deviceCashflowSchema || {};
    deviceCashflowRows.innerHTML = previewUnits.length
      ? previewUnits
          .map((candidate) => {
            return `
              <tr>
                <td>${escapeHtml(candidate.macroFlowName)}</td>
                <td>${escapeHtml(candidate.frontContractName)}</td>
                <td>${escapeHtml(candidate.backContractName)}</td>
                <td><span class="status-pill warn">待拆分</span><br /><small>${escapeHtml(candidate.itemName)}</small></td>
                <td><small>前向 ${escapeHtml(candidate.frontTaxRate)}</small><br /><small>后向 ${escapeHtml(candidate.backTaxRate)}</small></td>
                <td><span class="status-pill warn">${escapeHtml(candidate.nodeId)}</span><br /><small>${escapeHtml(candidate.siteName)}</small></td>
                <td><span class="status-pill warn">${escapeHtml(candidate.ownerAuditStatus)}</span><br /><small>审计确认金额 ${escapeHtml(candidate.auditConfirmedAmount)}</small></td>
                <td><span class="status-pill warn">${escapeHtml(candidate.downstreamStatus)}</span><br /><small>${escapeHtml(candidate.risk)}</small></td>
              </tr>
            `;
          })
          .join("")
      : `<tr><td colspan="8">尚无已确认的前后向合同关系。请先在上方候选关系中确认，再进入设备级资金流拆分。最小单元：${escapeHtml(schema.minimumUnit || "前向合同明细 → 后向合同明细 → 清单池行 → NodeID/点位 → 业主审计 → 阶段付款")}。</td></tr>`;
  }

  if (ownerAuditRows) {
    const previewUnits = buildDeviceCashflowPreviewUnits(macroById, 30);
    ownerAuditRows.innerHTML = previewUnits.length
      ? previewUnits
          .map((unit, index) => `
            <tr>
              <td><strong>设备级单元 ${index + 1}</strong><br /><small>${escapeHtml(unit.itemName)}</small></td>
              <td><span class="status-pill warn">${escapeHtml(unit.nodeId)}</span><br /><small>${escapeHtml(unit.siteName)}</small></td>
              <td><span class="status-pill warn">${escapeHtml(unit.ownerAuditStatus)}</span></td>
              <td>申报/确认金额待业主审计批次写入</td>
              <td>未通过设备级审计前，不得作为上游可回款依据</td>
              <td>未形成设备级审计结果和上游资金池前，后向付款保持锁定</td>
            </tr>
          `)
          .join("")
      : `<tr><td colspan="6">尚无已确认关系。设备级审计跟踪必须在前后向关系确认、设备明细拆分和 NodeID 绑定后生成。</td></tr>`;
  }

  if (paymentStageRows) {
    const stages = cr.paymentStageTemplates?.length
      ? cr.paymentStageTemplates
      : [
          { stageName: "预付款", triggerCondition: "合同签署、预付款条款满足", ownerAuditRequirement: "通常不要求设备审计", upstreamRequirement: "背靠背条款下需上游预付款到账", releaseRule: "按设备级资金池和后向条款释放" },
          { stageName: "到货款", triggerCondition: "到货、入库、发票", ownerAuditRequirement: "可按到货批次预审", upstreamRequirement: "对应设备到货款或可用资金池到账", releaseRule: "按到货数量/金额部分释放" },
          { stageName: "安装款", triggerCondition: "安装签证、安装数量确认", ownerAuditRequirement: "可按安装点位预审", upstreamRequirement: "对应设备安装阶段回款或资金池可用", releaseRule: "按实际安装数量折算释放" },
          { stageName: "验收款", triggerCondition: "验收单、验收数量确认", ownerAuditRequirement: "验收资料进入业主审计", upstreamRequirement: "对应设备验收阶段回款或资金池可用", releaseRule: "按验收通过数量/金额释放" },
          { stageName: "审计款", triggerCondition: "业主审计确认数量和金额", ownerAuditRequirement: "必须有正式审计结果", upstreamRequirement: "审计款到账或可归集到设备资金池", releaseRule: "以审计确认金额为基数释放" },
          { stageName: "质保金", triggerCondition: "质保期届满、无遗留问题", ownerAuditRequirement: "受最终审计确认金额约束", upstreamRequirement: "质保/尾款到账或可用", releaseRule: "扣除问题和审计扣减后释放" },
        ];
    paymentStageRows.innerHTML = stages
      .map(
        (stage) => `
          <tr>
            <td><strong>${escapeHtml(stage.stageName || "-")}</strong><br /><small>${escapeHtml(stage.stageCode || "")}</small></td>
            <td>${escapeHtml(stage.triggerCondition || "-")}</td>
            <td>${escapeHtml(stage.ownerAuditRequirement || "-")}</td>
            <td>${escapeHtml(stage.upstreamRequirement || "-")}</td>
            <td>${escapeHtml(stage.releaseRule || "-")}</td>
          </tr>
        `,
      )
      .join("");
  }
}

function renderWarehouse() {
  const devices = allDevices();
  const delivered = devices.reduce((sum, item) => sum + (item.deliveredQty || 0), 0);
  const inbound = devices.reduce((sum, item) => sum + (item.inboundQty || 0), 0);
  const installed = devices.reduce((sum, item) => sum + (item.installedQty || 0), 0);
  const warehouse = devices.reduce((sum, item) => sum + (item.warehouseQty || 0), 0);
  const cards = [
    ["已送货数量", delivered, "送货单数量汇总"],
    ["已入库数量", inbound, "入库单数量汇总"],
    ["已安装数量", installed, "点位安装记录汇总"],
    ["仓库数量", warehouse, "在仓库数量汇总"],
    ["送货未安装", Math.max(0, delivered - installed), "履约风险排查"],
    ["未匹配设备池", data.stats.unmatchedRows, "需要人工治理"],
  ];
  $("#warehouseGrid").innerHTML = cards
    .map(
      ([title, value, note]) => `
        <div class="warehouse-card">
          <span>${title}</span>
          <strong>${formatNumber(value)}</strong>
          <small>${note}</small>
        </div>
      `,
    )
    .join("");
}

function renderUnmatchedRows() {
  $("#unmatchedRows").innerHTML = unmatchedItems
    .map(
      (item) => `
        <tr>
          <td><span class="status-pill ${item.status === "待治理" ? "warn" : "done"}">${item.status}</span></td>
          <td>${item.assetType}</td>
          <td><strong>${item.material}</strong><br /><small>${item.id}</small></td>
          <td>${item.supplier}</td>
          <td>${item.contract}</td>
          <td>${item.target}</td>
          <td>
            <button class="link-btn" data-govern-unmatched="${item.id}">治理</button>
          </td>
        </tr>
      `,
    )
    .join("");
}

function renderImportCenter() {
  const history = state.importHistory.length
    ? state.importHistory
    : [
        {
          id: "BOOT-SITES",
          type: "sites",
          fileName: "点位管理表-无锡车路云.xlsx",
          validRows: data.stats.siteTotal,
          invalidRows: 1,
          status: "内置数据",
          createdAt: "当前样例",
        },
        {
          id: "BOOT-DEVICES",
          type: "devices",
          fileName: "设备安装位置表.xlsx",
          validRows: data.stats.deviceRows,
          invalidRows: 0,
          matchedRows: data.stats.deviceRows,
          unmatchedRows: data.stats.unmatchedRows,
          status: "内置数据",
          createdAt: "当前样例",
        },
      ];
  $("#importTaskList").innerHTML = history
    .map(
      (task) => `
        <div class="task-item">
          <div>
            <strong>${task.fileName}</strong>
            <span>${task.id} / ${task.createdAt}</span>
          </div>
          <dl>
            <dt>有效</dt><dd>${formatNumber(task.validRows)}</dd>
            <dt>失败</dt><dd>${formatNumber(task.invalidRows)}</dd>
            <dt>${task.type === "devices" ? "未匹配" : "状态"}</dt><dd>${task.type === "devices" ? formatNumber(task.unmatchedRows) : task.status}</dd>
          </dl>
        </div>
      `,
    )
    .join("");
  $("#importErrorList").innerHTML = importErrors
    .map(
      (error) => `
        <div class="error-item ${error.level}">
          <strong>${error.code}</strong>
          <span>${error.sheet} 第 ${error.row} 行</span>
          <p>${error.message}</p>
        </div>
      `,
    )
    .join("");
}

function coordinateIssueSites() {
  const issues = sites.filter((site) => site.issueCount > 0 || !site.lngGcj || !site.latGcj).slice(0, 30);
  return issues.length ? issues : sites.slice(0, 12).map((site, index) => ({ ...site, syntheticIssue: index % 3 }));
}

function renderCoordinateIssues() {
  const rows = coordinateIssueSites();
  const mapItems = rows.filter((site) => site.lngGcj && site.latGcj).slice(0, 20);
  if ($("#coordinateMap")) {
    renderMap($("#coordinateMap"), mapItems.length ? mapItems : sites.slice(0, 10), {
      label: "坐标异常点位复核，地图仅使用 GCJ-02 坐标",
    });
  }
  $("#coordinateRows").innerHTML = rows
    .map((site, index) => {
      const reason = site.issueCount > 0 ? "坐标转换失败或数据待复核" : index % 2 ? "原始坐标缺失审计说明" : "经开区归并后需复核区域";
      return `
        <tr data-open-site="${site.nodeId}" tabindex="0">
          <td><strong>${site.nodeId}</strong></td>
          <td>${site.name}</td>
          <td>${districtBadge(site)}</td>
          <td>${reason}</td>
          <td>保留 CGCS2000 原始值，重新生成 GCJ-02 展示坐标</td>
        </tr>
      `;
    })
    .join("");
  bindSiteOpenButtons();
}

function renderDocuments() {
  const docState = state.documentAssets;
  const summary = docState.summary || { categoryCounts: {}, fileCount: 0, totalSizeBytes: 0, duplicateFileCount: 0, unlinkedFileCount: 0 };
  const categoryCounts = summary.categoryCounts || {};
  const status = $("#documentAssetStatus");
  if (status) {
    status.className = `document-asset-status${docState.error ? " error" : ""}`;
    status.textContent = docState.loading
      ? "正在读取文档资产扫描报告……"
      : docState.error
        ? `读取失败：${docState.error}`
        : `批次 ${docState.batchNo || "-"} · ${docState.generatedAt || "-"} · ${docState.reportPath || "sample"}`;
  }

  const metrics = $("#documentMetrics");
  if (metrics) {
    metrics.innerHTML = [
      ["文件总数", formatNumber(summary.fileCount), "元数据记录，不加载原件"],
      ["容量规模", formatBytes(summary.totalSizeBytes), "用于评估本地库/NAS/对象存储"],
      ["重复文件", formatNumber(summary.duplicateFileCount), "sha256 去重治理"],
      ["未关联", formatNumber(summary.unlinkedFileCount), "需补充业务对象关系"],
    ]
      .map(([label, value, note]) => `
        <article class="metric document-metric">
          <span>${label}</span>
          <strong>${value}</strong>
          <small>${note}</small>
        </article>
      `)
      .join("");
  }

  const grid = $("#documentGrid");
  if (grid) {
    const categories = Object.entries(documentCategoryMeta);
    grid.innerHTML = categories
      .map(([key, [title, note]]) => `
        <article class="document-card ${key === "unclassified" && (categoryCounts[key] || 0) ? "warning" : ""}">
          <span>${title}</span>
          <strong>${formatNumber(categoryCounts[key] || 0)}</strong>
          <p>${note}</p>
          <button class="link-btn" type="button">${key === "unclassified" ? "治理分类" : "查看元数据"}</button>
        </article>
      `)
      .join("");
  }

  const categoryCount = $("#documentCategoryCount");
  if (categoryCount) categoryCount.textContent = `${Object.keys(categoryCounts).length} 类已有文件`;

  const rows = $("#documentRows");
  if (rows) {
    const records = (docState.records || []).slice(0, 20);
    rows.innerHTML = records.length
      ? records
          .map((record) => {
            const fields = record.extractedFields || {};
            const qualityClass = record.qualityStatus === "normal" && fields.documentStatusHint !== "voided" ? "done" : "warn";
            const qualityText = fields.documentStatusHint === "voided" ? "需复核" : statusLabel(record.qualityStatus);
            return `
              <tr>
                <td title="${record.storageKey || ""}">${record.fileName || "-"}</td>
                <td>${documentCategoryDisplay(record)}</td>
                <td>${contractInsightDisplay(record)}</td>
                <td>${fields.amountText || "-"}</td>
                <td>${fields.signDate || "-"}</td>
                <td>${fieldFlagsLabel(record)}</td>
                <td>${(record.fileExt || "-").toUpperCase()} · ${formatBytes(record.fileSizeBytes)}</td>
                <td><span class="status-pill ${qualityClass}">${qualityText}</span></td>
              </tr>
            `;
          })
          .join("")
      : `<tr><td colspan="8">暂无扫描记录。运行 tools/scan_document_assets.py 生成 document_assets/import_batches/latest-document-scan.json。</td></tr>`;
  }

  const recordCount = $("#documentRecordCount");
  if (recordCount) recordCount.textContent = `展示 ${Math.min((docState.records || []).length, 20)} / ${formatNumber(summary.fileCount)} 条`;
}

function isRoadsideEnabled(device) {
  return device.enabledStatus === "启用";
}

function isRoadsideAbnormal(device) {
  return device.status === "异常";
}

function roadsideRowsForSelectedDate() {
  if (state.opsDate === roadsideStatusState.currentDate) return roadsideStatusState.currentRows;
  const archive = roadsideStatusState.archives.find((item) => item.importDate === state.opsDate);
  return archive?.rows || roadsideStatusState.currentRows;
}

function roadsideSummary(rows) {
  const total = rows.length;
  const enabled = rows.filter(isRoadsideEnabled);
  const disabled = total - enabled.length;
  const online = enabled.filter((device) => device.status === "在线").length;
  const offline = enabled.filter((device) => device.status === "离线").length;
  const abnormal = enabled.filter(isRoadsideAbnormal).length;
  return { total, enabled: enabled.length, disabled, online, offline, abnormal };
}

function percent(part, total) {
  if (!total) return "0%";
  return `${((part / total) * 100).toFixed(1)}%`;
}

function roadsideTypeStats(rows) {
  const enabled = rows.filter(isRoadsideEnabled);
  const groups = new Map();
  enabled.forEach((device) => {
    const type = device.deviceTypeName || "未标注";
    if (!groups.has(type)) groups.set(type, { type, online: 0, offline: 0, abnormal: 0, total: 0 });
    const group = groups.get(type);
    group.total += 1;
    if (device.status === "在线") group.online += 1;
    if (device.status === "离线") group.offline += 1;
    if (isRoadsideAbnormal(device)) group.abnormal += 1;
  });
  return Array.from(groups.values()).sort((a, b) => b.total - a.total);
}

function roadsideStatusCounts(devices) {
  const enabled = devices.filter(isRoadsideEnabled);
  return {
    offline: enabled.filter((device) => device.status === "离线").length,
    abnormal: enabled.filter(isRoadsideAbnormal).length,
  };
}

function roadsideAbnormalDevices(rows) {
  return rows
    .filter((device) => isRoadsideEnabled(device) && ["离线", "异常"].includes(device.status))
    .sort((a, b) => {
      const priority = (device) => (device.status === "异常" ? 2 : device.status === "离线" ? 1 : 0);
      return priority(b) - priority(a) || String(a.area || "").localeCompare(String(b.area || ""), "zh-CN") || String(a.deviceTypeName || "").localeCompare(String(b.deviceTypeName || ""), "zh-CN");
    });
}

function importedRoadsideDays() {
  const days = [{ importDate: roadsideStatusState.currentDate, rows: roadsideStatusState.currentRows, isCurrent: true }, ...roadsideStatusState.archives];
  const byDate = new Map();
  days.forEach((day) => {
    if (!day.importDate || byDate.has(day.importDate)) return;
    const summary = roadsideSummary(day.rows || []);
    byDate.set(day.importDate, {
      date: day.importDate,
      isCurrent: Boolean(day.isCurrent),
      abnormal: summary.abnormal,
      offline: summary.offline,
      enabled: summary.enabled,
      ratio: percent(summary.abnormal, summary.enabled),
    });
  });
  return Array.from(byDate.values()).sort((a, b) => b.date.localeCompare(a.date));
}

function opsCalendarDays() {
  const imported = new Map(importedRoadsideDays().map((day) => [day.date, day]));
  const [year, month] = (state.opsCalendarCursor || state.opsDate.slice(0, 7)).split("-").map(Number);
  const daysInMonth = new Date(year, month, 0).getDate();
  const firstWeekday = new Date(year, month - 1, 1).getDay();
  const blanks = Array.from({ length: firstWeekday }, () => ({ empty: true }));
  const days = Array.from({ length: daysInMonth }, (_, index) => {
    const date = `${year}-${String(month).padStart(2, "0")}-${String(index + 1).padStart(2, "0")}`;
    return imported.get(date) || { date, hasImport: false };
  });
  return [...blanks, ...days];
}

function opsCalendarTitle() {
  const [year, month] = (state.opsCalendarCursor || state.opsDate.slice(0, 7)).split("-");
  return state.opsCalendarMode === "year" ? `${year} 年` : `${year} 年 ${Number(month)} 月`;
}

function opsYearMonths() {
  const imported = importedRoadsideDays();
  const year = Number((state.opsCalendarCursor || state.opsDate.slice(0, 7)).slice(0, 4));
  return Array.from({ length: 12 }, (_, index) => {
    const month = `${year}-${String(index + 1).padStart(2, "0")}`;
    const days = imported.filter((day) => day.date.startsWith(month));
    return {
      month,
      label: `${index + 1} 月`,
      importedDays: days.length,
      abnormal: days.reduce((sum, day) => sum + day.abnormal, 0),
      offline: days.reduce((sum, day) => sum + day.offline, 0),
    };
  });
}

function opsTrendDays() {
  const days = importedRoadsideDays()
    .filter((day) => day.date <= state.opsDate)
    .sort((a, b) => a.date.localeCompare(b.date));
  return days.length ? days : importedRoadsideDays().sort((a, b) => a.date.localeCompare(b.date));
}

function smoothSvgPath(points) {
  if (!points.length) return "";
  if (points.length === 1) return `M ${points[0].x} ${points[0].y}`;
  const commands = [`M ${points[0].x} ${points[0].y}`];
  for (let index = 0; index < points.length - 1; index += 1) {
    const current = points[index];
    const next = points[index + 1];
    const previous = points[index - 1] || current;
    const afterNext = points[index + 2] || next;
    const cp1 = {
      x: current.x + (next.x - previous.x) / 6,
      y: current.y + (next.y - previous.y) / 6,
    };
    const cp2 = {
      x: next.x - (afterNext.x - current.x) / 6,
      y: next.y - (afterNext.y - current.y) / 6,
    };
    commands.push(`C ${cp1.x.toFixed(2)} ${cp1.y.toFixed(2)}, ${cp2.x.toFixed(2)} ${cp2.y.toFixed(2)}, ${next.x.toFixed(2)} ${next.y.toFixed(2)}`);
  }
  return commands.join(" ");
}

function bindOpsTrendHover(target) {
  const groups = Array.from(target.querySelectorAll(".ops-trend-hit-group"));
  groups.forEach((group) => {
    const zone = group.querySelector(".ops-trend-hit-zone");
    if (!zone) return;
    zone.setAttribute("tabindex", "0");
    const show = () => group.setAttribute("data-active", "true");
    const hide = () => group.removeAttribute("data-active");
    zone.addEventListener("pointerenter", show);
    zone.addEventListener("pointerleave", hide);
    zone.addEventListener("mouseenter", show);
    zone.addEventListener("mouseleave", hide);
    zone.addEventListener("focus", () => group.setAttribute("data-active", "true"));
    zone.addEventListener("blur", () => group.removeAttribute("data-active"));
  });
}

function renderOpsTrendChart() {
  const target = $("#opsTrendCard");
  if (!target) return;
  const days = opsTrendDays();
  if (!days.length) {
    target.innerHTML = `
      <div class="ops-trend-empty">
        <strong>离线/异常趋势</strong>
        <span>暂无可绘制的历史导入数据</span>
      </div>
    `;
    return;
  }
  const width = 960;
  const height = 320;
  const pad = { top: 42, right: 34, bottom: 72, left: 58 };
  const innerWidth = width - pad.left - pad.right;
  const innerHeight = height - pad.top - pad.bottom;
  const maxValue = Math.max(1, ...days.flatMap((day) => [day.offline, day.abnormal]));
  const xFor = (index) => pad.left + (days.length === 1 ? innerWidth / 2 : (index / (days.length - 1)) * innerWidth);
  const yFor = (value) => pad.top + innerHeight - (value / maxValue) * innerHeight;
  const offlinePoints = days.map((day, index) => ({ x: xFor(index), y: yFor(day.offline), value: day.offline, date: day.date }));
  const abnormalPoints = days.map((day, index) => ({ x: xFor(index), y: yFor(day.abnormal), value: day.abnormal, date: day.date }));
  const offlinePath = smoothSvgPath(offlinePoints);
  const abnormalPath = smoothSvgPath(abnormalPoints);
  const areaPath = `${offlinePath} L ${offlinePoints.at(-1).x} ${pad.top + innerHeight} L ${offlinePoints[0].x} ${pad.top + innerHeight} Z`;
  const selectedIndex = days.findIndex((day) => day.date === state.opsDate);
  const activeIndex = selectedIndex >= 0 ? selectedIndex : days.length - 1;
  const selected = days[selectedIndex] || days.at(-1);
  const selectedPoint = offlinePoints[activeIndex] || offlinePoints.at(-1);
  const ticks = Array.from({ length: 4 }, (_, index) => Math.round((maxValue / 3) * index));
  const hitWidth = Math.max(18, innerWidth / Math.max(1, days.length - 1));
  const labelRotation = days.length > 10 ? -36 : 0;
  target.innerHTML = `
    <div class="ops-trend-head">
      <div>
        <strong>离线/异常趋势</strong>
        <span>${days[0].date} 至 ${selected.date}</span>
      </div>
      <div class="ops-trend-kpi">
        <b>${formatNumber(selected.offline)}</b><span>离线</span>
        <b>${formatNumber(selected.abnormal)}</b><span>异常</span>
      </div>
    </div>
    <svg class="ops-trend-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="离线和异常趋势折线图">
      <defs>
        <linearGradient id="opsOfflineStroke" x1="0" x2="1" y1="0" y2="0">
          <stop offset="0%" stop-color="#fff8ea" />
          <stop offset="55%" stop-color="#ffb454" />
          <stop offset="100%" stop-color="#ffd879" />
        </linearGradient>
        <linearGradient id="opsOfflineArea" x1="0" x2="0" y1="0" y2="1">
          <stop offset="0%" stop-color="#ffb454" stop-opacity="0.5" />
          <stop offset="64%" stop-color="#ffb454" stop-opacity="0.12" />
          <stop offset="100%" stop-color="#ffb454" stop-opacity="0" />
        </linearGradient>
        <filter id="opsGlassGlow" x="-20%" y="-30%" width="140%" height="180%">
          <feGaussianBlur stdDeviation="5" result="blur" />
          <feMerge>
            <feMergeNode in="blur" />
            <feMergeNode in="SourceGraphic" />
          </feMerge>
        </filter>
      </defs>
      ${ticks
        .map((tick) => {
          const y = yFor(tick);
          return `<g class="ops-trend-grid horizontal"><line x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}" /><text x="${pad.left - 12}" y="${y + 4}">${tick}</text></g>`;
        })
        .join("")}
      ${days
        .map((day, index) => {
          const x = xFor(index);
          return `<g class="ops-trend-grid vertical"><line x1="${x}" y1="${pad.top}" x2="${x}" y2="${pad.top + innerHeight}" /></g>`;
        })
        .join("")}
      <path class="ops-trend-area" d="${areaPath}" />
      <path class="ops-trend-line offline" d="${offlinePath}" />
      <path class="ops-trend-line abnormal" d="${abnormalPath}" />
      <line class="ops-trend-cursor active" x1="${selectedPoint.x}" y1="${pad.top}" x2="${selectedPoint.x}" y2="${pad.top + innerHeight}" />
      <circle class="ops-trend-dot offline" cx="${selectedPoint.x}" cy="${selectedPoint.y}" r="6" />
      <text class="ops-trend-tip" x="${Math.min(width - 126, selectedPoint.x + 10)}" y="${Math.max(28, selectedPoint.y - 12)}">${formatNumber(selected.offline)} 离线</text>
      ${days
        .map((day, index) => {
          const x = xFor(index);
          const transform = labelRotation ? ` transform="rotate(${labelRotation} ${x} ${height - 28})"` : "";
          return `<text class="ops-trend-date" x="${x}" y="${height - 28}"${transform}>${day.date.slice(5)}</text>`;
        })
        .join("")}
      ${days
        .map((day, index) => {
          const x = xFor(index);
          const offlinePoint = offlinePoints[index];
          const abnormalPoint = abnormalPoints[index];
          const tipX = Math.min(width - 164, Math.max(pad.left + 6, x + 12));
          const tipY = Math.max(pad.top + 14, Math.min(pad.top + innerHeight - 64, Math.min(offlinePoint.y, abnormalPoint.y) - 52));
          return `
            <g class="ops-trend-hit-group">
              <rect class="ops-trend-hit-zone" x="${x - hitWidth / 2}" y="${pad.top - 10}" width="${hitWidth}" height="${innerHeight + 28}" />
              <g class="ops-trend-hover">
                <line class="ops-trend-cursor" x1="${x}" y1="${pad.top}" x2="${x}" y2="${pad.top + innerHeight}" />
                <circle class="ops-trend-dot offline" cx="${offlinePoint.x}" cy="${offlinePoint.y}" r="6" />
                <circle class="ops-trend-dot abnormal" cx="${abnormalPoint.x}" cy="${abnormalPoint.y}" r="5" />
                <rect class="ops-trend-tooltip-bg" x="${tipX}" y="${tipY}" width="150" height="62" rx="10" />
                <text class="ops-trend-tooltip" x="${tipX + 12}" y="${tipY + 21}">
                  <tspan class="date">${day.date}</tspan>
                  <tspan x="${tipX + 12}" dy="19" class="offline">离线 ${formatNumber(day.offline)}</tspan>
                  <tspan x="${tipX + 82}" dy="0" class="abnormal">异常 ${formatNumber(day.abnormal)}</tspan>
                </text>
              </g>
            </g>
          `;
        })
        .join("")}
    </svg>
    <div class="ops-trend-legend">
      <span><i class="offline"></i>离线数量</span>
      <span><i class="abnormal"></i>异常数量</span>
    </div>
  `;
  bindOpsTrendHover(target);
}

function shiftOpsCalendar(direction) {
  const [year, month] = (state.opsCalendarCursor || state.opsDate.slice(0, 7)).split("-").map(Number);
  if (state.opsCalendarMode === "year") {
    state.opsCalendarCursor = `${year + direction}-${String(month).padStart(2, "0")}`;
  } else {
    const next = new Date(year, month - 1 + direction, 1);
    state.opsCalendarCursor = `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, "0")}`;
  }
}

async function moveRoadsideDay(sourceDate, targetDate) {
  if (!sourceDate || !targetDate || sourceDate === targetDate) return;
  const sourceRows = roadsideRowsForSelectedDateFor(sourceDate);
  if (!sourceRows.length) return;
  removeRoadsideDay(targetDate);
  if (sourceDate === roadsideStatusState.currentDate) {
    const previousCurrentDate = roadsideStatusState.currentDate;
    const previousCurrentRows = roadsideStatusState.currentRows;
    const replacement = importedRoadsideDays().find((day) => day.date !== sourceDate && day.date !== targetDate);
    roadsideStatusState.archives.unshift({
      importDate: targetDate,
      rows: previousCurrentRows,
      archivedAt: new Date().toLocaleString("zh-CN", { hour12: false }),
    });
    if (replacement) {
      const replacementRows = roadsideRowsForSelectedDateFor(replacement.date);
      removeRoadsideDay(replacement.date);
      roadsideStatusState.currentDate = replacement.date;
      roadsideStatusState.currentRows = replacementRows;
    } else {
      roadsideStatusState.currentDate = previousCurrentDate;
      roadsideStatusState.currentRows = [];
    }
  } else {
    const archive = roadsideStatusState.archives.find((item) => item.importDate === sourceDate);
    if (archive) archive.importDate = targetDate;
  }
  state.opsDate = targetDate;
  state.opsCalendarCursor = targetDate.slice(0, 7);
  const backendSaved = await saveRoadsideStatusState();
  if (!backendSaved) alert("本地文件持久化服务未连接。本次日期调整只保存到了浏览器缓存，请通过 http://127.0.0.1:4173 访问并确认本地服务已启动。");
  renderOps();
}

function roadsideRowsForSelectedDateFor(date) {
  if (date === roadsideStatusState.currentDate) return roadsideStatusState.currentRows;
  const archive = roadsideStatusState.archives.find((item) => item.importDate === date);
  return archive?.rows || [];
}

function removeRoadsideDay(date) {
  roadsideStatusState.archives = roadsideStatusState.archives.filter((item) => item.importDate !== date);
}

function siteByNodeId(nodeId) {
  const normalized = String(nodeId || "").trim();
  if (!normalized) return null;
  return sites.find((site) => site.nodeId === normalized) || null;
}

function roadsideSiteForDevice(device) {
  return siteByNodeId(device.intersectionId);
}

function roadsidePointTitle(point) {
  return point.site ? `${point.site.name} / NodeID ${point.site.nodeId}` : point.fallbackTitle;
}

function roadsidePointVendorLabel(point) {
  const vendors = Array.from(new Set(point.devices.map((device) => device.vendorName).filter(Boolean)));
  if (!vendors.length) return "厂商：-";
  if (vendors.length === 1) return `厂商：${vendors[0]}`;
  return `厂商：多厂商：${vendors.slice(0, 3).join("、")}${vendors.length > 3 ? `等 ${vendors.length} 家` : ""}`;
}

function buildRoadsidePointGroups(rows) {
  const groups = new Map();
  rows.filter(isRoadsideEnabled).forEach((device) => {
    const site = roadsideSiteForDevice(device);
    const hasSiteCoordinate = site && site.lngGcj && site.latGcj;
    const hasDeviceCoordinate = device.lng && device.lat;
    if (!hasSiteCoordinate && !hasDeviceCoordinate) return;
    const key = hasSiteCoordinate ? `site:${site.nodeId}` : `device:${device.deviceId}`;
    if (!groups.has(key)) {
      groups.set(key, {
        key,
        site: hasSiteCoordinate ? site : null,
        lng: hasSiteCoordinate ? site.lngGcj : device.lng,
        lat: hasSiteCoordinate ? site.latGcj : device.lat,
        fallbackTitle: device.devicePosition || device.relatedIntersection || device.deviceId,
        devices: [],
      });
    }
    groups.get(key).devices.push(device);
  });
  return Array.from(groups.values())
    .map((point) => ({
      ...point,
      statusCounts: roadsideStatusCounts(point.devices),
      displayStatus: point.devices.some(isRoadsideAbnormal) ? "异常" : "离线",
    }))
    .filter((point) => point.statusCounts.offline > 0 || point.statusCounts.abnormal > 0);
}

function roadsidePointDeviceSummary(point) {
  const rowsByTypeVendor = new Map();
  point.devices.forEach((device) => {
    const type = device.deviceTypeName || "未标注";
    const vendor = device.vendorName || "-";
    const key = `${type}__${vendor}`;
    if (!rowsByTypeVendor.has(key)) rowsByTypeVendor.set(key, { type, vendor, total: 0, offline: 0, abnormal: 0 });
    const row = rowsByTypeVendor.get(key);
    row.total += 1;
    if (device.status === "离线") row.offline += 1;
    if (isRoadsideAbnormal(device)) row.abnormal += 1;
  });
  const rows = Array.from(rowsByTypeVendor.values()).sort((a, b) => (b.offline + b.abnormal) - (a.offline + a.abnormal) || b.total - a.total || a.vendor.localeCompare(b.vendor, "zh-CN"));
  return `
    <div class="ops-point-summary">
      <div class="ops-point-summary-row ops-point-summary-head">
        <span>设备类型</span>
        <span>厂商</span>
        <span>设备数量</span>
        <span class="ops-alert-count">离线数量</span>
        <span class="ops-alert-count">异常数量</span>
      </div>
      ${rows
        .map(
          (row) => `
            <div class="ops-point-summary-row">
              <span class="ops-point-type">${row.type}</span>
              <span class="ops-point-vendor-name">${row.vendor}</span>
              <span><b>${row.total}</b></span>
              <span class="ops-alert-count"><b>${row.offline}</b></span>
              <span class="ops-alert-count"><b>${row.abnormal}</b></span>
            </div>
          `,
        )
        .join("")}
    </div>
  `;
}

function renderAmapOpsMap(container, pointGroups, unmatchedCount) {
  if (!window.AMap || !pointGroups.length) return false;
  const previousMap = amapInstances.get(container);
  if (previousMap) {
    previousMap.destroy();
    amapInstances.delete(container);
  }
  container.innerHTML = "";
  container.classList.add("amap-live");
  const first = pointGroups[0];
  const map = new window.AMap.Map(container, {
    zoom: 12,
    center: [first.lng, first.lat],
    viewMode: "2D",
    mapStyle: "amap://styles/normal",
  });
  amapInstances.set(container, map);

  const info = new window.AMap.InfoWindow({
    offset: new window.AMap.Pixel(0, -8),
    autoMove: true,
  });
  let closeInfoTimer = null;

  const markers = pointGroups.map((point) => {
    const title = roadsidePointTitle(point);
    const marker = new window.AMap.Marker({
      position: [point.lng, point.lat],
      title,
      offset: new window.AMap.Pixel(-8, -8),
      zIndex: point.displayStatus === "异常" ? 80 : point.site ? 70 : 65,
      content: `<button class="amap-ops-marker ${point.displayStatus === "异常" ? "abnormal" : "offline"} ${point.site ? "" : "unmatched"}" title="${title}"></button>`,
    });
    marker.on("mouseover", () => {
      if (closeInfoTimer) window.clearTimeout(closeInfoTimer);
      info.setContent(`<div class="amap-ops-info"><strong>${title}</strong>${roadsidePointDeviceSummary(point)}</div>`);
      info.open(map, marker.getPosition());
    });
    marker.on("click", () => {
      if (closeInfoTimer) window.clearTimeout(closeInfoTimer);
      info.setContent(`<div class="amap-ops-info"><strong>${title}</strong>${roadsidePointDeviceSummary(point)}</div>`);
      info.open(map, marker.getPosition());
    });
    marker.on("mouseout", () => {
      closeInfoTimer = window.setTimeout(() => info.close(), 600);
    });
    return marker;
  });
  map.add(markers);
  if (markers.length > 1) map.setFitView(markers, false, [42, 42, 42, 42], 15);

  const label = document.createElement("div");
  label.className = "map-label";
  const deviceCount = pointGroups.reduce((sum, point) => sum + point.statusCounts.offline + point.statusCounts.abnormal, 0);
  label.textContent = `高德地图 | 离线/异常点位 ${pointGroups.length} 个，设备 ${deviceCount} 台${unmatchedCount ? `，${unmatchedCount} 台未匹配点位` : ""}`;
  container.appendChild(label);
  return true;
}

function renderOpsMap(rows) {
  const target = $("#opsDeviceMap");
  if (!target) return;
  const devices = rows
    .filter((device) => isRoadsideEnabled(device) && ["离线", "异常"].includes(device.status));
  const pointGroups = buildRoadsidePointGroups(rows).slice(0, 300);
  const unmatchedCount = pointGroups.filter((point) => !point.site).reduce((sum, point) => sum + point.statusCounts.offline + point.statusCounts.abnormal, 0);
  try {
    if (renderAmapOpsMap(target, pointGroups, unmatchedCount)) return;
  } catch (error) {
    console.warn("AMap ops render failed, fallback map enabled.", error);
  }
  target.innerHTML = "";
  target.classList.remove("amap-live");
  const pseudoSites = pointGroups.map((point) => ({
    nodeId: point.site?.nodeId || point.key,
    name: point.site?.name || point.fallbackTitle,
    district: point.site?.district || point.devices[0]?.area || "未匹配",
    lngGcj: point.lng,
    latGcj: point.lat,
    status: point.displayStatus,
    issueCount: point.statusCounts.abnormal,
  }));
  const bounds = mapBounds(pseudoSites.length ? pseudoSites : sites.slice(0, 10));
  pointGroups.forEach((point) => {
    const pos = markerPosition({ lngGcj: point.lng, latGcj: point.lat }, bounds);
    const marker = document.createElement("div");
    marker.className = `marker ops-device-marker ${point.displayStatus === "异常" ? "abnormal" : "offline"}`;
    marker.style.left = `${pos.left}%`;
    marker.style.top = `${pos.top}%`;
    marker.innerHTML = `
      <button aria-label="${roadsidePointTitle(point)}"></button>
      <span>${point.displayStatus}</span>
      <div class="ops-map-tooltip">
        <strong>${roadsidePointTitle(point)}</strong>
        ${roadsidePointDeviceSummary(point)}
      </div>
    `;
    target.appendChild(marker);
  });
  const label = document.createElement("div");
  label.className = "map-label";
  const deviceCount = pointGroups.reduce((sum, point) => sum + point.statusCounts.offline + point.statusCounts.abnormal, 0);
  label.textContent = `静态地图回退 | 离线/异常点位 ${pointGroups.length} 个，设备 ${deviceCount} 台${unmatchedCount ? `，${unmatchedCount} 台未匹配点位` : ""}`;
  target.appendChild(label);
}

function renderOps() {
  renderPersistenceStatus();
  const rows = roadsideRowsForSelectedDate();
  const summary = roadsideSummary(rows);
  $("#opsCurrentDate").textContent = `当前日期：${state.opsDate}`;
  $("#opsCalendarTitle").textContent = opsCalendarTitle();
  $("#opsSummaryGrid").innerHTML = [
    ["全部路侧设备总数", summary.total, "导入表全部设备"],
    ["已启用路侧设备总数", summary.enabled, "统计口径基数"],
    ["未启用路侧设备总数", summary.disabled, "不纳入在线率统计"],
    ["在线设备", summary.online, percent(summary.online, summary.enabled)],
    ["离线设备", summary.offline, percent(summary.offline, summary.enabled)],
    ["异常设备", summary.abnormal, percent(summary.abnormal, summary.enabled)],
  ]
    .map(
      ([title, value, note]) => `
        <div class="metric ops-metric">
          <span>${title}</span>
          <div class="ops-metric-value-row">
            <strong>${formatNumber(value)}</strong>
            <small>${note}</small>
          </div>
        </div>
      `,
    )
    .join("");
  $("#opsStatsRows").innerHTML = roadsideTypeStats(rows)
    .map(
      (item) => `
        <tr>
          <td><strong>${item.type}</strong></td>
          <td>${formatNumber(item.online)}</td>
          <td>${percent(item.online, item.total)}</td>
          <td>${formatNumber(item.offline)}</td>
          <td>${percent(item.offline, item.total)}</td>
          <td>${formatNumber(item.abnormal)}</td>
          <td>${percent(item.abnormal, item.total)}</td>
          <td>${formatNumber(item.total)}</td>
        </tr>
      `,
    )
    .join("");
  const abnormalDevices = roadsideAbnormalDevices(rows);
  const detailRows = abnormalDevices.slice(0, 500);
  $("#opsDetailCount").textContent = `显示 ${formatNumber(detailRows.length)} / ${formatNumber(abnormalDevices.length)} 条异常/离线设备`;
  $("#opsDeviceRows").innerHTML = detailRows
    .map(
      (device, index) => `
        <tr>
          <td>${index + 1}</td>
          <td>${device.devicePosition || device.relatedIntersection || "-"}<br /><small>${device.intersectionId || ""}</small></td>
          <td>${device.area || "-"}</td>
          <td><span class="device-type">${device.deviceTypeName || "-"}</span></td>
          <td>${device.vendorName || "-"}</td>
          <td><strong>${device.deviceId}</strong></td>
          <td><span class="status-pill ${device.status === "在线" ? "done" : "warn"}">${device.status || "-"}</span></td>
        </tr>
      `,
    )
    .join("");
  $$("[data-ops-calendar-mode]").forEach((button) => button.classList.toggle("active", button.dataset.opsCalendarMode === state.opsCalendarMode));
  $("#opsWeekdays").innerHTML = state.opsCalendarMode === "month"
    ? ["日", "一", "二", "三", "四", "五", "六"].map((day) => `<span>${day}</span>`).join("")
    : "";
  if (state.opsCalendarMode === "year") {
    $("#opsCalendar").classList.add("year-mode");
    $("#opsCalendar").innerHTML = opsYearMonths()
      .map(
        (month) => `
          <button class="year-month ${month.importedDays ? "" : "no-data"}" data-ops-month="${month.month}" title="${month.month} 导入 ${month.importedDays} 天">
            <strong class="calendar-day">${month.label}</strong>
            <span class="calendar-counts"><b>导入 ${month.importedDays} 天</b><b>异常 ${formatNumber(month.abnormal)}</b></span>
            <small>离线 ${formatNumber(month.offline)}</small>
          </button>
        `,
      )
      .join("");
  } else {
    $("#opsCalendar").classList.remove("year-mode");
    const calendarDays = opsCalendarDays();
    $("#opsCalendar").innerHTML = calendarDays
      .map((day) => {
        if (day.empty) return `<div class="calendar-blank" aria-hidden="true"></div>`;
        return `
          <button class="${day.date === state.opsDate ? "active" : ""} ${day.hasImport === false ? "no-data" : ""}" ${day.hasImport === false ? "" : `data-ops-date="${day.date}" draggable="true" data-ops-drag-date="${day.date}"`} data-ops-drop-date="${day.date}" title="${day.hasImport === false ? `${day.date} 无导入，可拖放导入日数据到此日期` : `${day.date} 异常 ${day.abnormal}，离线 ${day.offline}，异常占比 ${day.ratio}`}" >
            <strong class="calendar-day">${Number(day.date.slice(-2))}</strong>
            ${day.hasImport === false
              ? `<span class="calendar-muted">无导入</span><small>${day.date}</small>`
              : `<span class="calendar-counts"><b>异常 ${formatNumber(day.abnormal)}</b><b>离线 ${formatNumber(day.offline)}</b></span><small>${day.isCurrent ? "当前" : "历史"} · ${day.ratio}</small>`}
          </button>
        `;
      })
      .join("");
  }
  renderOpsTrendChart();
  renderOpsMap(rows);
}

function exportRoadsideExcel(kind) {
  const rows = roadsideRowsForSelectedDate();
  if (kind === "stats") {
    exportRoadsideStatsWorkbook(rows);
  } else {
    exportRoadsideAbnormalWorkbook(rows);
  }
}

function writeWorkbookFile(fileName, sheetName, rows, columns) {
  if (!window.XLSX) {
    alert("缺少 XLSX 导出库，请确认 prototype/vendor/xlsx.full.min.js 已加载。");
    return;
  }
  const worksheet = window.XLSX.utils.aoa_to_sheet(rows);
  if (columns) worksheet["!cols"] = columns;
  const workbook = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  window.XLSX.writeFile(workbook, fileName);
}

function exportRoadsideStatsWorkbook(rows) {
  const workbookRows = [
      ["设备类型", "在线数量", "在线占比", "离线数量", "离线占比", "异常数量", "异常占比", "合计"],
      ...roadsideTypeStats(rows).map((item) => [
        item.type,
        item.online,
        percent(item.online, item.total),
        item.offline,
        percent(item.offline, item.total),
        item.abnormal,
        percent(item.abnormal, item.total),
        item.total,
      ]),
  ];
  writeWorkbookFile(`当日异常设备统计表-${state.opsDate}.xlsx`, "异常统计", workbookRows, [
    { wch: 18 },
    { wch: 12 },
    { wch: 12 },
    { wch: 12 },
    { wch: 12 },
    { wch: 12 },
    { wch: 12 },
    { wch: 10 },
  ]);
}

function exportRoadsideAbnormalWorkbook(rows) {
  const workbookRows = [
    ["序号", "点位/位置", "区域", "设备类型", "厂商", "设备编号", "状态"],
    ...roadsideAbnormalDevices(rows).map((device, index) => [
      index + 1,
      device.devicePosition || device.relatedIntersection || "-",
      device.area || "-",
      device.deviceTypeName || "-",
      device.vendorName || "-",
      device.deviceId || "-",
      device.status || "-",
    ]),
  ];
  writeWorkbookFile(`路侧设备运行状态异常详表-${state.opsDate}.xlsx`, "异常详表", workbookRows, [
    { wch: 8 },
    { wch: 30 },
    { wch: 12 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 10 },
  ]);
}

function nearestSites(site) {
  return sites
    .filter((item) => item.nodeId !== site.nodeId)
    .map((item) => ({
      ...item,
      distance: Math.hypot(item.lngGcj - site.lngGcj, item.latGcj - site.latGcj),
    }))
    .sort((a, b) => a.distance - b.distance)
    .slice(0, 5);
}

function renderDrawer(site) {
  $("#drawerNodeId").textContent = `NodeID ${site.nodeId}`;
  $("#drawerTitle").textContent = site.name;
  const nearby = [site, ...nearestSites(site)];
  renderMap($("#drawerMap"), nearby, {
    label: `当前点位：${site.name}；周边点位：${nearby.length - 1} 个`,
    centerSite: site,
    currentNodeId: site.nodeId,
    zoom: 15,
    maxZoom: 16,
  });

  $("#detailBase").innerHTML = `
    <div class="kv-grid">
      <div class="kv"><span>行政区域</span><strong>${districtBadge(site)}</strong></div>
      <div class="kv"><span>点位类型</span><strong>${site.type}</strong></div>
      <div class="kv"><span>信号机厂商</span><strong>${site.vendor}</strong></div>
      <div class="kv"><span>感知点位类型</span><strong>${perceptionType(site)}</strong></div>
      <div class="kv"><span>GCJ-02 经度</span><strong>${site.lngGcj.toFixed(8)}</strong></div>
      <div class="kv"><span>GCJ-02 纬度</span><strong>${site.latGcj.toFixed(8)}</strong></div>
      <div class="kv"><span>是否自适应点位</span><strong>${site.adaptive ? "是" : "否"}</strong></div>
      <div class="kv"><span>是否可变车道路口</span><strong>${isVariableLane(site) ? "是" : "否"}</strong></div>
      <div class="kv"><span>CGCS2000</span><strong>仅原始数据审计</strong></div>
    </div>
  `;

  $("#detailDevice").innerHTML = `
    <div class="table-wrap">
      <table>
        <thead><tr><th>设备类型</th><th>安装位置</th><th>设备</th><th>供应商</th><th>已安装</th></tr></thead>
        <tbody>
          ${site.devices
            .map(
              (device) => `
                <tr>
                  <td><span class="device-type">${deviceType(device)}</span></td>
                  <td>${device.position || "-"}</td>
                  <td>${device.name || "-"}<br /><small>${device.model || ""}</small></td>
                  <td>${device.supplier || "-"}</td>
                  <td>${device.installedQty || 0}</td>
                </tr>
              `,
            )
            .join("")}
        </tbody>
      </table>
    </div>
  `;

  $("#detailArchive").innerHTML = `
    <div class="kv-grid">
      <div class="kv"><span>1. 基础档案</span><strong>NodeID、CrossID、点位名称、区域、类型、信号机厂商</strong></div>
      <div class="kv"><span>2. 坐标与地图</span><strong>GCJ-02 坐标、当前点位、周边点位距离</strong></div>
      <div class="kv"><span>3. 信号系统信息</span><strong>${site.vendor || "-"} / 附件待归集</strong></div>
      <div class="kv"><span>4. 感知与自适应</span><strong>${perceptionType(site)} / ${site.adaptive ? "自适应" : "非自适应"}</strong></div>
      <div class="kv"><span>5. 可变车道</span><strong>${isVariableLane(site) ? `${site.variableLaneCount || "-"} / ${site.variableLaneDirection || "-"}` : "无"}</strong></div>
      <div class="kv"><span>6. 设备安装清单</span><strong>${site.devices.length} 条，按设备类型逐个展示</strong></div>
      <div class="kv"><span>7. 合同与供应商</span><strong>${new Set(site.devices.map((d) => d.supplier).filter(Boolean)).size} 家供应商</strong></div>
      <div class="kv"><span>8. 单据资料</span><strong>送货、入库、领料、开箱验收索引</strong></div>
      <div class="kv"><span>9. 施工与验收</span><strong>${site.archiveCompleteness}% 完整度</strong></div>
      <div class="kv"><span>10. 运维工单</span><strong>${site.issueCount ? "有异常待闭环" : "暂无未闭环异常"}</strong></div>
      <div class="kv"><span>11. 操作日志</span><strong>导入、编辑、匹配、删除审计</strong></div>
    </div>
  `;

  const deviceAmount = site.devices.reduce((sum, device) => sum + (device.backAmount || 0), 0);
  $("#detailContract").innerHTML = `
    <div class="kv-grid">
      <div class="kv"><span>合同流路径</span><strong>天安 -> 万集</strong></div>
      <div class="kv"><span>后向合同金额</span><strong>${formatMoney(deviceAmount)}</strong></div>
      <div class="kv"><span>付款依赖</span><strong>上游回款后触发</strong></div>
      <div class="kv"><span>设备级匹配</span><strong>${site.devices.length ? "自动匹配待确认" : "无设备"}</strong></div>
    </div>
  `;
}

function openSite(nodeId) {
  const site = sites.find((item) => item.nodeId === String(nodeId));
  if (!site) return;
  state.selectedSite = site;
  try {
    renderDrawer(site);
  } catch (error) {
    console.error("Render site detail failed.", error);
    renderFallbackDrawer(site);
  }
  $("#detailDrawer").classList.add("open");
  $("#detailDrawer").setAttribute("aria-hidden", "false");
  $("#scrim").classList.add("open");
}

function renderFallbackDrawer(site) {
  $("#drawerNodeId").textContent = `NodeID ${site.nodeId}`;
  $("#drawerTitle").textContent = site.name;
  renderFallbackMap($("#drawerMap"), [site, ...nearestSites(site)], {
    label: `当前点位：${site.name}`,
    currentNodeId: site.nodeId,
  });
  $("#detailBase").innerHTML = `<div class="kv"><span>点位</span><strong>${site.name}</strong></div>`;
  $("#detailDevice").innerHTML = "";
  $("#detailArchive").innerHTML = "";
  $("#detailContract").innerHTML = "";
}

function closeDrawer() {
  $("#detailDrawer").classList.remove("open");
  $("#detailDrawer").setAttribute("aria-hidden", "true");
  $("#scrim").classList.remove("open");
}

function bindEvents() {
  $$(".nav-item").forEach((item) => item.addEventListener("click", () => setPanel(item.dataset.panel)));
  $$("[data-panel-link]").forEach((item) => item.addEventListener("click", () => setPanel(item.dataset.panelLink)));
  $("#globalSearch").addEventListener("input", (event) => {
    state.query = event.target.value;
    renderSites();
    renderDevices();
    renderMapAssets();
  });
  $$(".segmented button").forEach((button) =>
    button.addEventListener("click", () => {
      state.view = button.dataset.view;
      $$(".segmented button").forEach((item) => item.classList.toggle("active", item === button));
      $$(".site-view").forEach((view) => view.classList.toggle("active", view.id === `${state.view}View`));
      requestAnimationFrame(applySiteViewLayoutSettled);
      if (state.view === "map") {
        requestAnimationFrame(() => {
          applySiteViewLayoutSettled();
          renderMap($("#siteMap"), filteredSites());
        });
      }
    }),
  );
  document.body.addEventListener("click", (event) => {
    const chooseImport = event.target.closest("[data-choose-import]");
    if (chooseImport) {
      const type = chooseImport.dataset.chooseImport;
      const input = importInputForType(type);
      input.value = "";
      input.click();
      return;
    }
    const simulateImport = event.target.closest("[data-simulate-import]");
    if (simulateImport) {
      startImport(simulateImport.dataset.simulateImport);
      return;
    }
    const confirmButton = event.target.closest("[data-confirm-import]");
    if (confirmButton) {
      confirmImport(confirmButton.dataset.confirmImport);
      return;
    }
    const contractEvidence = event.target.closest("[data-contract-candidate-evidence]");
    if (contractEvidence) {
      state.contractReview.selectedCandidateId = contractEvidence.dataset.contractCandidateEvidence || "";
      renderContractSystemViews();
      return;
    }
    const contractDecision = event.target.closest("[data-contract-candidate-decision]");
    if (contractDecision) {
      setContractCandidateDecision(contractDecision.dataset.candidateId, contractDecision.dataset.contractCandidateDecision);
      return;
    }
    const macroFlowSelect = event.target.closest("[data-select-macro-flow]");
    if (macroFlowSelect) {
      state.contractRelationships.selectedMacroFlowId = macroFlowSelect.dataset.selectMacroFlow || "";
      renderContractSystemViews();
      return;
    }
    const finishButton = event.target.closest("[data-import-finish]");
    if (finishButton) {
      const finishType = finishButton.dataset.importFinish;
      setPanel(finishType === "sites" ? "sites" : finishType === "roadsideStatus" ? "ops" : "devices");
      return;
    }
    const deleter = event.target.closest("[data-delete-site]");
    if (deleter) {
      event.preventDefault();
      event.stopPropagation();
      const site = sites.find((item) => item.nodeId === deleter.dataset.deleteSite);
      if (!site) return;
      if (site.devices.length) {
        alert(`点位 ${site.name} 下挂有 ${site.devices.length} 条设备记录。请先到设备管理删除该点位下设备后，再删除点位。`);
      } else {
        const confirmed = confirm(`确认删除点位 ${site.name}？该操作会写入删除审计。`);
        if (confirmed) {
          const index = sites.findIndex((item) => item.nodeId === site.nodeId);
          if (index >= 0) sites.splice(index, 1);
          data.stats.siteTotal = Math.max(0, data.stats.siteTotal - 1);
          renderMetrics();
          renderFilters();
          renderSites();
          renderCoordinateIssues();
        }
      }
      return;
    }
    const openImport = event.target.closest("[data-open-import]");
    if (openImport) {
      showImportPanel(openImport.dataset.openImport);
      return;
    }
    const governButton = event.target.closest("[data-govern-unmatched]");
    if (governButton) {
      const item = unmatchedItems.find((entry) => entry.id === governButton.dataset.governUnmatched);
      if (!item) return;
      item.status = "已建议";
      alert(`${item.material} 已按“${item.target}”生成治理建议，等待业务确认。`);
      renderUnmatchedRows();
      return;
    }
    const siteFilterTrigger = event.target.closest("[data-site-filter-open]");
    if (siteFilterTrigger) {
      const key = siteFilterTrigger.dataset.siteFilterOpen;
      state.activeSiteFilter = state.activeSiteFilter === key ? null : key;
      renderSiteColumnFilters();
      return;
    }
    const siteFilterClear = event.target.closest("[data-site-filter-clear]");
    if (siteFilterClear) {
      state.siteColumnFilters[siteFilterClear.dataset.siteFilterClear] = [];
      renderSites();
      return;
    }
    const siteFilterOption = event.target.closest("[data-site-filter-option]");
    if (siteFilterOption) {
      const key = siteFilterOption.dataset.siteFilterOption;
      const selected = new Set(state.siteColumnFilters[key] || []);
      if (siteFilterOption.checked) selected.add(siteFilterOption.value);
      else selected.delete(siteFilterOption.value);
      state.siteColumnFilters[key] = Array.from(selected);
      renderSites();
      return;
    }
    if (state.activeSiteFilter && !event.target.closest("#siteFilterPopover")) {
      state.activeSiteFilter = null;
      renderSiteColumnFilters();
    }
    const filterTrigger = event.target.closest("[data-map-asset-filter-open]");
    if (filterTrigger) {
      const key = filterTrigger.dataset.mapAssetFilterOpen;
      state.activeMapAssetFilter = state.activeMapAssetFilter === key ? null : key;
      renderMapAssetColumnFilters();
      return;
    }
    const filterClear = event.target.closest("[data-map-asset-filter-clear]");
    if (filterClear) {
      state.mapAssetColumnFilters[filterClear.dataset.mapAssetFilterClear] = [];
      state.selectedMapAssetId = null;
      renderMapAssets();
      return;
    }
    const filterOption = event.target.closest("[data-map-asset-filter-option]");
    if (filterOption) {
      const key = filterOption.dataset.mapAssetFilterOption;
      const selected = new Set(state.mapAssetColumnFilters[key] || []);
      if (filterOption.checked) selected.add(filterOption.value);
      else selected.delete(filterOption.value);
      state.mapAssetColumnFilters[key] = Array.from(selected);
      state.selectedMapAssetId = null;
      renderMapAssets();
      return;
    }
    if (state.activeMapAssetFilter && !event.target.closest("#mapAssetFilterPopover")) {
      state.activeMapAssetFilter = null;
      renderMapAssetColumnFilters();
    }
    const opsDateButton = event.target.closest("[data-ops-date]");
    if (opsDateButton) {
      state.opsDate = opsDateButton.dataset.opsDate;
      state.opsCalendarCursor = state.opsDate.slice(0, 7);
      renderOps();
      return;
    }
    const opsMonthButton = event.target.closest("[data-ops-month]");
    if (opsMonthButton) {
      state.opsCalendarCursor = opsMonthButton.dataset.opsMonth;
      state.opsCalendarMode = "month";
      renderOps();
      return;
    }
    const calendarShift = event.target.closest("[data-ops-calendar-shift]");
    if (calendarShift) {
      shiftOpsCalendar(Number(calendarShift.dataset.opsCalendarShift));
      renderOps();
      return;
    }
    const calendarMode = event.target.closest("[data-ops-calendar-mode]");
    if (calendarMode) {
      state.opsCalendarMode = calendarMode.dataset.opsCalendarMode;
      renderOps();
      return;
    }
    const opener = event.target.closest("[data-open-site]");
    if (opener) openSite(opener.dataset.openSite);
    const mapAssetOpener = event.target.closest("[data-open-map-asset]");
    if (mapAssetOpener) {
      state.selectedMapAssetId = mapAssetOpener.dataset.openMapAsset;
      renderMapAssets();
    }
  });
  document.body.addEventListener("mousedown", (event) => {
    const handle = event.target.closest(".column-resizer");
    if (handle?.dataset.resizeTable === "site") startSiteColumnResize(event, handle);
    else if (handle) startMapAssetColumnResize(event, handle);
    const preview = event.target.closest("[data-map-asset-preview]");
    if (preview) startMapAssetPreviewDrag(event, preview);
  });
  document.body.addEventListener(
    "wheel",
    (event) => {
      const preview = event.target.closest("[data-map-asset-preview]");
      if (preview) handleMapAssetPreviewWheel(event, preview);
    },
    { passive: false },
  );
  document.body.addEventListener("dblclick", (event) => {
    const preview = event.target.closest("[data-map-asset-preview]");
    if (preview) setMapAssetPreviewTransform(preview, { scale: 1, x: 0, y: 0 });
  });
  document.addEventListener("mousemove", moveMapAssetPreviewDrag);
  document.addEventListener("mouseup", stopMapAssetPreviewDrag);
  document.addEventListener("mouseleave", stopMapAssetPreviewDrag);
  document.body.addEventListener("dragstart", (event) => {
    const dragDate = event.target.closest("[data-ops-drag-date]");
    if (!dragDate) return;
    event.dataTransfer.setData("text/plain", dragDate.dataset.opsDragDate);
    event.dataTransfer.effectAllowed = "move";
  });
  document.body.addEventListener("dragover", (event) => {
    const dropDate = event.target.closest("[data-ops-drop-date]");
    if (!dropDate) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = "move";
  });
  document.body.addEventListener("drop", (event) => {
    const dropDate = event.target.closest("[data-ops-drop-date]");
    if (!dropDate) return;
    event.preventDefault();
    moveRoadsideDay(event.dataTransfer.getData("text/plain"), dropDate.dataset.opsDropDate);
  });
  $$(".detail-tabs button").forEach((button) =>
    button.addEventListener("click", () => {
      $$(".detail-tabs button").forEach((item) => item.classList.toggle("active", item === button));
      $$(".detail-pane").forEach((pane) => pane.classList.toggle("active", pane.id === `detail${button.dataset.detailTab[0].toUpperCase()}${button.dataset.detailTab.slice(1)}`));
    }),
  );
  $("#closeDrawer").addEventListener("click", closeDrawer);
  $("#scrim").addEventListener("click", closeDrawer);
  $("#visualThemeToggle").addEventListener("click", toggleVisualTheme);
  $("#openFirstSite").addEventListener("click", () => openSite(sites[0].nodeId));
  $("#importSiteTable").addEventListener("click", () => showImportPanel("sites"));
  $("#importDeviceTable").addEventListener("click", () => showImportPanel("devices"));
  $("#siteImportInput").addEventListener("change", (event) => {
    const [file] = event.target.files;
    if (file) startImport("sites", file);
  });
  $("#deviceImportInput").addEventListener("change", (event) => {
    const [file] = event.target.files;
    if (file) startImport("devices", file);
  });
  $("#roadsideStatusImportInput").addEventListener("change", (event) => {
    const [file] = event.target.files;
    if (file) startImport("roadsideStatus", file);
  });
  $("#showUnmatched").addEventListener("click", () => {
    $("#unmatchedPanel").classList.toggle("open");
    $("#unmatchedTableWrap").classList.toggle("open");
  });
  $("#refreshMapAssets").addEventListener("click", () => loadMapAssets(true));
  $("#mapAssetRoot").addEventListener("click", () => {
    alert(mapAssetState.root || "地图资产目录尚未读取。");
  });
  $("#exportAllMapAssets").addEventListener("click", exportAllMapAssets);
  $("#exportIncompleteMapAssets").addEventListener("click", exportIncompleteMapAssets);
  document.body.addEventListener("input", (event) => {
    const filter = event.target.closest("[data-contract-filter]");
    if (!filter) return;
    state.contractReview.filters[filter.dataset.contractFilter] = filter.value;
    state.contractReview.selectedCandidateId = "";
    renderContractSystemViews();
  });
  document.body.addEventListener("change", (event) => {
    const filter = event.target.closest("[data-contract-filter]");
    if (!filter) return;
    state.contractReview.filters[filter.dataset.contractFilter] = filter.value;
    state.contractReview.selectedCandidateId = "";
    renderContractSystemViews();
  });
  $("#refreshDocumentAssets")?.addEventListener("click", loadDocumentAssets);
  $("#refreshContractRelationships")?.addEventListener("click", loadContractRelationships);
  $("#showDocumentStorageRule")?.addEventListener("click", () => {
    alert("文档原件进入 document_assets/raw 或对象存储；数据库保存 storage_key、sha256、版本、业务关系、解析结果和审计，不把大文件写入前端 data.js。");
  });
  $("#overviewUnmatchedCard").addEventListener("click", () => {
    setPanel("devices");
    $("#unmatchedPanel").classList.add("open");
    $("#unmatchedTableWrap").classList.add("open");
  });
  $("#reconvertCoordinates").addEventListener("click", () => {
    importErrors.unshift({
      sheet: "坐标异常",
      row: "-",
      level: "warning",
      code: "COORDINATE_RECONVERTED",
      message: "已按 CGCS2000≈WGS84 -> GCJ-02 重新转换当前异常清单。",
    });
    renderImportCenter();
    renderCoordinateIssues();
  });
  $("#drawerSave").addEventListener("click", () => alert("点位基础信息已保存，操作日志已记录。"));
  $("#drawerArchive").addEventListener("click", () => {
    $$(".detail-tabs button").forEach((item) => item.classList.toggle("active", item.dataset.detailTab === "archive"));
    $$(".detail-pane").forEach((pane) => pane.classList.toggle("active", pane.id === "detailArchive"));
  });
  $("#exportAbnormalList").addEventListener("click", () => exportRoadsideExcel("list"));
  $("#exportAbnormalStats").addEventListener("click", () => exportRoadsideExcel("stats"));
  window.addEventListener("amap-ready", () => {
    if (state.panel === "ops") renderOpsMap(roadsideRowsForSelectedDate());
    if (state.view === "map") renderMap($("#siteMap"), filteredSites());
  });
  window.addEventListener("resize", () => {
    applySiteViewLayoutSettled();
    if (state.activeSiteFilter) renderSiteFilterPopover();
    if (state.activeMapAssetFilter) renderMapAssetFilterPopover();
  });
}

async function init() {
  state.visualTheme = loadVisualTheme();
  loadContractReviewState();
  applyVisualTheme(state.visualTheme);
  await loadRoadsideStatusState();
  renderPageHeader();
  renderMetrics();
  renderStatusProgress();
  renderContractStrip();
  renderOverviewMap();
  renderFilters();
  renderSites();
  renderDevices();
  renderContracts();
  renderWarehouse();
  renderImportCenter();
  renderCoordinateIssues();
  renderDocuments();
  renderOps();
  bindEvents();
  loadMapAssets();
  loadDocumentAssets();
  loadContractRelationships();
}

init();
