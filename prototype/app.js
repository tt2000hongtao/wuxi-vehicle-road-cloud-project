const state = {
  panel: "overview",
  view: "table",
  query: "",
  district: "全部区域",
  perception: "全部感知类型",
  adaptive: "全部自适应",
  variableLane: "全部可变车道",
  mapNodeIdWarningBlink: false,
  selectedSite: null,
  importType: null,
  importTask: null,
  importHistory: [],
  opsDate: "2026-05-18",
  opsCalendarMode: "month",
  opsCalendarCursor: "2026-05",
  visualTheme: "command",
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
  contractManualConfirmations: {},
  contractPageScale: 1,
  contractRelationships: {
    loading: true,
    error: "",
    summary: null,
    macroFlows: [],
    contracts: [],
    frontContractItems: [],
    backContractItems: [],
    contractToMacroFlowMatches: [],
    frontBackRelationshipCandidates: [],
    deviceItemMatchCandidates: [],
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
const CONTRACT_MANUAL_CONFIRMATIONS_STORAGE_KEY = "wuxi-contract-manual-confirmations-v1";
const CONTRACT_PAGE_SCALE_STORAGE_KEY = "wuxi-contract-page-scale-v1";
const CURRENT_PANEL_STORAGE_KEY = "wuxi-current-panel-v1";
const SIDEBAR_EXPANDED_STORAGE_KEY = "wuxi-sidebar-expanded-v1";
const visualThemes = ["command", "trajectory"];
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
  contracts: ["合同管理", "基于前向销售合同、后向采购合同和 Word 明细表，重构设备级资金关联关系"],
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

function splitMacroFlowPackageName(packageName) {
  const value = String(packageName || "").trim();
  const match = value.match(/^(.*?)[（(]([^（）()]*)[）)]\s*$/);
  if (!match) return { title: value || "-", content: "-" };
  return {
    title: match[1].trim() || value,
    content: match[2].trim() || "-",
  };
}

function extractTaxRatesFromText(value) {
  return Array.from(String(value || "").matchAll(/税率\s*(?:为|：|:)?\s*(\d{1,2}%)/g)).map((match) => match[1]);
}

const contractTaxRateDisplayOrder = ["13%", "9%", "6%"];

function sortTaxRates(rates) {
  const allowed = new Set(contractTaxRateDisplayOrder);
  return Array.from(new Set(rates.filter((rate) => allowed.has(rate)))).sort(
    (a, b) => contractTaxRateDisplayOrder.indexOf(a) - contractTaxRateDisplayOrder.indexOf(b),
  );
}

function macroFlowTaxRateDisplay(flow, cr) {
  const rates = [];
  (cr.deviceItemMatchCandidates || []).forEach((candidate) => {
    if ((candidate.sharedMacroFlowIds || []).includes(flow.id)) rates.push(candidate.frontTaxRate);
  });
  const frontContractIds = new Set(
    (cr.contractToMacroFlowMatches || [])
      .filter((item) => item.direction === "front_sales" && (item.matches || []).some((match) => match.macroFlowId === flow.id))
      .map((item) => item.contractId),
  );
  if (frontContractIds.size) {
    (cr.frontContractItems || []).forEach((item) => {
      if (!frontContractIds.has(item.contractId)) return;
      rates.push(item.taxRate);
      rates.push(...extractTaxRatesFromText([item.itemNo, item.itemName, item.detailName, item.specModel, item.unit, ...(item.rawCells || [])].join(" ")));
    });
  }
  return sortTaxRates(rates).join(" / ") || "待确认";
}

function macroFlowFrontContractNames(flow, cr) {
  const names = (cr.contractToMacroFlowMatches || [])
    .filter((item) => item.direction === "front_sales" && (item.matches || []).some((match) => match.macroFlowId === flow.id))
    .map((item) => item.contractName)
    .filter(Boolean);
  return Array.from(new Set(names));
}

function frontContractAmountYuan(contractId, cr) {
  const itemSum = (cr.frontContractItems || [])
    .filter((item) => item.contractId === contractId)
    .reduce((sum, item) => sum + (Number(item.amountTaxIncluded) || 0), 0);
  if (itemSum > 0) return itemSum;
  const doc = (cr.contracts || []).find((item) => item.id === contractId);
  return Number(doc?.amounts?.[0]) || 0;
}

function macroFlowEntryAmountYuan(flow, contractId, cr) {
  const flowAmount = Number(flow?.tiananAmount || flow?.packageAmount || 0);
  if (flowAmount > 0) return flowAmount;
  return frontContractAmountYuan(contractId, cr);
}

function frontContractTaxRateDisplay(contractId, doc, cr) {
  const rates = [...(doc?.taxRates || [])];
  (cr.frontContractItems || []).forEach((item) => {
    if (item.contractId !== contractId) return;
    rates.push(item.taxRate);
    rates.push(...extractTaxRatesFromText([item.itemNo, item.itemName, item.detailName, item.specModel, item.unit, ...(item.rawCells || [])].join(" ")));
  });
  (cr.deviceItemMatchCandidates || []).forEach((candidate) => {
    if (candidate.frontContractId === contractId) rates.push(candidate.frontTaxRate);
  });
  return sortTaxRates(rates).join(" / ") || "待确认";
}

function frontContractContentDisplay(doc, fallbackContent) {
  const keywords = (doc?.keywords || []).filter(Boolean).slice(0, 8);
  return keywords.length ? keywords.join(" / ") : fallbackContent || "-";
}

function macroFlowMasterContractNote(flow) {
  const amount = Number(flow?.sourceMasterContractAmountCny || 0);
  if (!flow?.sourceMasterContractName || amount <= 0) return "";
  return `源自联合体总合同 ${formatAccountingYuan(amount)} 元`;
}

function macroFlowFrontContracts(flow, cr, fallbackContent) {
  const docById = new Map((cr.contracts || []).map((doc) => [doc.id, doc]));
  const rows = [];
  const seen = new Set();
  (cr.contractToMacroFlowMatches || [])
    .filter((item) => item.direction === "front_sales" && (item.matches || []).some((match) => match.macroFlowId === flow.id))
    .forEach((item) => {
      const key = item.contractId || item.contractName;
      if (!key || seen.has(key)) return;
      seen.add(key);
      const doc = docById.get(item.contractId);
      rows.push({
        id: key,
        fileName: item.contractName || doc?.fileName || "-",
        sourcePath: doc?.sourcePath || "",
        amount: macroFlowEntryAmountYuan(flow, item.contractId, cr),
        taxRates: frontContractTaxRateDisplay(item.contractId, doc, cr),
        content: [frontContractContentDisplay(doc, fallbackContent), macroFlowMasterContractNote(flow)].filter(Boolean).join("；"),
      });
    });
  return rows;
}

async function openContractFileWithDefaultApp(sourcePath) {
  if (!sourcePath) return;
  try {
    const response = await fetch("/api/contract-file/open", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: sourcePath }),
    });
    if (!response.ok) {
      const payload = await response.json().catch(() => null);
      const message = payload?.message || payload?.error || `HTTP ${response.status}`;
      throw new Error(message);
    }
  } catch (error) {
    alert(
      `合同文件调阅失败：${error.message || error}\n\n请确认当前页面通过 node prototype/server.js 启动的本地服务访问，而不是静态文件服务；修改后也需要重启原型服务。`,
    );
  }
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

function formatAccountingYuan(value) {
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(Number(value) || 0);
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

function loadContractManualConfirmations() {
  try {
    const raw = globalThis.localStorage?.getItem(CONTRACT_MANUAL_CONFIRMATIONS_STORAGE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    state.contractManualConfirmations = parsed?.confirmations || {};
  } catch (error) {
    state.contractManualConfirmations = {};
  }
}

function saveContractManualConfirmations() {
  try {
    globalThis.localStorage?.setItem(
      CONTRACT_MANUAL_CONFIRMATIONS_STORAGE_KEY,
      JSON.stringify({
        updatedAt: new Date().toISOString(),
        confirmations: state.contractManualConfirmations,
      }),
    );
  } catch (error) {
    // Keep in-memory confirmations if localStorage is unavailable.
  }
}

function normalizeContractPageScale(value) {
  const scale = Number(value);
  if (!Number.isFinite(scale)) return 1;
  return Math.min(1, Math.max(0.65, Math.round(scale * 100) / 100));
}

function loadContractPageScale() {
  try {
    const raw = globalThis.localStorage?.getItem(CONTRACT_PAGE_SCALE_STORAGE_KEY);
    state.contractPageScale = normalizeContractPageScale(raw || 1);
  } catch (error) {
    state.contractPageScale = 1;
  }
}

function saveContractPageScale() {
  try {
    globalThis.localStorage?.setItem(CONTRACT_PAGE_SCALE_STORAGE_KEY, String(state.contractPageScale));
  } catch (error) {
    // Keep the in-memory scale when localStorage is unavailable.
  }
}

function applyContractPageScale() {
  const surface = $('#contractZoomSurface');
  if (!surface) return;
  const scale = normalizeContractPageScale(state.contractPageScale);
  state.contractPageScale = scale;
  surface.style.setProperty('--contract-page-scale', String(scale));
  surface.style.zoom = String(scale);
  surface.classList.toggle('is-scaled', scale < 0.999);
  const value = $('#contractZoomValue');
  if (value) value.textContent = `${Math.round(scale * 100)}%`;
  requestAnimationFrame(() => renderContractManualColumnHandles());
}

function fitContractPageToActiveWorkspace() {
  const active = $('[data-contract-workspace-panel].active');
  const surface = $('#contractZoomSurface');
  const table = active?.querySelector('.contract-workspace-table-wrap table');
  if (!surface || !table) {
    setContractPageScale(0.85);
    return;
  }
  const available = Math.max(320, surface.clientWidth - 24);
  const tableWidth = Math.max(table.scrollWidth, table.offsetWidth, available);
  const fitted = normalizeContractPageScale(Math.min(1, (available / tableWidth) * 0.98));
  setContractPageScale(fitted);
}

function setContractPageScale(scale, options = {}) {
  state.contractPageScale = normalizeContractPageScale(scale);
  if (options.persist !== false) saveContractPageScale();
  applyContractPageScale();
}

function adjustContractPageScale(action) {
  if (action === 'reset') {
    setContractPageScale(1);
    return;
  }
  if (action === 'fit') {
    fitContractPageToActiveWorkspace();
    return;
  }
  const step = action === 'in' ? 0.05 : -0.05;
  setContractPageScale(state.contractPageScale + step);
}

function getContractManualConfirmation(contractId) {
  return state.contractManualConfirmations[contractId] || null;
}

function upsertContractManualConfirmation(contractId, patch) {
  if (!contractId) return;
  const current = getContractManualConfirmation(contractId) || {};
  state.contractManualConfirmations[contractId] = {
    ...current,
    ...patch,
    updatedAt: new Date().toISOString(),
  };
  saveContractManualConfirmations();
}

function resetContractManualConfirmation(contractId) {
  if (!contractId) return;
  delete state.contractManualConfirmations[contractId];
  saveContractManualConfirmations();
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
      evidence: status === "confirmed" ? "原型人工确认：前向合同明细项与后向采购明细项形成设备/服务级候选关联，仍需正式系统复核合同附件、审计口径和付款条件" : "原型人工否决：该设备/服务明细项候选不进入正式设备级资金流",
    };
  }
  state.contractReview.selectedCandidateId = candidateId;
  saveContractReviewState();
  renderContracts();
  renderMetrics();
  renderContractStrip();
}

function contractReviewStats() {
  const candidates = state.contractRelationships.deviceItemMatchCandidates || [];
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

async function rebuildContractRelationships() {
  const button = $("#refreshContractRelationships");
  const originalText = button?.textContent || "刷新候选数据";
  if (button) {
    button.disabled = true;
    button.textContent = "正在扫描并重建...";
  }
  state.contractRelationships.loading = true;
  state.contractRelationships.error = "";
  renderContracts();
  try {
    const response = await fetch("/api/contract-relationships/rebuild", { method: "POST", cache: "no-store" });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.message || payload.error || `HTTP ${response.status}`);
    await loadDocumentAssets();
    await loadContractRelationships();
  } catch (error) {
    state.contractRelationships = {
      ...state.contractRelationships,
      loading: false,
      error: error.message || "合同候选数据重建失败",
    };
    renderContracts();
    alert(`合同候选数据刷新失败：${state.contractRelationships.error}`);
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = originalText;
    }
  }
}

async function loadContractRelationships() {
  state.contractRelationships.loading = true;
  state.contractRelationships.error = "";
  renderContracts();
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
      frontContractItems: payload.frontContractItems || [],
      backContractItems: payload.backContractItems || [],
      contractToMacroFlowMatches: payload.contractToMacroFlowMatches || [],
      frontBackRelationshipCandidates: payload.frontBackRelationshipCandidates || [],
      deviceItemMatchCandidates: payload.deviceItemMatchCandidates || [],
      deviceCashflowSchema: payload.deviceCashflowSchema || null,
      ownerAuditTrackingModel: payload.ownerAuditTrackingModel || null,
      paymentStageTemplates: payload.paymentStageTemplates || [],
      selectedMacroFlowId: state.contractRelationships.selectedMacroFlowId || "",
    };
  } catch (error) {
    state.contractRelationships = {
      ...state.contractRelationships,
      loading: false,
      error: error.message || "合同关系重构数据读取失败",
      summary: null,
      macroFlows: [],
      contracts: [],
      frontContractItems: [],
      backContractItems: [],
      contractToMacroFlowMatches: [],
      frontBackRelationshipCandidates: [],
      deviceItemMatchCandidates: [],
      deviceCashflowSchema: null,
      ownerAuditTrackingModel: null,
      paymentStageTemplates: [],
      selectedMacroFlowId: "",
    };
  }
  renderContracts();
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

function mapAssetForSite(site) {
  const nodeId = String(site?.nodeId || "").trim();
  if (!nodeId) return null;
  return mapAssetState.intersections.find((asset) => String(asset.nodeId || "").trim() === nodeId) || null;
}

function siteHasMapXmlNodeIdWarning(site) {
  const asset = mapAssetForSite(site);
  if (!asset) return false;
  const missingItems = Array.isArray(asset.missing) ? asset.missing : [];
  return Boolean(
    asset.quality?.mapXmlNodeIdShort ||
    missingItems.some((item) => String(item || "").includes("MAP-XML上下游 NodeID<5位"))
  );
}

function renderSiteMap(rows = filteredSites()) {
  renderMap($("#siteMap"), rows, {
    nodeIdWarningBlink: state.mapNodeIdWarningBlink,
  });
}

function renderDrawerMapXmlPreview(site) {
  const preview = $("#drawerMapXmlPreview");
  const meta = $("#drawerMapXmlMeta");
  if (!preview || !meta) return;
  sizeDrawerMapXmlPreview();
  const asset = mapAssetForSite(site);
  if (!asset) {
    meta.textContent = mapAssetState.loading ? "地图资产扫描中" : "未匹配地图资产";
    preview.innerHTML = `<div class="empty-detail">${mapAssetState.loading ? "正在扫描地图资产目录" : "未找到该点位的 MAP-XML SVG 预览"}</div>`;
    return;
  }
  const svgFile = mapAssetFileByPath(asset, asset.mapXmlSvgPath);
  if (!asset.mapXmlSvgPath) {
    meta.textContent = "未找到 map_xml SVG";
    preview.innerHTML = `<div class="empty-detail">该点位 map_xml 目录没有 SVG 预览文件</div>`;
    return;
  }
  const displayName = mapAssetDisplayName(asset);
  const svgUrl = `/api/map-assets/file?path=${encodeURIComponent(asset.mapXmlSvgPath)}`;
  meta.textContent = svgFile ? `由 ${svgFile.fileName} 生成` : "map_xml 目录 SVG";
  preview.innerHTML = `<img src="${svgUrl}" alt="${displayName} MAP-XML 转换 SVG 高精地图" />`;
  requestAnimationFrame(sizeDrawerMapXmlPreview);
}

async function ensureDrawerMapXmlPreview(site) {
  renderDrawerMapXmlPreview(site);
  if (!mapAssetForSite(site) && !mapAssetState.loading) {
    await loadMapAssets(false);
    renderDrawerMapXmlPreview(site);
  }
}

function sizeDrawerMapXmlPreview() {
  const preview = $("#drawerMapXmlPreview");
  const panel = $(".drawer-map-xml-panel");
  if (!preview || !panel) return;
  const panelRect = panel.getBoundingClientRect();
  const headerHeight = panel.querySelector(".drawer-panel-head")?.getBoundingClientRect().height || 42;
  const availableHeight = Math.max(240, panelRect.height - headerHeight - 14);
  const availableWidth = Math.max(240, panelRect.width - 2);
  const size = Math.floor(Math.min(availableWidth, availableHeight));
  preview.style.width = `${size}px`;
  preview.style.height = `${size}px`;
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

function validPanelIds() {
  return new Set($$('.panel').map((panel) => panel.id).filter(Boolean));
}

function normalizePanelId(panelId) {
  const text = String(panelId || '').replace(/^#/, '').trim();
  return validPanelIds().has(text) ? text : 'overview';
}

function panelIdFromLocation() {
  try {
    return normalizePanelId(decodeURIComponent(globalThis.location?.hash || '').replace(/^#/, ''));
  } catch (error) {
    return 'overview';
  }
}

function loadSavedPanelId() {
  try {
    const fromHash = panelIdFromLocation();
    if (fromHash !== 'overview' || globalThis.location?.hash) return fromHash;
    return normalizePanelId(globalThis.localStorage?.getItem(CURRENT_PANEL_STORAGE_KEY));
  } catch (error) {
    return 'overview';
  }
}

function persistCurrentPanel(panelId, options = {}) {
  try {
    globalThis.localStorage?.setItem(CURRENT_PANEL_STORAGE_KEY, panelId);
  } catch (error) {
    // Keep navigation usable when storage is unavailable.
  }
  if (options.updateHash === false) return;
  const nextHash = `#${encodeURIComponent(panelId)}`;
  if (globalThis.location?.hash !== nextHash) {
    globalThis.history?.replaceState(null, '', nextHash);
  }
}

function setPanel(panelId, options = {}) {
  panelId = normalizePanelId(panelId);
  state.panel = panelId;
  persistCurrentPanel(panelId, options);
  $$(".panel").forEach((panel) => panel.classList.toggle("active", panel.id === panelId));
  $$(".nav-item").forEach((item) => item.classList.toggle("active", item.dataset.panel === panelId));
  renderPageHeader();
  if (panelId === "overview") requestAnimationFrame(renderOverviewMap);
  if (panelId === "sites") requestAnimationFrame(applySiteViewLayoutSettled);
  if (panelId === "sites" && state.view === "map") requestAnimationFrame(() => { applySiteViewLayoutSettled(); renderSiteMap(); });
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
    if (state.panel === "sites" && state.view === "map") {
      renderSiteMap();
    }
    if ($("#detailDrawer")?.classList.contains("open") && state.selectedSite) {
      renderDrawerMapXmlPreview(state.selectedSite);
    }
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

function renderContractManualColumnHandles(table = null) {
  const tables = table ? [table] : Array.from(document.querySelectorAll(".contract-manual-table"));
  tables.forEach((currentTable) => {
    currentTable.querySelectorAll(".contract-manual-column-resizer").forEach((handle) => handle.remove());
    Array.from(currentTable.querySelectorAll("thead th")).forEach((th, index, headers) => {
      if (index === headers.length - 1) return;
      const handle = document.createElement("button");
      handle.type = "button";
      handle.className = "column-resizer contract-manual-column-resizer";
      handle.setAttribute("aria-label", `调整 ${th.textContent.trim()} 列宽`);
      handle.dataset.columnIndex = String(index);
      handle.dataset.resizeTable = "contract-manual";
      th.appendChild(handle);
    });
  });
}

function startContractManualColumnResize(event, handle) {
  event.preventDefault();
  event.stopPropagation();
  const table = handle.closest("table");
  const th = handle.closest("th");
  if (!table || !th) return;
  const columnIndex = Number(handle.dataset.columnIndex);
  const startX = event.clientX;
  const startWidth = th.getBoundingClientRect().width;
  const cells = Array.from(table.querySelectorAll(`tr > *:nth-child(${columnIndex + 1})`));
  const onMove = (moveEvent) => {
    const minWidth = columnIndex === 0 ? 58 : 86;
    const width = Math.max(minWidth, startWidth + moveEvent.clientX - startX);
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
    return visualThemes.includes(saved) ? saved : "command";
  } catch {
    return "command";
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
  state.visualTheme = visualThemes.includes(theme) ? theme : "command";
  document.body.dataset.visualTheme = state.visualTheme;
  const toggle = $("#visualThemeToggle");
  if (!toggle) return;
  const themeLabels = {
    command: "大屏风格",
    trajectory: "轨迹风格",
  };
  const themeTitles = {
    command: "切换为轨迹实时指标风格",
    trajectory: "切换为大屏指挥风格",
  };
  toggle.textContent = themeLabels[state.visualTheme];
  toggle.setAttribute("aria-pressed", String(state.visualTheme === "trajectory"));
  toggle.title = themeTitles[state.visualTheme];
}

function toggleVisualTheme() {
  const currentIndex = visualThemes.indexOf(state.visualTheme);
  applyVisualTheme(visualThemes[(currentIndex + 1) % visualThemes.length]);
  saveVisualTheme(state.visualTheme);
}

function loadSidebarExpanded() {
  try {
    return localStorage.getItem(SIDEBAR_EXPANDED_STORAGE_KEY) === "true";
  } catch {
    return false;
  }
}

function saveSidebarExpanded(expanded) {
  try {
    localStorage.setItem(SIDEBAR_EXPANDED_STORAGE_KEY, String(expanded));
  } catch {
    // Storage can be unavailable in restricted browser contexts; the live toggle still works.
  }
}

function applySidebarExpanded(expanded = false) {
  document.body.classList.toggle("sidebar-expanded", Boolean(expanded));
  const toggle = $("#sidebarWidthToggle");
  if (!toggle) return;
  toggle.setAttribute("aria-expanded", String(expanded));
  toggle.setAttribute("aria-label", expanded ? "缩进左侧菜单" : "延展左侧菜单");
  toggle.title = expanded ? "缩进左侧菜单" : "延展左侧菜单";
}

function toggleSidebarExpanded() {
  const expanded = !document.body.classList.contains("sidebar-expanded");
  applySidebarExpanded(expanded);
  saveSidebarExpanded(expanded);
  requestAnimationFrame(applySiteViewLayoutSettled);
}

function currentFullscreenElement() {
  return document.fullscreenElement || document.webkitFullscreenElement || null;
}

function renderFullscreenToggle() {
  const toggle = $("#fullscreenToggle");
  if (!toggle) return;
  const isFullscreen = Boolean(currentFullscreenElement());
  toggle.textContent = isFullscreen ? "退出全屏" : "全屏";
  toggle.setAttribute("aria-pressed", String(isFullscreen));
  toggle.title = isFullscreen ? "退出全屏显示" : "进入全屏显示";
}

async function toggleFullscreen() {
  try {
    if (currentFullscreenElement()) {
      const exitFullscreen = document.exitFullscreen || document.webkitExitFullscreen;
      if (exitFullscreen) await exitFullscreen.call(document);
    } else {
      const target = document.documentElement;
      const requestFullscreen = target.requestFullscreen || target.webkitRequestFullscreen;
      if (requestFullscreen) await requestFullscreen.call(target);
    }
  } catch (error) {
    console.warn("Fullscreen toggle failed.", error);
  } finally {
    renderFullscreenToggle();
  }
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
    const candidates = summary.deviceItemMatchCandidateCount ?? cr.deviceItemMatchCandidates.length;
    contractMetricHint.textContent = cr.loading ? "读取合同关系" : `${formatNumber(candidates)} 条设备级候选`;
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
      <div class="contract-row contract-row-muted"><div><b>正在读取合同关系重构数据</b><span>读取前向销售合同、后向采购合同和 Word 明细表</span></div><span class="status-pill warn">读取中</span></div>
      <div class="contract-row contract-row-muted"><div><b>设备级资金流闸门</b><span>未确认的设备/服务明细关系不会进入付款或毛利计算</span></div><span class="status-pill warn">锁定</span></div>`;
    return;
  }
  if (cr.error) {
    holder.innerHTML = `<div class="contract-row contract-row-muted"><div><b>合同关系数据读取失败</b><span>${escapeHtml(cr.error)}</span></div><span class="status-pill error">失败</span></div>`;
    return;
  }
  const rows = [
    ["前向销售合同", `${formatNumber(summary.macroFlowCount ?? cr.macroFlows.length)} 个前向资金入口，${formatNumber(summary.frontContractItemCount || cr.frontContractItems.length || 0)} 条前向明细`, "前向", "done"],
    ["后向采购合同", `${formatNumber(summary.backContractCount || 0)} 份后向合同，${formatNumber(summary.backContractItemCount || cr.backContractItems.length || 0)} 条后向明细`, "后向", "done"],
    ["设备级逐条确认", `${formatNumber(summary.deviceItemMatchCandidateCount ?? cr.deviceItemMatchCandidates.length)} 条候选；${formatNumber(reviewStats.pending)} 条待确认`, "待确认", reviewStats.pending ? "warn" : "done"],
    ["资金计算闸门", "只有前向明细与后向明细逐条确认后，才能生成设备级资金流单元；未确认关系保持锁定", "锁定", "warn"],
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
    const hasMapXmlNodeIdWarning = siteHasMapXmlNodeIdWarning(site);
    const shouldBlinkNodeIdWarning = options.nodeIdWarningBlink && hasMapXmlNodeIdWarning;
    const titlePrefix = isCurrent ? "当前点位：" : "周边点位：";
    const warningTitle = hasMapXmlNodeIdWarning ? "（MAP-XML上下游 NodeID 位数小于5位）" : "";
    const marker = document.createElement("div");
    marker.className = `marker ${districtClass(site)} ${isCurrent ? "current" : "neighbor"} ${site.issueCount ? "has-issue" : ""} ${shouldBlinkNodeIdWarning ? "map-nodeid-warning" : ""}`;
    marker.style.left = `${pos.left}%`;
    marker.style.top = `${pos.top}%`;
    marker.title = `${titlePrefix}${site.nodeId} ${site.name}${warningTitle}`;
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
    const hasMapXmlNodeIdWarning = siteHasMapXmlNodeIdWarning(site);
    const shouldBlinkNodeIdWarning = options.nodeIdWarningBlink && hasMapXmlNodeIdWarning;
    const warningTitle = hasMapXmlNodeIdWarning ? "（MAP-XML上下游 NodeID 位数小于5位）" : "";
    const marker = new window.AMap.Marker({
      position: [site.lngGcj, site.latGcj],
      title: `${isCurrent ? "当前点位：" : "周边点位："}${site.nodeId} ${site.name}${warningTitle}`,
      offset: new window.AMap.Pixel(isCurrent ? -13 : -7, isCurrent ? -13 : -7),
      zIndex: isCurrent ? 120 : 100,
      content: `<button class="amap-site-marker ${districtClass(site)} ${isCurrent ? "current" : "neighbor"} ${site.issueCount ? "has-issue" : ""} ${shouldBlinkNodeIdWarning ? "map-nodeid-warning" : ""}" title="${site.name}${warningTitle}" data-node="${site.nodeId}">${isCurrent ? "当前" : ""}</button>`,
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
    renderSiteMap(rows);
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
  renderContractManualWorkspaces();
  applyContractPageScale();
}

function setContractWorkspace(workspace) {
  if (!workspace) return;
  $$('[data-contract-workspace-panel]').forEach((panel) => {
    panel.classList.toggle('active', panel.dataset.contractWorkspacePanel === workspace);
  });
  $$('.contract-workspace-tabs [data-contract-workspace]').forEach((button) => {
    button.classList.toggle('active', button.dataset.contractWorkspace === workspace);
  });
  requestAnimationFrame(applyContractPageScale);
}

function contractManualStatusText() {
  return '<span class="status-pill warn">候选 / 待人工确认</span>';
}

function allowedTaxRateDisplay(rates) {
  const allowed = ['13%', '9%', '6%'];
  const list = [...new Set((rates || []).filter((rate) => allowed.includes(String(rate).trim())))];
  return list.length ? list.join(' / ') : '待确认';
}

function itemDisplayName(item) {
  return [item?.itemName, item?.detailName, item?.specModel].filter(Boolean).join(' / ') || '未命名清单行';
}

function sumAmount(items, field = 'amountTaxIncluded') {
  return (items || []).reduce((sum, item) => sum + (Number(item?.[field]) || 0), 0);
}

function contractDocumentById(contractId) {
  return (state.contractRelationships.contracts || []).find((contract) => contract.id === contractId) || null;
}

function candidateMacroFlowAmountForContract(contractId) {
  const match = (state.contractRelationships.contractToMacroFlowMatches || []).find((item) => item.contractId === contractId);
  const top = match?.matches?.[0];
  if (!top?.macroFlowId) return 0;
  const flow = (state.contractRelationships.macroFlows || []).find((item) => item.id === top.macroFlowId);
  return Number(flow?.tiananAmount || flow?.packageAmount || 0) || 0;
}

function candidateAmountForGroup(group) {
  const macroFlowAmount = candidateMacroFlowAmountForContract(group.contractId);
  if (macroFlowAmount > 0) return macroFlowAmount;
  const itemAmount = sumAmount(group.items);
  if (itemAmount > 0) return itemAmount;
  const doc = contractDocumentById(group.contractId);
  const amounts = (doc?.amounts || []).map(Number).filter((value) => Number.isFinite(value) && value > 0);
  if (!amounts.length) return 0;
  return Math.min(...amounts.filter((value) => value >= 10000)) || Math.max(...amounts);
}

function manualAmountForContractGroup(group) {
  const candidateAmount = candidateAmountForGroup(group);
  const confirmation = getContractManualConfirmation(group.contractId) || {};
  const hasSavedAmount = confirmation.amount !== undefined && confirmation.amount !== null && String(confirmation.amount).trim() !== '';
  const savedAmount = hasSavedAmount ? Number(confirmation.amount) : NaN;
  return Number.isFinite(savedAmount) ? savedAmount : candidateAmount;
}

function confirmedAmountForContractGroup(group) {
  const confirmation = getContractManualConfirmation(group.contractId) || {};
  if (confirmation.status !== 'confirmed') return 0;
  const hasSavedAmount = confirmation.amount !== undefined && confirmation.amount !== null && String(confirmation.amount).trim() !== '';
  if (!hasSavedAmount) return 0;
  const savedAmount = Number(confirmation.amount);
  return Number.isFinite(savedAmount) ? savedAmount : 0;
}

function subtotalAmountForContractFlowGroup(flowGroup) {
  return (flowGroup.contracts || []).reduce((sum, group) => sum + confirmedAmountForContractGroup(group), 0);
}

function totalAmountForContractGroups(groups) {
  return (groups || []).reduce((sum, group) => sum + confirmedAmountForContractGroup(group), 0);
}

function confirmedManualContractGroups(groups) {
  return (groups || []).filter((group) => {
    const confirmation = getContractManualConfirmation(group.contractId) || {};
    return confirmation.status === 'confirmed';
  });
}

function totalAmountForConfirmedContractGroups(groups) {
  return totalAmountForContractGroups(confirmedManualContractGroups(groups));
}

function confirmedFrontAmountByMacroFlow(groups) {
  const result = new Map();
  confirmedManualContractGroups(groups).forEach((group) => {
    const confirmation = getContractManualConfirmation(group.contractId) || {};
    const flowId = confirmation.macroFlowId || '';
    if (!flowId) return;
    result.set(flowId, (result.get(flowId) || 0) + confirmedAmountForContractGroup(group));
  });
  return result;
}

function candidateTaxRatesForGroup(group) {
  const allowed = ['13%', '9%', '6%'];
  const doc = contractDocumentById(group.contractId);
  return [...new Set([...(group.items || []).map((item) => item.taxRate), ...(doc?.taxRates || [])]
    .map((rate) => String(rate || '').trim())
    .filter((rate) => allowed.includes(rate)))];
}

function groupItemsByContract(items) {
  const grouped = new Map();
  (items || []).forEach((item) => {
    const id = item.contractId || item.contractName || 'unknown';
    if (!grouped.has(id)) {
      grouped.set(id, {
        contractId: id,
        contractName: item.contractName || '-',
        counterparty: item.counterparty || '-',
        sourcePath: item.sourcePath || '',
        items: [],
      });
    }
    grouped.get(id).items.push(item);
  });
  return [...grouped.values()];
}

function normalizedContractPartyName(value) {
  const text = String(value || '').trim();
  if (!text || text === '-' || text.length > 36) return '';
  if (/^V?\d+(?:\.\d+)*$/i.test(text) || /^终改\d*$/i.test(text)) return '';
  if (/合同|协议|项目|采购|补充|一体化|应用试点/.test(text) && !/(公司|集团|研究所|天安|移动|华通|万集|浪潮|尚行|车城|四维|大华|海康|希迪|大为|交科所|金中天|隆顺|网盈|合创|工业安装|中电鸿信)/.test(text)) return '';
  return text;
}

function contractSupplierName(contract, itemGroup) {
  const confirmation = getContractManualConfirmation(contract?.id || itemGroup?.contractId) || {};
  const manualSupplierName = String(confirmation.supplierName || '').trim();
  if (manualSupplierName) return manualSupplierName;
  const partyBFullName = String(contract?.partyBFullName || '').trim();
  if (partyBFullName) return partyBFullName;
  const parties = Array.isArray(contract?.parties) ? contract.parties : [];
  const sourcePath = String(contract?.sourcePath || itemGroup?.sourcePath || '');
  const fileName = String(contract?.fileName || itemGroup?.contractName || '');
  const knownParties = ['中电鸿信', '交科所', '华通', '金中天', '隆顺', '工业安装', '合创', '车城', '尚行', '浪潮', '四维图新', '中通服网盈', '万集', '大华', '海康智联', '希迪', '大为'];
  const pathParts = sourcePath.split(/[\/]/).filter(Boolean);
  const backContractIndex = pathParts.findIndex((part) => part === '后向采购合同');
  const explicitScopeParts = [fileName];
  if (backContractIndex >= 0) explicitScopeParts.push(...pathParts.slice(backContractIndex + 1));
  else explicitScopeParts.push(...pathParts.slice(-3));
  for (const part of explicitScopeParts) {
    const hit = knownParties.find((name) => part.includes(name));
    if (hit) return hit;
  }
  const preferred = parties
    .map(normalizedContractPartyName)
    .filter(Boolean)
    .filter((name) => !['天安', '天安智联', '中国移动', '移动', '车联网', '车联网集团', '无锡车联网集团', '无锡市车联网产业集团'].includes(name));
  if (preferred.length === 1) return preferred[0];
  const counterparty = normalizedContractPartyName(contract?.counterparty || itemGroup?.counterparty);
  if (counterparty) return counterparty;
  return '-';
}

function groupContractsWithItems(direction, items) {
  const itemGroups = new Map(groupItemsByContract(items).map((group) => [group.contractId, group]));
  return (state.contractRelationships.contracts || [])
    .filter((contract) => contract.direction === direction)
    .map((contract) => {
      const itemGroup = itemGroups.get(contract.id);
      return {
        contractId: contract.id,
        contractName: contract.fileName || itemGroup?.contractName || '-',
        counterparty: contract.counterparty || itemGroup?.counterparty || '-',
        supplierName: contractSupplierName(contract, itemGroup),
        partyAFullName: contract.partyAFullName || '',
        partyBFullName: contract.partyBFullName || '',
        parties: contract.parties || [],
        sourcePath: contract.sourcePath || itemGroup?.sourcePath || '',
        direction: contract.direction,
        documentAmounts: contract.amounts || [],
        documentTaxRates: contract.taxRates || [],
        items: itemGroup?.items || [],
        hasLineItems: Boolean(itemGroup?.items?.length),
      };
    });
}

function renderContractKpis(target, cards) {
  const node = $(target);
  if (!node) return;
  node.innerHTML = cards
    .map(
      ([label, value, note = '', cls = '', action = '']) => `
        <article class="metric contract-workspace-kpi ${cls} ${action ? 'clickable' : ''}" ${action ? `role="button" tabindex="0" data-contract-kpi-action="${escapeHtml(action)}"` : ''}>
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(formatMetricValue(value))}</strong>
          ${note ? `<small>${escapeHtml(note)}</small>` : ''}
        </article>
      `,
    )
    .join('');
}

function contractFileButton(name, path) {
  if (!path) return escapeHtml(name || '-');
  return `<button type="button" class="contract-file-link" data-contract-open-file="${escapeHtml(path)}">${escapeHtml(name || '-')}</button>`;
}

function macroFlowSummaryText(flow) {
  const packageName = String(flow?.packageName || '');
  const match = packageName.match(/[（(]([^（）()]*)[）)]\s*$/);
  return match?.[1]?.trim() || String(flow?.notes || '').replace(String(flow?.displayText || ''), '').trim() || '待补充设备级清单';
}

function macroFlowParties(flow) {
  const display = String(flow?.displayText || '').trim();
  if (display) return display.split(/\s*→\s*/).map((item) => item.trim()).filter(Boolean);
  const packageName = String(flow?.packageName || '').replace(/[（(].*$/, '').trim();
  return packageName.split('-').map((item) => item.trim()).filter(Boolean);
}

function renderFrontSalesFlowTreeDiagram(flows, amountByFlow = new Map()) {
  const flowMeta = (flows || []).map((flow) => {
    const parties = macroFlowParties(flow);
    const directParty = parties.slice(0, -1).at(-1) || '';
    return {
      flow,
      parties,
      directParty,
      summary: macroFlowSummaryText(flow),
      amount: Number(amountByFlow.get(flow.id) || 0),
    };
  });
  const directByParty = new Map(flowMeta.map((item) => [item.directParty, item]));
  const directMeta = (party) => directByParty.get(party) || { summary: '', amount: 0 };
  const amountLabel = (party) => {
    const amount = Number(directMeta(party).amount || 0);
    return amount > 0 ? formatAccountingYuan(amount) : '待确认';
  };
  const summaryLabel = (party) => directMeta(party).summary || '待补充';
  const nodes = [
    { id: 'owner', label: ['业主', '无锡市车联网产业集团'], x: 36, y: 282, w: 320, h: 108, cls: 'owner hero' },
    { id: 'vehicle', label: ['车联网集团', amountLabel('车联网') || amountLabel('车联网集团')], x: 640, y: 36, w: 360, h: 74, cls: 'main direct wide' },
    { id: 'mobile', label: ['中国移动', '联合体成员'], x: 500, y: 330, w: 166, h: 104, cls: 'main mobile' },
    { id: 'langchao', label: ['浪潮', amountLabel('浪潮')], x: 786, y: 176, w: 214, h: 66, cls: 'mid direct' },
    { id: 'shangyan', label: ['上研'], x: 708, y: 292, w: 88, h: 60, cls: 'mid small' },
    { id: 'hechuang', label: ['合创', amountLabel('合创')], x: 846, y: 260, w: 154, h: 60, cls: 'mid direct compact' },
    { id: 'gongye', label: ['工业安装', amountLabel('工业安装')], x: 846, y: 344, w: 154, h: 60, cls: 'mid direct compact' },
    { id: 'wangying', label: ['中通服网盈', amountLabel('中通服网盈')], x: 786, y: 426, w: 214, h: 66, cls: 'mid direct' },
    { id: 'shangxing', label: ['尚行', amountLabel('尚行')], x: 786, y: 516, w: 214, h: 66, cls: 'mid direct' },
    { id: 'siwei', label: ['四维图新', amountLabel('四维图新')], x: 786, y: 606, w: 214, h: 66, cls: 'mid direct' },
    { id: 'sumVehicle', label: [summaryLabel('车联网') || summaryLabel('车联网集团')], x: 1082, y: 44, w: 242, h: 58, cls: 'summary' },
    { id: 'sumLangchao', label: [summaryLabel('浪潮')], x: 1082, y: 180, w: 242, h: 58, cls: 'summary' },
    { id: 'sumHechuang', label: [summaryLabel('合创')], x: 1082, y: 262, w: 242, h: 58, cls: 'summary' },
    { id: 'sumGongye', label: [summaryLabel('工业安装')], x: 1082, y: 346, w: 242, h: 58, cls: 'summary' },
    { id: 'sumWangying', label: [summaryLabel('中通服网盈')], x: 1082, y: 430, w: 242, h: 58, cls: 'summary' },
    { id: 'sumShangxing', label: [summaryLabel('尚行')], x: 1082, y: 520, w: 242, h: 58, cls: 'summary' },
    { id: 'sumSiwei', label: [summaryLabel('四维图新')], x: 1082, y: 610, w: 242, h: 58, cls: 'summary' },
    { id: 'tianan', label: ['天安智联'], x: 1480, y: 352, w: 132, h: 64, cls: 'tianan final' },
  ];
  const nodeById = new Map(nodes.map((node) => [node.id, node]));
  const right = (id) => {
    const n = nodeById.get(id);
    return { x: n.x + n.w, y: n.y + n.h / 2 };
  };
  const left = (id) => {
    const n = nodeById.get(id);
    return { x: n.x, y: n.y + n.h / 2 };
  };
  const polyline = (points, cls = '') => `<polyline class="${cls}" points="${points.map((pt) => `${pt.x},${pt.y}`).join(' ')}"></polyline>`;
  const ownerRight = right('owner');
  const mobileRight = right('mobile');
  const mobileBusX = 688;
  const shangyanRight = right('shangyan');
  const shangyanBusX = 822;
  const bracketX = 1406;
  const bracketTopY = left('sumVehicle').y;
  const bracketBottomY = left('sumSiwei').y;
  const tiananLeft = left('tianan');
  const summaryIds = ['sumVehicle', 'sumLangchao', 'sumHechuang', 'sumGongye', 'sumWangying', 'sumShangxing', 'sumSiwei'];
  const edges = [
    polyline([ownerRight, { x: 382, y: ownerRight.y }, { x: 382, y: left('vehicle').y }, left('vehicle')], 'owner-edge'),
    polyline([ownerRight, { x: 382, y: ownerRight.y }, { x: 382, y: left('mobile').y }, left('mobile')], 'owner-edge'),
    polyline([mobileRight, { x: mobileBusX, y: mobileRight.y }, { x: mobileBusX, y: left('langchao').y }, left('langchao')]),
    polyline([mobileRight, { x: mobileBusX, y: mobileRight.y }, { x: mobileBusX, y: left('shangyan').y }, left('shangyan')]),
    polyline([mobileRight, { x: mobileBusX, y: mobileRight.y }, { x: mobileBusX, y: left('wangying').y }, left('wangying')]),
    polyline([mobileRight, { x: mobileBusX, y: mobileRight.y }, { x: mobileBusX, y: left('shangxing').y }, left('shangxing')]),
    polyline([mobileRight, { x: mobileBusX, y: mobileRight.y }, { x: mobileBusX, y: left('siwei').y }, left('siwei')]),
    polyline([shangyanRight, { x: shangyanBusX, y: shangyanRight.y }, { x: shangyanBusX, y: left('hechuang').y }, left('hechuang')]),
    polyline([shangyanRight, { x: shangyanBusX, y: shangyanRight.y }, { x: shangyanBusX, y: left('gongye').y }, left('gongye')]),
    polyline([right('vehicle'), left('sumVehicle')], 'summary-edge'),
    polyline([right('langchao'), left('sumLangchao')], 'summary-edge'),
    polyline([right('hechuang'), left('sumHechuang')], 'summary-edge'),
    polyline([right('gongye'), left('sumGongye')], 'summary-edge'),
    polyline([right('wangying'), left('sumWangying')], 'summary-edge'),
    polyline([right('shangxing'), left('sumShangxing')], 'summary-edge'),
    polyline([right('siwei'), left('sumSiwei')], 'summary-edge'),
    ...summaryIds.map((id) => polyline([right(id), { x: bracketX, y: right(id).y }], 'bracket-feed')),
    polyline([{ x: bracketX, y: bracketTopY }, { x: bracketX, y: bracketBottomY }], 'right-bracket'),
    polyline([{ x: bracketX, y: tiananLeft.y }, tiananLeft], 'final-edge'),
  ];
  return `
    <div class="front-flow-tree-wrap">
      <svg class="front-flow-tree-svg xmind screenshot-structure" viewBox="0 0 1660 728" role="img" aria-label="前向合同流链路树状汇聚框图">
        <defs>
          <marker id="frontFlowArrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto" markerUnits="strokeWidth">
            <path d="M0,0 L0,6 L9,3 z"></path>
          </marker>
        </defs>
        <g class="front-flow-svg-edges">${edges.join('')}</g>
        <g class="front-flow-svg-nodes">
          ${nodes.map((node) => `
            <g class="front-flow-svg-node ${node.cls}">
              <rect x="${node.x}" y="${node.y}" width="${node.w}" height="${node.h}" rx="16"></rect>
              ${node.label.map((line, lineIndex) => {
                const lineGap = node.cls.includes('hero') ? 30 : node.cls.includes('direct') ? 26 : 20;
                const y = node.y + node.h / 2 + (lineIndex - (node.label.length - 1) / 2) * lineGap + 6;
                return `<text class="${lineIndex === 1 && node.cls.includes('direct') ? 'node-amount' : ''}" x="${node.x + node.w / 2}" y="${y}" text-anchor="middle">${escapeHtml(line)}</text>`;
              }).join('')}
            </g>
          `).join('')}
        </g>
      </svg>
    </div>
  `;
}

function openFrontSalesFlowDiagram() {
  const flows = state.contractRelationships.macroFlows || [];
  const frontItems = state.contractRelationships.frontContractItems || [];
  const frontContractGroups = groupContractsWithItems('front_sales', frontItems);
  const amountByFlow = confirmedFrontAmountByMacroFlow(frontContractGroups);
  const totalAmount = totalAmountForConfirmedContractGroups(frontContractGroups);
  const existing = $('#frontSalesFlowDiagramModal');
  if (existing) existing.remove();
  const modal = document.createElement('div');
  modal.id = 'frontSalesFlowDiagramModal';
  modal.className = 'contract-flow-modal open';
  modal.innerHTML = `
    <div class="contract-flow-modal-backdrop" data-contract-flow-diagram-close="1"></div>
    <section class="contract-flow-modal-panel" role="dialog" aria-modal="true" aria-label="前向合同流链路框图">
      <header class="contract-flow-modal-head">
        <div class="contract-flow-title-copy">
          <span class="eyebrow">Front sales contract flow</span>
          <h3>前向合同流链路框图</h3>
          <p>同一合同主体只出现一个图框；金额口径为表格中已人工确认的前向销售合同金额。</p>
        </div>
        <aside class="front-flow-total-card compact">
          <span>天安智联</span>
          <em>已确认前向销售合同金额总额</em>
          <strong>${escapeHtml(formatAccountingYuan(totalAmount))}</strong>
        </aside>
        <button type="button" class="ghost-btn" data-contract-flow-diagram-close="1">关闭</button>
      </header>
      <div class="front-flow-stage">
        <div class="front-flow-stage-label left">业主入口 → 前向合同主体树状展开 → 合同摘要 → 天安智联</div>
      </div>
      <div class="front-flow-canvas tree">
        ${flows.length ? renderFrontSalesFlowTreeDiagram(flows, amountByFlow) : '<div class="contract-empty-note">暂无前向合同流链路数据，请先刷新候选数据。</div>'}
      </div>
    </section>
  `;
  document.body.appendChild(modal);
}

function closeFrontSalesFlowDiagram() {
  $('#frontSalesFlowDiagramModal')?.remove();
}

function itemFieldValue(item, keys, fallback = '-') {
  for (const key of keys) {
    const value = item?.[key];
    if (value !== undefined && value !== null && String(value).trim() !== '') return value;
  }
  return fallback;
}

function openBackContractLinePreview(contractId) {
  const items = (state.contractRelationships.backContractItems || []).filter((item) => item.contractId === contractId);
  const contract = (state.contractRelationships.contracts || []).find((item) => item.id === contractId) || items[0] || {};
  const existing = $('#backContractLinePreviewModal');
  if (existing) existing.remove();
  const modal = document.createElement('div');
  modal.id = 'backContractLinePreviewModal';
  modal.className = 'contract-flow-modal open';
  modal.innerHTML = `
    <div class="contract-flow-modal-backdrop" data-contract-line-preview-close="1"></div>
    <section class="contract-flow-modal-panel contract-line-preview-panel" role="dialog" aria-modal="true" aria-label="后向采购合同清单预览">
      <header class="contract-flow-modal-head">
        <div class="contract-flow-title-copy">
          <span class="eyebrow">Back procurement line preview</span>
          <h3>后向采购合同清单预览</h3>
          <p>${escapeHtml(contract.fileName || contract.contractName || '未命名合同')}；当前仅展示前 ${Math.min(items.length, 80)} / ${items.length} 行候选清单，正式关联仍需人工确认。</p>
        </div>
        <button type="button" class="ghost-btn" data-contract-line-preview-close="1">关闭</button>
      </header>
      <div class="table-wrap contract-line-preview-table-wrap">
        <table class="contract-manual-table contract-line-preview-table">
          <thead><tr><th>序号</th><th>分项 / 设备服务</th><th>规格型号</th><th>数量</th><th>含税金额（元）</th><th>税率</th><th>来源</th></tr></thead>
          <tbody>
            ${items.length ? items.slice(0, 80).map((item, index) => `
              <tr>
                <td>${index + 1}</td>
                <td><strong>${escapeHtml(itemDisplayName(item))}</strong><br /><small>${escapeHtml(itemFieldValue(item, ['detailName', 'itemNo'], ''))}</small></td>
                <td>${escapeHtml(itemFieldValue(item, ['specModel', 'model', 'brand'], '-'))}</td>
                <td>${escapeHtml(itemFieldValue(item, ['quantity'], '待确认'))} ${escapeHtml(item.unit || '')}</td>
                <td class="money-cell">${escapeHtml(formatAccountingYuan(item.amountTaxIncluded || 0))}</td>
                <td>${escapeHtml(allowedTaxRateDisplay([item.taxRate]))}</td>
                <td>表 ${escapeHtml(item.tableIndex ?? '-')} / 行 ${escapeHtml(item.rowNumber ?? '-')}</td>
              </tr>`).join('') : '<tr><td colspan="7">该合同尚未解析出清单行，后续需要人工补清单。</td></tr>'}
          </tbody>
        </table>
      </div>
    </section>
  `;
  document.body.appendChild(modal);
}

function closeBackContractLinePreview() {
  $('#backContractLinePreviewModal')?.remove();
}

function contractMacroFlowOptions(selectedId = '') {
  const flows = state.contractRelationships.macroFlows || [];
  const options = ['<option value="">请选择合同流链路</option>'];
  flows.forEach((flow) => {
    const label = flow.packageName || flow.displayText || flow.id;
    options.push(`<option value="${escapeHtml(flow.id)}" ${flow.id === selectedId ? 'selected' : ''}>${escapeHtml(label)}</option>`);
  });
  return options.join('');
}

function contractMacroFlowHelper(flowId) {
  const flow = (state.contractRelationships.macroFlows || []).find((item) => item.id === flowId);
  if (!flow) return '来源：Excel《合同流》下部描述；请选择后人工确认。';
  const parts = [flow.displayText, flow.notes].filter(Boolean);
  return parts.join('；') || '来源：Excel《合同流》下部描述。';
}

function activeContractWorkspace() {
  return $('[data-contract-workspace-panel].active')?.dataset.contractWorkspacePanel || 'front';
}

function normalizeContractImportKind(kind) {
  const text = String(kind || 'front');
  if (text === 'active-file' || text === 'active-dir') {
    const workspace = activeContractWorkspace();
    const isBack = workspace === 'back';
    const isBatch = text === 'active-dir';
    return `${isBack ? 'back' : 'front'}${isBatch ? '-batch' : ''}`;
  }
  return text;
}

function contractImportInput(kind) {
  kind = normalizeContractImportKind(kind);
  const isBack = String(kind || '').startsWith('back');
  const isBatch = String(kind || '').includes('batch');
  if (isBack) return isBatch ? $('#contractBackDirInput') : $('#contractBackFileInput');
  return isBatch ? $('#contractFrontDirInput') : $('#contractFrontFileInput');
}

function handleContractImport(kind) {
  kind = normalizeContractImportKind(kind);
  const isBack = String(kind || '').startsWith('back');
  setContractWorkspace(isBack ? 'back' : 'front');
  const input = contractImportInput(kind);
  if (!input) {
    alert('合同导入控件缺失，请检查页面初始化。');
    return;
  }
  input.value = '';
  input.click();
}

async function uploadContractImportFiles(kind, files) {
  const fileList = Array.from(files || []);
  if (!fileList.length) return;
  const isBack = String(kind || '').startsWith('back');
  const direction = isBack ? 'back' : 'front';
  const directionText = isBack ? '后向采购合同' : '前向销售合同';
  const refreshButton = $('#refreshContractRelationships');
  const originalText = refreshButton?.textContent || '刷新候选数据';
  if (refreshButton) {
    refreshButton.disabled = true;
    refreshButton.textContent = `正在导入${directionText}...`;
  }
  try {
    for (let index = 0; index < fileList.length; index += 1) {
      const file = fileList[index];
      const relativeName = file.webkitRelativePath || file.name;
      if (refreshButton) refreshButton.textContent = `正在导入 ${index + 1}/${fileList.length}`;
      const response = await fetch(`/api/contract-import?direction=${encodeURIComponent(direction)}&fileName=${encodeURIComponent(relativeName)}`, {
        method: 'POST',
        body: file,
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(payload.message || payload.error || `合同文件上传失败：HTTP ${response.status}`);
    }
    if (refreshButton) refreshButton.textContent = '正在扫描并重建...';
    await rebuildContractRelationships();
    const summary = state.contractRelationships.summary || {};
    alert(`${directionText}导入完成：${fileList.length} 个文件已进入系统导入区，并已刷新候选数据。\n\n当前合同数：${formatNumber(summary.contractDocumentCount || state.contractRelationships.contracts.length)}\n前向合同：${formatNumber(summary.frontContractCount || 0)}\n后向合同：${formatNumber(summary.backContractCount || 0)}`);
  } catch (error) {
    alert(`${directionText}导入失败：${error.message || error}`);
  } finally {
    if (refreshButton) {
      refreshButton.disabled = false;
      refreshButton.textContent = originalText;
    }
  }
}

function contractConfirmationStatusPill(confirmation) {
  if (confirmation?.status === 'confirmed') return '<span class="status-pill done">已确认</span>';
  if (confirmation?.status === 'draft') return '<span class="status-pill warn">已暂存</span>';
  return '<span class="status-pill warn">待人工确认</span>';
}

function contractTaxRateChecks(contractId, values) {
  const selected = new Set(values || []);
  const rates = ['13%', '9%', '6%'];
  const checked = rates
    .map(
      (rate) => `
        <label class="contract-tax-check">
          <input type="checkbox" data-contract-manual-field="taxRates" data-contract-id="${escapeHtml(contractId)}" value="${escapeHtml(rate)}" ${selected.has(rate) ? 'checked' : ''} />
          <span>${escapeHtml(rate)}</span>
        </label>`,
    )
    .join('');
  return `<div class="contract-tax-checks">${checked}</div>`;
}

function taxRatesForManualContract(group, confirmation, amount) {
  if (Array.isArray(confirmation.taxRates)) return confirmation.taxRates;
  if (Number(amount) === 0) return [];
  return candidateTaxRatesForGroup(group);
}

function candidateMacroFlowIdForContract(contractId) {
  const match = (state.contractRelationships.contractToMacroFlowMatches || []).find((item) => item.contractId === contractId);
  return match?.matches?.[0]?.macroFlowId || '';
}

function contractFlowLabel(flowId) {
  const flow = (state.contractRelationships.macroFlows || []).find((item) => item.id === flowId);
  if (!flowId) return '未分配合同流';
  return flow?.packageName || flow?.displayText || flowId;
}

function contractFlowSortIndex(flowId) {
  if (!flowId) return Number.MAX_SAFE_INTEGER;
  const index = (state.contractRelationships.macroFlows || []).findIndex((item) => item.id === flowId);
  return index >= 0 ? index : Number.MAX_SAFE_INTEGER - 1;
}

function supplierSortName(group) {
  return String(group?.supplierName || group?.partyBFullName || group?.counterparty || '').trim();
}

function sortManualContractsInFlow(a, b, direction, isConfirmedFlow) {
  if (direction === 'back' && isConfirmedFlow) {
    const supplierDiff = supplierSortName(a).localeCompare(supplierSortName(b), 'zh-CN');
    if (supplierDiff !== 0) return supplierDiff;
    const aIsZero = Math.abs(Number(confirmedAmountForContractGroup(a)) || 0) < 0.005;
    const bIsZero = Math.abs(Number(confirmedAmountForContractGroup(b)) || 0) < 0.005;
    if (aIsZero !== bIsZero) return aIsZero ? -1 : 1;
    return 0;
  }
  if (direction === 'back') return 0;
  return String(a.contractName || '').localeCompare(String(b.contractName || ''), 'zh-CN');
}

function buildManualContractFlowGroups(groups, direction = 'front') {
  const buckets = new Map();
  (groups || []).forEach((group) => {
    const confirmation = getContractManualConfirmation(group.contractId) || {};
    const isConfirmedFlow = confirmation.status === 'confirmed' && Boolean(confirmation.macroFlowId);
    const selectedFlowId = confirmation.macroFlowId || candidateMacroFlowIdForContract(group.contractId) || '';
    const mergeFlowId = isConfirmedFlow ? confirmation.macroFlowId : '';
    const key = isConfirmedFlow ? `confirmed:${mergeFlowId}` : `unconfirmed:${group.contractId}`;
    if (!buckets.has(key)) {
      buckets.set(key, {
        flowId: mergeFlowId,
        selectedFlowId,
        isConfirmedFlow,
        contracts: [],
        firstOrder: buckets.size,
      });
    }
    buckets.get(key).contracts.push(group);
  });
  return [...buckets.values()]
    .map((bucket) => ({
      ...bucket,
      contracts: bucket.contracts.sort((a, b) => sortManualContractsInFlow(a, b, direction, bucket.isConfirmedFlow)),
    }))
    .sort((a, b) => {
      if (a.isConfirmedFlow !== b.isConfirmedFlow) return a.isConfirmedFlow ? -1 : 1;
      if (a.isConfirmedFlow && b.isConfirmedFlow) {
        const indexDiff = contractFlowSortIndex(a.flowId) - contractFlowSortIndex(b.flowId);
        if (indexDiff !== 0) return indexDiff;
        return contractFlowLabel(a.flowId).localeCompare(contractFlowLabel(b.flowId), 'zh-CN');
      }
      if (direction === 'back') return (a.firstOrder || 0) - (b.firstOrder || 0);
      return String(a.contracts[0]?.contractName || '').localeCompare(String(b.contracts[0]?.contractName || ''), 'zh-CN');
    });
}

function renderManualContractRows(target, groups, direction) {
  const node = $(target);
  if (!node) return;
  const colspan = direction === 'front' ? 7 : 8;
  if (!groups.length) {
    node.innerHTML = `<tr><td colspan="${colspan}">暂无候选合同记录。请先导入合同文件并由人工确认。</td></tr>`;
    renderContractManualColumnHandles(node.closest("table"));
    return;
  }
  const flowGroups = buildManualContractFlowGroups(groups, direction).slice(0, 120);
  node.innerHTML = flowGroups
    .map((flowGroup, flowIndex) => {
      const visibleContracts = flowGroup.contracts;
      const rowSpan = Math.max(visibleContracts.length, 1);
      const groupContractIds = visibleContracts.map((item) => item.contractId);
      const selectedFlowId = flowGroup.selectedFlowId || flowGroup.flowId || '';
      const subtotalAmount = subtotalAmountForContractFlowGroup(flowGroup);
      const subtotalLabel = visibleContracts.length > 1 ? '链路小计' : '本行金额';
      const groupSelect = `<select class="contract-inline-select flow" data-contract-flow-group-field="macroFlowId" data-contract-ids="${escapeHtml(groupContractIds.join('|'))}">${contractMacroFlowOptions(selectedFlowId)}</select><div class="contract-flow-subtotal"><span>${subtotalLabel}</span><strong>${escapeHtml(formatAccountingYuan(subtotalAmount))}</strong></div>`;
      return visibleContracts
        .map((group, rowIndex) => {
          const contractId = group.contractId;
          const confirmation = getContractManualConfirmation(contractId) || {};
          const amount = manualAmountForContractGroup(group);
          const taxRates = taxRatesForManualContract(group, confirmation, amount);
          const status = contractConfirmationStatusPill(confirmation);
          const amountInput = `<input class="contract-inline-input money" data-contract-manual-field="amount" data-contract-id="${escapeHtml(contractId)}" value="${escapeHtml(formatAccountingYuan(amount))}" inputmode="decimal" />`;
          const actions = `<div class="contract-table-actions"><button class="link-btn" data-contract-manual-action="confirm" data-contract-id="${escapeHtml(contractId)}">确认</button><button class="link-btn muted" data-contract-manual-action="draft" data-contract-id="${escapeHtml(contractId)}">暂存</button><button class="link-btn danger" data-contract-manual-action="reset" data-contract-id="${escapeHtml(contractId)}">重置候选</button><button class="link-btn danger solid" data-contract-system-remove="${escapeHtml(contractId)}" data-contract-name="${escapeHtml(group.contractName)}" data-contract-source-path="${escapeHtml(group.sourcePath)}">删除</button></div>`;
          const rowClass = confirmation.status === 'confirmed' ? 'contract-confirmed-row' : '';
          const seqCell = rowIndex === 0 ? `<td rowspan="${rowSpan}" class="contract-group-seq">${flowIndex + 1}</td>` : '';
          const flowCell = rowIndex === 0 ? `<td rowspan="${rowSpan}" class="contract-flow-cell contract-flow-group-cell">${groupSelect}</td>` : '';
          if (direction === 'front') {
            return `
              <tr class="${rowClass}" data-manual-contract-id="${escapeHtml(contractId)}">
                ${seqCell}
                ${flowCell}
                <td>${contractFileButton(group.contractName, group.sourcePath)}</td>
                <td class="money-cell">${amountInput}</td>
                <td>${contractTaxRateChecks(contractId, taxRates)}</td>
                <td>${status}</td>
                <td>${actions}</td>
              </tr>`;
          }
          const backFlowCell = rowIndex === 0 ? `<td rowspan="${rowSpan}" class="contract-flow-cell contract-flow-group-cell">${groupSelect}</td>` : '';
          return `
            <tr class="${rowClass}" data-manual-contract-id="${escapeHtml(contractId)}">
              ${seqCell}
              ${backFlowCell}
              <td>${contractFileButton(group.contractName, group.sourcePath)}</td>
              <td><input class="contract-inline-input supplier" data-contract-manual-field="supplierName" data-contract-id="${escapeHtml(contractId)}" value="${escapeHtml(confirmation.supplierName || group.supplierName || '')}" placeholder="填写乙方全称" /></td>
              <td class="money-cell">${amountInput}</td>
              <td>${contractTaxRateChecks(contractId, taxRates)}</td>
              <td>${status}</td>
              <td>${actions}</td>
            </tr>`;
        })
        .join('');
    })
    .join('');
  renderContractManualColumnHandles(node.closest("table"));
}

function selectContractManualRow(row) {
  if (!row) return;
  const table = row.closest('.contract-manual-table');
  if (!table) return;
  table.querySelectorAll('tr.contract-selected-row').forEach((item) => item.classList.remove('contract-selected-row'));
  row.classList.add('contract-selected-row');
}

function parseContractAmountInput(value) {
  const normalized = String(value || '').replace(/,/g, '').replace(/，/g, '').trim();
  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function renderedContractFlowIdForContract(contractId) {
  const flowGroupSelects = Array.from(document.querySelectorAll('[data-contract-flow-group-field="macroFlowId"]'));
  const matched = flowGroupSelects.find((select) => String(select.dataset.contractIds || '').split('|').includes(contractId));
  if (matched) return matched.value || '';
  const confirmation = getContractManualConfirmation(contractId) || {};
  return confirmation.macroFlowId || candidateMacroFlowIdForContract(contractId) || '';
}

function readContractManualRow(contractId) {
  const row = document.querySelector(`[data-manual-contract-id="${CSS.escape(contractId)}"]`);
  if (!row) return {};
  const amount = parseContractAmountInput(row.querySelector('[data-contract-manual-field="amount"]')?.value || '');
  const supplierInput = row.querySelector('[data-contract-manual-field="supplierName"]');
  return {
    macroFlowId: renderedContractFlowIdForContract(contractId),
    amount,
    supplierName: supplierInput ? supplierInput.value.trim() : undefined,
    taxRates: amount === 0 ? [] : Array.from(row.querySelectorAll('[data-contract-manual-field="taxRates"]:checked')).map((item) => item.value),
  };
}

function handleContractFlowGroupChange(element) {
  const contractIds = String(element.dataset.contractIds || '').split('|').filter(Boolean);
  const macroFlowId = element.value || '';
  contractIds.forEach((contractId) => {
    upsertContractManualConfirmation(contractId, { macroFlowId, status: 'draft' });
  });
  renderContracts();
}

function handleContractManualFieldChange(element) {
  const contractId = element.dataset.contractId;
  if (!contractId) return;
  const patch = readContractManualRow(contractId);
  upsertContractManualConfirmation(contractId, { ...patch, status: 'draft' });
  renderContracts();
}

async function removeContractFromSystem(contractId, contractName, sourcePath) {
  if (!contractId && !sourcePath) return;
  const confirmed = confirm(`确认从系统中删除该合同？\n\n${contractName || contractId}\n\n该操作只会从系统列表和候选关系中移除，不会删除或移动源文件。`);
  if (!confirmed) return;
  try {
    const response = await fetch('/api/contract-system-remove', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ contractId, fileName: contractName, sourcePath }),
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.message || payload.error || `HTTP ${response.status}`);
    resetContractManualConfirmation(contractId);
    await loadContractRelationships();
    await loadDocumentAssets();
    alert('已从系统合同列表中删除；源文件未被修改。');
  } catch (error) {
    alert(`合同删除失败：${error.message || error}`);
  }
}

function handleContractManualAction(action, contractId) {
  if (!contractId) return;
  if (action === 'reset') {
    resetContractManualConfirmation(contractId);
  } else {
    const patch = readContractManualRow(contractId);
    upsertContractManualConfirmation(contractId, { ...patch, status: action === 'confirm' ? 'confirmed' : 'draft' });
  }
  renderContracts();
  renderMetrics();
  renderContractStrip();
}

function renderLineItems(target, items, emptyText) {
  const node = $(target);
  if (!node) return;
  node.innerHTML = (items || []).length
    ? items
        .slice(0, 80)
        .map(
          (item) => `
            <tr>
              <td>${contractFileButton(item.contractName, item.sourcePath)}</td>
              <td><strong>${escapeHtml(itemDisplayName(item))}</strong><br /><small>表 ${escapeHtml(item.tableIndex ?? '-')} / 行 ${escapeHtml(item.rowNumber ?? '-')}</small></td>
              <td>${escapeHtml(item.quantity ?? '待确认')} ${escapeHtml(item.unit || '')}</td>
              <td class="money-cell">${escapeHtml(formatAccountingYuan(item.amountTaxIncluded || 0))}</td>
              <td>${escapeHtml(allowedTaxRateDisplay([item.taxRate]))}</td>
            </tr>
          `,
        )
        .join('')
    : `<tr><td colspan="5">${escapeHtml(emptyText)}</td></tr>`;
}

function renderManualRelationRows(candidates, macroById) {
  const node = $('#lineRelationRows');
  if (!node) return;
  const rows = [...(candidates || [])]
    .sort((a, b) => (b.score || 0) - (a.score || 0))
    .slice(0, 100);
  node.innerHTML = rows.length
    ? rows
        .map((candidate) => {
          const decision = contractCandidateDecision(candidate.id).status || 'candidate';
          const statusText = decision === 'confirmed' ? '已人工确认' : decision === 'rejected' ? '已否决' : '候选 / 待人工确认';
          const statusClass = decision === 'confirmed' ? 'done' : decision === 'rejected' ? 'error' : 'warn';
          const profit = (Number(candidate.frontAmountTaxIncluded) || 0) - (Number(candidate.backAmountTaxIncluded) || 0);
          return `
            <tr>
              <td><strong>${escapeHtml(candidate.frontItemName || '-')}</strong> ↔ <strong>${escapeHtml(candidate.backItemName || '-')}</strong><br /><small>${escapeHtml(candidate.frontContractName || '-')} → ${escapeHtml(candidate.backContractName || '-')}</small></td>
              <td class="money-cell">${escapeHtml(formatAccountingYuan(candidate.frontAmountTaxIncluded || 0))}</td>
              <td class="money-cell">${escapeHtml(formatAccountingYuan(candidate.backAmountTaxIncluded || 0))}</td>
              <td>${escapeHtml(allowedTaxRateDisplay([candidate.frontTaxRate, candidate.backTaxRate]))}</td>
              <td class="money-cell">${escapeHtml(formatAccountingYuan(profit))}</td>
              <td><span class="status-pill ${statusClass}">${escapeHtml(statusText)}</span><div class="contract-review-actions"><button class="link-btn" data-contract-candidate-decision="confirmed" data-candidate-id="${escapeHtml(candidate.id)}">确认</button><button class="link-btn danger" data-contract-candidate-decision="rejected" data-candidate-id="${escapeHtml(candidate.id)}">否决</button></div></td>
            </tr>`;
        })
        .join('')
    : '<tr><td colspan="6">暂无前后向清单行候选。正式系统中应由人工选择前向行、后向行后建立关系。</td></tr>';
}

function renderAuditPaymentRows(candidates) {
  const node = $('#auditPaymentRows');
  if (!node) return;
  const confirmed = (candidates || []).filter((candidate) => contractCandidateDecision(candidate.id).status === 'confirmed').slice(0, 60);
  node.innerHTML = confirmed.length
    ? confirmed
        .map((candidate, index) => `
          <tr>
            <td><strong>设备级单元 ${index + 1}</strong><br /><small>${escapeHtml(candidate.frontItemName || candidate.backItemName || '-')}</small></td>
            <td><span class="status-pill warn">待绑定</span><br /><small>NodeID/点位需按审计资料确认</small></td>
            <td class="money-cell">0.00</td>
            <td class="money-cell">0.00</td>
            <td class="money-cell">0.00</td>
            <td class="money-cell">0.00</td>
            <td>待导入业主设备级审计结果</td>
          </tr>`)
        .join('')
    : '<tr><td colspan="7">尚无已人工确认的设备级清单关联。未确认关系不得进入审计、收款、付款台账。</td></tr>';
}

function confirmedContractAmountsByDirection() {
  const cr = state.contractRelationships;
  const frontGroups = groupContractsWithItems('front_sales', cr.frontContractItems || []);
  const backGroups = groupContractsWithItems('back_procurement', cr.backContractItems || []);
  const front = totalAmountForConfirmedContractGroups(frontGroups);
  const back = totalAmountForConfirmedContractGroups(backGroups);
  return { front, back, frontGroups, backGroups };
}

function renderProfitDashboard(candidates) {
  const confirmed = (candidates || []).filter((candidate) => contractCandidateDecision(candidate.id).status === 'confirmed');
  const { front, back, frontGroups, backGroups } = confirmedContractAmountsByDirection();
  renderContractKpis('#profitKpiRows', [
    ['已确认关联', confirmed.length, '只统计人工确认清单行关系', 'done'],
    ['前向确认金额', formatAccountingYuan(front), '只取前向合同表已人工确认金额', ''],
    ['后向确认金额', formatAccountingYuan(back), '只取后向合同表已人工确认金额', ''],
    ['利润留存', formatAccountingYuan(front - back), '只由合同表人工确认金额计算', front - back >= 0 ? 'done' : 'risk'],
  ]);
  const byFlow = $('#profitByFlowRows');
  if (byFlow) {
    const frontByFlow = confirmedFrontAmountByMacroFlow(frontGroups);
    const backByFlow = new Map();
    confirmedManualContractGroups(backGroups).forEach((group) => {
      const confirmation = getContractManualConfirmation(group.contractId) || {};
      const flowId = confirmation.macroFlowId || '';
      if (!flowId) return;
      backByFlow.set(flowId, (backByFlow.get(flowId) || 0) + confirmedAmountForContractGroup(group));
    });
    const rows = (state.contractRelationships.macroFlows || [])
      .map((flow) => {
        const retained = (frontByFlow.get(flow.id) || 0) - (backByFlow.get(flow.id) || 0);
        return { flow, retained };
      })
      .filter((item) => Math.abs(item.retained) > 0.005);
    byFlow.innerHTML = rows.length
      ? rows.map(({ flow, retained }) => `<div class="profit-row"><span>${escapeHtml(flow.packageName || flow.displayText || flow.id)}</span><strong>${escapeHtml(formatAccountingYuan(retained))}</strong></div>`).join('')
      : '<p class="contract-empty-note">暂无正式汇总。请先在前向/后向合同表中人工确认合同流链路和金额。</p>';
  }
  const risks = $('#profitRiskRows');
  if (risks) {
    const pending = (candidates || []).filter((candidate) => (contractCandidateDecision(candidate.id).status || 'candidate') === 'candidate').length;
    risks.innerHTML = `
      <div class="profit-row"><span>待人工确认候选</span><strong>${escapeHtml(formatNumber(pending))}</strong></div>
      <div class="profit-row"><span>未绑定 NodeID/点位</span><strong>${escapeHtml(formatNumber(confirmed.length))}</strong></div>
      <div class="profit-row"><span>未导入审计批次</span><strong>${escapeHtml(formatNumber(confirmed.length))}</strong></div>`;
  }
}

function downloadTextFile(fileName, content, mimeType = 'text/csv;charset=utf-8') {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function csvCell(value) {
  const text = String(value ?? '');
  return `"${text.replaceAll('"', '""')}"`;
}

function exportContractPendingRelations() {
  const rows = [['前向合同', '前向清单', '后向合同', '后向清单', '前向金额', '后向金额', '税率', '置信度', '状态']];
  (state.contractRelationships.deviceItemMatchCandidates || []).forEach((candidate) => {
    const decision = contractCandidateDecision(candidate.id).status || 'candidate';
    if (decision !== 'candidate') return;
    rows.push([
      candidate.frontContractName || '',
      [candidate.frontItemName, candidate.frontDetailName, candidate.frontSpecModel].filter(Boolean).join(' / '),
      candidate.backContractName || '',
      [candidate.backItemName, candidate.backDetailName, candidate.backSpecModel].filter(Boolean).join(' / '),
      formatAccountingYuan(candidate.frontAmountTaxIncluded || 0),
      formatAccountingYuan(candidate.backAmountTaxIncluded || 0),
      allowedTaxRateDisplay([candidate.frontTaxRate, candidate.backTaxRate]),
      candidate.confidence || '',
      '待人工确认',
    ]);
  });
  downloadTextFile(`合同待确认清单-${new Date().toISOString().slice(0, 10)}.csv`, rows.map((row) => row.map(csvCell).join(',')).join('\n'));
}

function exportContractProfitSummary() {
  const { front, back } = confirmedContractAmountsByDirection();
  const rows = [
    ['统计口径', '前向确认金额', '后向确认金额', '利润留存', '说明'],
    ['合同表人工确认金额', formatAccountingYuan(front), formatAccountingYuan(back), formatAccountingYuan(front - back), '不使用合同文件抽取金额、不使用候选清单金额'],
  ];
  downloadTextFile(`合同利润留存汇总-${new Date().toISOString().slice(0, 10)}.csv`, rows.map((row) => row.map(csvCell).join(',')).join('\n'));
}

function handleContractExport(kind) {
  if (kind === 'pending-relations') exportContractPendingRelations();
  else if (kind === 'profit-summary') exportContractProfitSummary();
}

function renderContractManualWorkspaces() {
  const cr = state.contractRelationships;
  const summary = cr.summary || {};
  const frontItems = cr.frontContractItems || [];
  const backItems = cr.backContractItems || [];
  const candidates = cr.deviceItemMatchCandidates || [];
  const frontContractGroups = groupContractsWithItems('front_sales', frontItems);
  const backContractGroups = groupContractsWithItems('back_procurement', backItems);
  const frontContractAmount = totalAmountForConfirmedContractGroups(frontContractGroups);
  const backContractAmount = totalAmountForConfirmedContractGroups(backContractGroups);

  const contractSummaryKpis = [
    ['前向销售合同数', frontContractGroups.length, '', 'done', 'front-sales-flow-diagram'],
    ['前向销售合同金额', formatAccountingYuan(frontContractAmount), '', ''],
    ['后向采购合同数', backContractGroups.length, '', 'done'],
    ['后向采购合同金额', formatAccountingYuan(backContractAmount), '', ''],
  ];

  renderContractKpis('#contractSummaryKpis', contractSummaryKpis);
  renderManualContractRows('#frontManualContractRows', frontContractGroups, 'front');

  renderManualContractRows('#backManualContractRows', backContractGroups, 'back');

  renderLineItems('#frontLineItemRows', frontItems, '暂无前向清单行候选。');
  renderLineItems('#backLineItemRows', backItems, '暂无后向清单行候选。');
  renderManualRelationRows(candidates);
  renderAuditPaymentRows(candidates);
  renderProfitDashboard(candidates);
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
  $("#drawerGcjSubtitle").textContent = `经度 ${site.lngGcj.toFixed(8)} / 纬度 ${site.latGcj.toFixed(8)}`;
  renderDrawerMapXmlPreview(site);
  ensureDrawerMapXmlPreview(site);
  const nearby = [site, ...nearestSites(site)];
  renderMap($("#drawerMap"), nearby, {
    label: `当前点位：${site.name}；周边点位：${nearby.length - 1} 个`,
    centerSite: site,
    currentNodeId: site.nodeId,
    zoom: 15,
    maxZoom: 16,
  });

  $("#detailBase").innerHTML = `
    <div class="kv-grid drawer-kv-compact">
      <div class="kv"><span>行政区域</span><strong>${districtBadge(site)}</strong></div>
      <div class="kv"><span>点位类型</span><strong>${site.type}</strong></div>
      <div class="kv"><span>信号机厂商</span><strong>${site.vendor}</strong></div>
      <div class="kv"><span>感知点位类型</span><strong>${perceptionType(site)}</strong></div>
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
  requestAnimationFrame(sizeDrawerMapXmlPreview);
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
  $("#sidebarWidthToggle")?.addEventListener("click", toggleSidebarExpanded);
  $$(".nav-item").forEach((item) => item.addEventListener("click", () => setPanel(item.dataset.panel)));
  $$("[data-panel-link]").forEach((item) => item.addEventListener("click", () => setPanel(item.dataset.panelLink)));
  globalThis.addEventListener?.("hashchange", () => setPanel(panelIdFromLocation(), { updateHash: false }));
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
          renderSiteMap();
        });
      }
    }),
  );
  $("#mapNodeIdWarningBlink")?.addEventListener("change", (event) => {
    state.mapNodeIdWarningBlink = event.target.checked;
    if (state.panel === "sites" && state.view === "map") renderSiteMap();
  });
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
    const manualContractRow = event.target.closest("tr[data-manual-contract-id]");
    if (manualContractRow) selectContractManualRow(manualContractRow);

    const contractSystemRemove = event.target.closest("[data-contract-system-remove]");
    if (contractSystemRemove) {
      removeContractFromSystem(
        contractSystemRemove.dataset.contractSystemRemove || '',
        contractSystemRemove.dataset.contractName || '',
        contractSystemRemove.dataset.contractSourcePath || '',
      );
      return;
    }
    const contractManualAction = event.target.closest("[data-contract-manual-action]");
    if (contractManualAction) {
      handleContractManualAction(contractManualAction.dataset.contractManualAction, contractManualAction.dataset.contractId);
      return;
    }
    const contractExport = event.target.closest("[data-contract-export]");
    if (contractExport) {
      handleContractExport(contractExport.dataset.contractExport);
      return;
    }
    const flowDiagramClose = event.target.closest("[data-contract-flow-diagram-close]");
    if (flowDiagramClose) {
      closeFrontSalesFlowDiagram();
      return;
    }
    const linePreviewClose = event.target.closest("[data-contract-line-preview-close]");
    if (linePreviewClose) {
      closeBackContractLinePreview();
      return;
    }
    const linePreview = event.target.closest("[data-contract-line-preview]");
    if (linePreview) {
      openBackContractLinePreview(linePreview.dataset.contractLinePreview || '');
      return;
    }
    const contractKpiAction = event.target.closest("[data-contract-kpi-action]");
    if (contractKpiAction) {
      if (contractKpiAction.dataset.contractKpiAction === "front-sales-flow-diagram") openFrontSalesFlowDiagram();
      return;
    }
    const contractZoom = event.target.closest("[data-contract-zoom]");
    if (contractZoom) {
      adjustContractPageScale(contractZoom.dataset.contractZoom);
      return;
    }
    const contractImport = event.target.closest("[data-contract-import]");
    if (contractImport) {
      handleContractImport(contractImport.dataset.contractImport);
      return;
    }
    const contractWorkspace = event.target.closest("[data-contract-workspace]");
    if (contractWorkspace) {
      setContractWorkspace(contractWorkspace.dataset.contractWorkspace || "front");
      return;
    }
    const contractEvidence = event.target.closest("[data-contract-candidate-evidence]");
    if (contractEvidence) {
      state.contractReview.selectedCandidateId = contractEvidence.dataset.contractCandidateEvidence || "";
      renderContracts();
      return;
    }
    const contractDecision = event.target.closest("[data-contract-candidate-decision]");
    if (contractDecision) {
      setContractCandidateDecision(contractDecision.dataset.candidateId, contractDecision.dataset.contractCandidateDecision);
      return;
    }
    const contractFileOpener = event.target.closest("[data-contract-open-file]");
    if (contractFileOpener) {
      event.preventDefault();
      event.stopPropagation();
      openContractFileWithDefaultApp(contractFileOpener.dataset.contractOpenFile || "");
      return;
    }
    const macroFlowSelect = event.target.closest("[data-select-macro-flow]");
    if (macroFlowSelect) {
      state.contractRelationships.selectedMacroFlowId = macroFlowSelect.dataset.selectMacroFlow || "";
      renderContracts();
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
    else if (handle?.dataset.resizeTable === "contract-manual") startContractManualColumnResize(event, handle);
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
  $("#fullscreenToggle").addEventListener("click", toggleFullscreen);
  document.addEventListener("fullscreenchange", renderFullscreenToggle);
  document.addEventListener("webkitfullscreenchange", renderFullscreenToggle);
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
  $('#contractFrontFileInput')?.addEventListener('change', (event) => uploadContractImportFiles('front', event.target.files));
  $('#contractFrontDirInput')?.addEventListener('change', (event) => uploadContractImportFiles('front-batch', event.target.files));
  $('#contractBackFileInput')?.addEventListener('change', (event) => uploadContractImportFiles('back', event.target.files));
  $('#contractBackDirInput')?.addEventListener('change', (event) => uploadContractImportFiles('back-batch', event.target.files));
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
    renderContracts();
  });
  document.body.addEventListener("change", (event) => {
    const flowGroupField = event.target.closest("[data-contract-flow-group-field]");
    if (flowGroupField) {
      handleContractFlowGroupChange(flowGroupField);
      return;
    }
    const manualField = event.target.closest("[data-contract-manual-field]");
    if (manualField) {
      handleContractManualFieldChange(manualField);
      return;
    }
    const filter = event.target.closest("[data-contract-filter]");
    if (!filter) return;
    state.contractReview.filters[filter.dataset.contractFilter] = filter.value;
    state.contractReview.selectedCandidateId = "";
    renderContracts();
  });
  document.body.addEventListener("keydown", (event) => {
    const contractKpiAction = event.target.closest?.("[data-contract-kpi-action]");
    if (contractKpiAction && (event.key === "Enter" || event.key === " ")) {
      event.preventDefault();
      if (contractKpiAction.dataset.contractKpiAction === "front-sales-flow-diagram") openFrontSalesFlowDiagram();
      return;
    }
    if (event.key === "Escape") {
      closeFrontSalesFlowDiagram();
      closeBackContractLinePreview();
    }
  });
  document.body.addEventListener("blur", (event) => {
    const manualField = event.target.closest?.("[data-contract-manual-field]");
    if (manualField?.dataset.contractManualField === "amount") {
      handleContractManualFieldChange(manualField);
    }
  }, true);
  $("#refreshDocumentAssets")?.addEventListener("click", loadDocumentAssets);
  $("#refreshContractRelationships")?.addEventListener("click", rebuildContractRelationships);
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
    if (state.view === "map") renderSiteMap();
  });
  window.addEventListener("resize", () => {
    applySiteViewLayoutSettled();
    if (state.panel === "contracts") applyContractPageScale();
    if ($("#detailDrawer")?.classList.contains("open")) sizeDrawerMapXmlPreview();
    if (state.activeSiteFilter) renderSiteFilterPopover();
    if (state.activeMapAssetFilter) renderMapAssetFilterPopover();
  });
}

async function init() {
  state.visualTheme = loadVisualTheme();
  loadContractReviewState();
  loadContractManualConfirmations();
  loadContractPageScale();
  applyVisualTheme(state.visualTheme);
  applySidebarExpanded(loadSidebarExpanded());
  state.panel = loadSavedPanelId();
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
  setPanel(state.panel, { updateHash: Boolean(globalThis.location?.hash) });
  bindEvents();
  renderFullscreenToggle();
  loadMapAssets();
  loadDocumentAssets();
  loadContractRelationships();
}

init();
