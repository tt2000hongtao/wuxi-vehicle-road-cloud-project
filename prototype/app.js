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
};

const data = window.PROTOTYPE_DATA;
const sites = data.sites;
const contracts = data.contracts;
const unmatchedItems = buildUnmatchedItems();
const documentBuckets = [
  ["勘察资料", "点位踏勘、杆件条件、机房资源", 128],
  ["施工照片", "安装过程、开箱验收、现场问题", 468],
  ["验收资料", "到货查验、入库、开箱、竣工验收", 216],
  ["合同附件", "前向合同、后向合同、分项包和设备明细", 74],
  ["运维工单", "异常闭环、变更记录、设备维护", 39],
  ["图纸报告", "设计图、接线表、测评报告", 93],
];
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
    purpose: "用于导入设备点位分布信息，并按 NodeID 关联到点位。",
    template: "/Users/tt2000/Documents/天安智联/AI/项目管理工具包/Project Data/设备安装位置表.xlsx",
    result: "生成设备安装台账；未匹配记录进入 2447 条未匹配池，支持治理到点位、仓库、机房或服务项。",
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

function formatNumber(value) {
  return new Intl.NumberFormat("zh-CN").format(value || 0);
}

function formatMoney(value) {
  if (!value) return "0";
  if (value >= 100000000) return `${(value / 100000000).toFixed(2)} 亿`;
  if (value >= 10000) return `${(value / 10000).toFixed(1)} 万`;
  return formatNumber(value);
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

function importPanelElement() {
  return state.panel === "imports" && $("#importCenterPanel") ? $("#importCenterPanel") : $("#importPanel");
}

function filteredSites() {
  const q = state.query.trim().toLowerCase();
  return sites.filter((site) => {
    const text = `${site.nodeId} ${site.name} ${site.vendor} ${site.district} ${site.type}`.toLowerCase();
    const districtOk = state.district === "全部区域" || site.district === state.district;
    const perceptionOk = state.perception === "全部感知类型" || perceptionType(site) === state.perception;
    const adaptiveOk =
      state.adaptive === "全部自适应" ||
      (state.adaptive === "自适应" ? site.adaptive : !site.adaptive);
    const variableOk =
      state.variableLane === "全部可变车道" ||
      (state.variableLane === "可变车道" ? isVariableLane(site) : !isVariableLane(site));
    return districtOk && perceptionOk && adaptiveOk && variableOk && (!q || text.includes(q));
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
  if (panelId === "sites" && state.view === "map") requestAnimationFrame(() => renderMap($("#siteMap"), filteredSites()));
  if (panelId === "coordinates") requestAnimationFrame(renderCoordinateIssues);
}

function renderMetrics() {
  $("#metricSiteTotal").textContent = formatNumber(data.stats.siteTotal);
  $("#metricDeviceRows").textContent = formatNumber(data.stats.deviceRows);
  $("#metricUnmatchedRows").textContent = formatNumber(data.stats.unmatchedRows);
  $("#metricContractRows").textContent = "前向 7 / 后向 10";
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
  $("#contractStrip").innerHTML = contracts
    .map(
      (contract) => `
        <div class="contract-row">
          <div>
            <b>${contract.package}</b>
            <span>${contract.path} / ${contract.risk}</span>
          </div>
          <span class="status-pill ${contract.status === "已签署" ? "done" : "warn"}">${contract.status}</span>
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
  return type === "sites" ? "点位管理表-无锡车路云.xlsx" : "设备安装位置表.xlsx";
}

function fileLooksLikeTemplate(type, fileName) {
  const normalized = fileName.replace(/\s+/g, "");
  if (type === "sites") return normalized.includes("点位管理表") && normalized.endsWith(".xlsx");
  return normalized.includes("设备安装位置") && normalized.endsWith(".xlsx");
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

async function createImportTaskFromWorkbook(type, file) {
  const workbook = await readWorkbookFile(file);
  if (type === "sites") return createSiteImportTask(file, workbook);
  return createDeviceImportTask(file, workbook);
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

function createImportTask(type, file) {
  const isSites = type === "sites";
  const sampleRows = isSites ? sites.slice(0, 8) : allDevices().slice(0, 8);
  const invalidRows = isSites ? 1 : 0;
  const matchedRows = allDevices().length;
  return {
    id: `${type}-${Date.now()}`,
    type,
    fileName: file?.name || expectedFileName(type),
    fileSize: file?.size || 0,
    templateMatched: file ? fileLooksLikeTemplate(type, file.name) : true,
    totalRows: isSites ? data.stats.siteTotal + invalidRows : data.stats.deviceRows + data.stats.unmatchedRows,
    validRows: isSites ? data.stats.siteTotal : data.stats.deviceRows,
    invalidRows,
    duplicateNodeIdCount: isSites ? 0 : null,
    coordinateErrorCount: isSites ? 1 : null,
    dictionaryErrorCount: isSites ? 3 : null,
    sheet1TotalRows: isSites ? null : data.stats.deviceRows,
    matchedRows: isSites ? null : matchedRows,
    unmatchedSheetRows: isSites ? null : data.stats.unmatchedRows,
    unpackingIssueRows: isSites ? null : 42,
    quantityErrorRows: isSites ? null : 18,
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
                  : `<div><strong>${row.nodeId || "未匹配"}</strong><span>${row.name || "-"} / ${row.supplier || "-"} / ${row.contract || "-"}</span></div>`,
              )
              .join("")}
          </div>
        </section>
      </div>
      <div class="import-footer">
        ${
          stage === "report"
            ? `<button class="primary-btn" data-import-finish="${task.type}">${isSites ? "查看点位管理" : "查看设备管理"}</button>
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
  return `
    <div class="import-report">
      <h3>导入报告</h3>
      <p>${isSites ? "点位库已更新，地图坐标以 GCJ-02 为准，CGCS2000 已进入审计字段。" : "设备安装台账已更新，未匹配记录已进入治理池。"}</p>
      <ul>
        <li>${isSites ? `成功导入 ${formatNumber(task.validRows)} 个点位` : `成功匹配 ${formatNumber(task.matchedRows)} 条设备记录`}</li>
        <li>${isSites ? `坐标异常 ${formatNumber(task.coordinateErrorCount)} 条，进入坐标治理` : `未匹配 ${formatNumber(task.unmatchedSheetRows)} 条，进入未匹配池`}</li>
        <li>${isSites ? `字典异常 ${formatNumber(task.dictionaryErrorCount)} 条，需人工复核` : `开箱验收问题 ${formatNumber(task.unpackingIssueRows)} 条，生成异常标识`}</li>
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

function confirmImport(type) {
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
  pushImportHistory(state.importTask);
  renderMetrics();
  renderStatusProgress();
  renderSites();
  renderDevices();
  renderWarehouse();
  renderImportCenter();
  renderCoordinateIssues();
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

function renderFilters() {
  const districts = ["全部区域", ...Array.from(new Set(sites.map((site) => site.district))).sort()];
  const perceptions = ["全部感知类型", ...Array.from(new Set(sites.map(perceptionType))).sort()];
  $("#districtFilter").innerHTML = districts.map((item) => `<option>${item}</option>`).join("");
  $("#perceptionFilter").innerHTML = perceptions.map((item) => `<option>${item}</option>`).join("");
  $("#adaptiveFilter").innerHTML = ["全部自适应", "自适应", "非自适应"].map((item) => `<option>${item}</option>`).join("");
  $("#variableLaneFilter").innerHTML = ["全部可变车道", "可变车道", "非可变车道"].map((item) => `<option>${item}</option>`).join("");
}

function renderSites() {
  const rows = filteredSites();
  $("#siteRows").innerHTML = rows
    .map(
      (site) => `
        <tr data-open-site="${site.nodeId}" tabindex="0" aria-label="打开 ${site.name} 编辑页">
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

  $("#siteCards").innerHTML = rows
    .map(
      (site) => `
        <article class="site-card">
          <div class="card-top">${districtBadge(site)} <span class="status-pill ${statusClass(site)}">${site.status}</span></div>
          <h3>${site.name}</h3>
          <p>${site.type} / NodeID ${site.nodeId}<br />${site.vendor || "-"} / ${perceptionType(site)} / ${site.adaptive ? "自适应" : "非自适应"} / ${isVariableLane(site) ? "可变车道" : "非可变车道"}</p>
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
  $("#flowBoard").innerHTML = contracts
    .map((contract) => {
      const nodes = contract.path.split("->").map((node) => node.trim());
      return `
        <div class="flow-row">
          <div class="flow-path">
            ${nodes
              .map((node, index) => {
                const arrow = index === nodes.length - 1 ? "" : '<span class="flow-arrow">→</span>';
                return `<span class="flow-node">${node}</span>${arrow}`;
              })
              .join("")}
          </div>
          <b>${contract.package}</b>
          <p>${formatMoney(contract.amount)} / ${contract.status} / ${contract.risk}</p>
        </div>
      `;
    })
    .join("");
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
  $("#documentGrid").innerHTML = documentBuckets
    .map(
      ([title, note, count]) => `
        <article class="document-card">
          <span>${title}</span>
          <strong>${formatNumber(count)}</strong>
          <p>${note}</p>
          <button class="link-btn">查看资料</button>
        </article>
      `,
    )
    .join("");
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
  });
  $("#districtFilter").addEventListener("change", (event) => {
    state.district = event.target.value;
    renderSites();
  });
  $("#perceptionFilter").addEventListener("change", (event) => {
    state.perception = event.target.value;
    renderSites();
  });
  $("#adaptiveFilter").addEventListener("change", (event) => {
    state.adaptive = event.target.value;
    renderSites();
  });
  $("#variableLaneFilter").addEventListener("change", (event) => {
    state.variableLane = event.target.value;
    renderSites();
  });
  $$(".segmented button").forEach((button) =>
    button.addEventListener("click", () => {
      state.view = button.dataset.view;
      $$(".segmented button").forEach((item) => item.classList.toggle("active", item === button));
      $$(".site-view").forEach((view) => view.classList.toggle("active", view.id === `${state.view}View`));
      if (state.view === "map") {
        requestAnimationFrame(() => renderMap($("#siteMap"), filteredSites()));
      }
    }),
  );
  document.body.addEventListener("click", (event) => {
    const chooseImport = event.target.closest("[data-choose-import]");
    if (chooseImport) {
      const type = chooseImport.dataset.chooseImport;
      const input = type === "sites" ? $("#siteImportInput") : $("#deviceImportInput");
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
    const finishButton = event.target.closest("[data-import-finish]");
    if (finishButton) {
      setPanel(finishButton.dataset.importFinish === "sites" ? "sites" : "devices");
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
    const opener = event.target.closest("[data-open-site]");
    if (opener) openSite(opener.dataset.openSite);
  });
  $$(".detail-tabs button").forEach((button) =>
    button.addEventListener("click", () => {
      $$(".detail-tabs button").forEach((item) => item.classList.toggle("active", item === button));
      $$(".detail-pane").forEach((pane) => pane.classList.toggle("active", pane.id === `detail${button.dataset.detailTab[0].toUpperCase()}${button.dataset.detailTab.slice(1)}`));
    }),
  );
  $("#closeDrawer").addEventListener("click", closeDrawer);
  $("#scrim").addEventListener("click", closeDrawer);
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
  $("#showUnmatched").addEventListener("click", () => {
    $("#unmatchedPanel").classList.toggle("open");
    $("#unmatchedTableWrap").classList.toggle("open");
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
}

function init() {
  renderMetrics();
  renderStatusProgress();
  renderContractStrip();
  renderMap($("#overviewMap"), sites.slice(0, 24));
  renderFilters();
  renderSites();
  renderDevices();
  renderContracts();
  renderWarehouse();
  renderImportCenter();
  renderCoordinateIssues();
  renderDocuments();
  bindEvents();
}

init();
