const http = require("http");
const fs = require("fs/promises");
const path = require("path");
const { execFile } = require("child_process");
const crypto = require("crypto");
const { fileURLToPath } = require("url");

const PORT = Number(process.env.PORT || 4173);
const ROOT_DIR = __dirname;
const STORAGE_DIR = path.join(ROOT_DIR, "storage");
const STATE_FILE = path.join(STORAGE_DIR, "roadside-status-state.json");
const MAP_ASSET_DIR = "/Users/tt2000/Documents/天安智联/AI/项目管理工具包/无锡车路云MAP、RSI及信号机配置文件Excel";
const CONTRACT_FILE_ROOTS = [
  "/Users/tt2000/Documents/天安智联/0000无锡市车路云一体化项目/0.合同/盖章版合同 WORD 版",
  "/Users/tt2000/Documents/天安智联/0000无锡市车路云一体化项目/0.合同/盖章版合同",
  path.join(ROOT_DIR, "..", "document_assets", "import_batches", "contract_manual_imports"),
];
const MAP_ASSET_CACHE_TTL_MS = 30 * 1000;
const DOCUMENT_ASSET_REPORT = path.join(ROOT_DIR, "..", "document_assets", "import_batches", "latest-document-scan.json");
const DOCUMENT_ASSET_SAMPLE_REPORT = path.join(ROOT_DIR, "..", "document_assets", "import_batches", "sample-document-assets.json");
const CONTRACT_REBUILD_SCRIPT = path.join(ROOT_DIR, "..", "tools", "rebuild_contract_relationships.py");
const CONTRACT_MANUAL_IMPORT_DIR = path.join(ROOT_DIR, "..", "document_assets", "import_batches", "contract_manual_imports");
const CONTRACT_EXCLUSIONS_FILE = path.join(STORAGE_DIR, "contract-exclusions.json");
const CONTRACT_RELATIONSHIPS_FILE = path.join(ROOT_DIR, "data", "contract-relationships.json");
const DOCUMENT_ASSETS_FILE = path.join(ROOT_DIR, "data", "document-assets.json");
const PYTHON_BIN = process.env.PYTHON_BIN || "/Library/Frameworks/Python.framework/Versions/3.11/bin/python3";
let mapAssetCache = null;

const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".doc": "application/msword",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".pdf": "application/pdf",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

const CRC32_TABLE = Array.from({ length: 256 }, (_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});

function sendJson(response, status, payload) {
  response.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  response.end(JSON.stringify(payload));
}

function sendNoContent(response) {
  response.writeHead(204, { "Cache-Control": "no-store" });
  response.end();
}

async function readRequestJson(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  const text = Buffer.concat(chunks).toString("utf8");
  return text ? JSON.parse(text) : null;
}


async function readJsonFile(filePath, fallback) {
  try {
    const body = await fs.readFile(filePath, "utf8");
    return JSON.parse(body);
  } catch (error) {
    if (error.code === "ENOENT") return fallback;
    throw error;
  }
}

async function writeJsonFile(filePath, payload) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
}

function normalizeExcludeKey(value) {
  return String(value || "").trim();
}

async function readContractExclusions() {
  const payload = await readJsonFile(CONTRACT_EXCLUSIONS_FILE, null);
  if (!payload) {
    return { schemaVersion: 1, updatedAt: "", exclusions: [] };
  }
  return {
    schemaVersion: 1,
    updatedAt: payload.updatedAt || "",
    exclusions: Array.isArray(payload.exclusions) ? payload.exclusions : [],
  };
}

async function restoreContractExclusionForImport(record) {
  const current = await readContractExclusions();
  const before = current.exclusions.length;
  const restored = [];
  current.exclusions = current.exclusions.filter((item) => {
    const matched = contractExcludedBy([item], record);
    if (matched) restored.push(item);
    return !matched;
  });
  if (current.exclusions.length === before) return { restoredCount: 0, restored };
  current.updatedAt = new Date().toISOString();
  await writeJsonFile(CONTRACT_EXCLUSIONS_FILE, current);
  return { restoredCount: before - current.exclusions.length, restored };
}

function contractExcludedBy(exclusions, record) {
  const ids = new Set(exclusions.map((item) => normalizeExcludeKey(item.contractId)).filter(Boolean));
  const paths = new Set(exclusions.map((item) => normalizeExcludeKey(item.sourcePath)).filter(Boolean));
  const hashes = new Set(exclusions.map((item) => normalizeExcludeKey(item.sha256)).filter(Boolean));
  return Boolean(
    ids.has(normalizeExcludeKey(record.id || record.contractId)) ||
      paths.has(normalizeExcludeKey(record.sourcePath)) ||
      hashes.has(normalizeExcludeKey(record.sha256)),
  );
}

function relationshipReferencesExcludedContract(item, exclusions) {
  return contractExcludedBy(exclusions, { contractId: item.contractId }) ||
    contractExcludedBy(exclusions, { contractId: item.frontContractId }) ||
    contractExcludedBy(exclusions, { contractId: item.backContractId });
}

function itemReferencesExcludedContract(item, exclusions) {
  return contractExcludedBy(exclusions, { contractId: item.contractId }) ||
    contractExcludedBy(exclusions, { contractId: item.frontContractId }) ||
    contractExcludedBy(exclusions, { contractId: item.backContractId });
}

async function applyContractExclusionsToGeneratedData(exclusions) {
  const relationships = await readJsonFile(CONTRACT_RELATIONSHIPS_FILE, null);
  if (relationships) {
    const contracts = relationships.contracts || [];
    relationships.contracts = contracts.filter((item) => !contractExcludedBy(exclusions, item));
    relationships.frontContractItems = (relationships.frontContractItems || []).filter((item) => !itemReferencesExcludedContract(item, exclusions));
    relationships.backContractItems = (relationships.backContractItems || []).filter((item) => !itemReferencesExcludedContract(item, exclusions));
    relationships.contractToMacroFlowMatches = (relationships.contractToMacroFlowMatches || []).filter((item) => !relationshipReferencesExcludedContract(item, exclusions));
    relationships.frontBackRelationshipCandidates = (relationships.frontBackRelationshipCandidates || []).filter((item) => !relationshipReferencesExcludedContract(item, exclusions));
    relationships.deviceItemMatchCandidates = (relationships.deviceItemMatchCandidates || []).filter((item) => !relationshipReferencesExcludedContract(item, exclusions));
    relationships.summary = {
      ...(relationships.summary || {}),
      contractDocumentCount: relationships.contracts.length,
      frontContractCount: relationships.contracts.filter((item) => item.direction === "front_sales").length,
      backContractCount: relationships.contracts.filter((item) => item.direction === "back_procurement").length,
      frontContractItemCount: relationships.frontContractItems.length,
      backContractItemCount: relationships.backContractItems.length,
      frontBackCandidateCount: relationships.frontBackRelationshipCandidates.length,
      deviceItemMatchCandidateCount: relationships.deviceItemMatchCandidates.length,
      excludedContractCount: exclusions.length,
    };
    await writeJsonFile(CONTRACT_RELATIONSHIPS_FILE, relationships);
  }

  for (const documentAssetFile of [DOCUMENT_ASSETS_FILE, DOCUMENT_ASSET_REPORT]) {
    const documentAssets = await readJsonFile(documentAssetFile, null);
    if (!documentAssets) continue;
    const before = documentAssets.records || [];
    documentAssets.records = before.filter((item) => !contractExcludedBy(exclusions, item));
    documentAssets.summary = {
      ...(documentAssets.summary || {}),
      fileCount: documentAssets.records.length,
      excludedFileCount: exclusions.length,
    };
    await writeJsonFile(documentAssetFile, documentAssets);
  }
}

async function sha256ForContractSourcePath(sourcePath) {
  if (!sourcePath) return "";
  for (const filePath of [DOCUMENT_ASSETS_FILE, DOCUMENT_ASSET_REPORT]) {
    const payload = await readJsonFile(filePath, null);
    const record = (payload?.records || []).find((item) => item.sourcePath === sourcePath);
    if (record?.sha256) return record.sha256;
  }
  return "";
}

async function handleContractSystemRemove(request, response) {
  if (request.method !== "POST") {
    response.writeHead(405, { Allow: "POST" });
    response.end();
    return;
  }
  try {
    const payload = await readRequestJson(request);
    const contractId = normalizeExcludeKey(payload?.contractId);
    const sourcePath = normalizeExcludeKey(payload?.sourcePath);
    const fileName = normalizeExcludeKey(payload?.fileName);
    if (!contractId && !sourcePath) return sendJson(response, 400, { error: "missing_contract_identifier" });
    const exclusions = await readContractExclusions();
    const sha256 = normalizeExcludeKey(payload?.sha256) || await sha256ForContractSourcePath(sourcePath);
    const exists = exclusions.exclusions.some((item) =>
      (contractId && item.contractId === contractId) ||
      (sourcePath && item.sourcePath === sourcePath) ||
      (sha256 && item.sha256 === sha256),
    );
    if (!exists) {
      exclusions.exclusions.push({
        contractId,
        sourcePath,
        fileName,
        sha256,
        reason: normalizeExcludeKey(payload?.reason) || "manual_remove_from_system",
        removedAt: new Date().toISOString(),
      });
    }
    exclusions.updatedAt = new Date().toISOString();
    await writeJsonFile(CONTRACT_EXCLUSIONS_FILE, exclusions);
    await applyContractExclusionsToGeneratedData(exclusions.exclusions);
    sendJson(response, 200, { ok: true, excludedContractCount: exclusions.exclusions.length, sourceFileTouched: false });
  } catch (error) {
    console.error(error);
    sendJson(response, 500, { error: "contract_system_remove_failed", message: error.message });
  }
}

function validateStatePayload(payload) {
  return Boolean(
    payload &&
      typeof payload.currentDate === "string" &&
      Array.isArray(payload.currentRows) &&
      Array.isArray(payload.archives),
  );
}

function extensionOf(fileName) {
  return path.extname(fileName).slice(1).toLowerCase();
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) crc = CRC32_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function dateToDosTime(date = new Date()) {
  const year = Math.max(1980, date.getFullYear());
  const dosTime = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const dosDate = ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { dosDate, dosTime };
}

function sanitizeArchiveName(value) {
  return String(value || "map-assets")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 120);
}

function mapAssetFilePath(relativePath) {
  const filePath = path.normalize(path.join(MAP_ASSET_DIR, relativePath));
  const relative = path.relative(MAP_ASSET_DIR, filePath);
  if (relative.startsWith("..") || path.isAbsolute(relative)) return "";
  return filePath;
}

function contractFilePath(rawPath) {
  const filePath = path.normalize(rawPath || "");
  if (!path.isAbsolute(filePath)) return "";
  const allowed = CONTRACT_FILE_ROOTS.some((root) => {
    const relative = path.relative(root, filePath);
    return relative && !relative.startsWith("..") && !path.isAbsolute(relative);
  });
  return allowed ? filePath : "";
}

function contentDispositionInline(filePath) {
  const fileName = path.basename(filePath);
  const asciiFallback = fileName.replace(/[^\x20-\x7e]/g, "_").replace(/["\\]/g, "_") || "contract-file";
  return `inline; filename="${asciiFallback}"; filename*=UTF-8''${encodeURIComponent(fileName)}`;
}

function openWithDefaultApp(filePath) {
  return new Promise((resolve, reject) => {
    execFile("open", [filePath], (error) => {
      if (error) reject(error);
      else resolve();
    });
  });
}

function buildZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  entries.forEach((entry) => {
    const nameBuffer = Buffer.from(entry.name, "utf8");
    const checksum = crc32(entry.body);
    const { dosDate, dosTime } = dateToDosTime(entry.modifiedAt ? new Date(entry.modifiedAt) : new Date());
    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0x0800, 6);
    localHeader.writeUInt16LE(0, 8);
    localHeader.writeUInt16LE(dosTime, 10);
    localHeader.writeUInt16LE(dosDate, 12);
    localHeader.writeUInt32LE(checksum, 14);
    localHeader.writeUInt32LE(entry.body.length, 18);
    localHeader.writeUInt32LE(entry.body.length, 22);
    localHeader.writeUInt16LE(nameBuffer.length, 26);
    localHeader.writeUInt16LE(0, 28);
    localParts.push(localHeader, nameBuffer, entry.body);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0x0800, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(dosTime, 12);
    centralHeader.writeUInt16LE(dosDate, 14);
    centralHeader.writeUInt32LE(checksum, 16);
    centralHeader.writeUInt32LE(entry.body.length, 20);
    centralHeader.writeUInt32LE(entry.body.length, 24);
    centralHeader.writeUInt16LE(nameBuffer.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);
    centralParts.push(centralHeader, nameBuffer);
    offset += localHeader.length + nameBuffer.length + entry.body.length;
  });

  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);
  return Buffer.concat([...localParts, ...centralParts, end]);
}

function parseNodeFileName(baseName) {
  const match = baseName.match(/^(?:map|rsi)_?(\d+)[_-](.+)$/) || baseName.match(/^(\d+)[_-](.+)$/);
  if (!match) return { nodeId: "", intersectionName: baseName, tail: "" };
  const nodeId = match[1];
  const rest = match[2];
  const versionMatch = rest.match(/(?:[-_])(20\d{6}T\d{4})(?:$|[-_])/) || rest.match(/(?:[-_])(20\d{10})(?:$|[-_])/) || rest.match(/(?:[-_])(20\d{6})(?:$|[-_])/);
  const versionIndex = versionMatch ? rest.indexOf(versionMatch[1]) : -1;
  const beforeVersion = versionIndex >= 0 ? rest.slice(0, Math.max(0, versionIndex - 1)) : rest;
  const segments = beforeVersion.split("-").filter(Boolean);
  const nameSegments = segments.filter((segment) => {
    if (/^\d{8,}$/.test(segment)) return false;
    if (/^\d+(?:、\d+)+$/.test(segment)) return false;
    if (/^\d+-\d+-/.test(segment)) return false;
    return true;
  });
  return {
    nodeId,
    intersectionName: nameSegments.join("-") || beforeVersion || rest,
    tail: versionIndex >= 0 ? rest.slice(versionIndex) : rest,
  };
}

function parseAssetFile(relativePath) {
  const parts = relativePath.split(path.sep);
  const district = parts[0] || "未分区";
  const folder = parts[1] || "";
  const fileName = parts.at(-1) || "";
  const ext = extensionOf(fileName);
  const baseName = fileName.replace(/\.[^.]+$/, "");
  const categoryMap = {
    map_xml: ext === "svg" ? "map_svg" : "map_xml",
    map_json: ext === "svg" ? "map_svg" : "map_json",
    rsi_xml: "rsi_xml",
    rsi_json: "rsi_json",
    信号机: "signal_excel",
  };
  const category = categoryMap[folder] || ext;
  const parsedName = parseNodeFileName(baseName);
  const nodeId = parsedName.nodeId;
  const intersectionName = parsedName.intersectionName;
  const tail = parsedName.tail;
  const versionMatch = tail.match(/(20\d{6}T\d{4})/) || tail.match(/(20\d{10})/) || tail.match(/(20\d{6})(?!\d)/);
  return {
    relativePath,
    district,
    folder,
    fileName,
    ext,
    category,
    nodeId,
    intersectionName,
    version: versionMatch?.[1] || "",
  };
}

function hasMapNodeId(value) {
  const text = String(value ?? "").trim();
  return Boolean(text && text !== "0");
}

function hasFiveDigitMapNodeId(value) {
  const text = String(value ?? "").trim();
  if (!text || text === "0") return false;
  return text.replace(/\D/g, "").length >= 5;
}

function isShortPresentMapNodeId(value) {
  const text = String(value ?? "").trim();
  if (!text) return false;
  const digitLength = text.replace(/\D/g, "").length;
  return digitLength > 0 && digitLength < 5;
}

function mapQualityEmpty() {
  return {
    nodeIdMissing: false,
    mapXmlNodeIdShort: false,
    stopLineMissing: false,
  };
}

function mergeMapQuality(target, source) {
  if (!source) return target;
  target.nodeIdMissing = target.nodeIdMissing || Boolean(source.nodeIdMissing);
  target.mapXmlNodeIdShort = target.mapXmlNodeIdShort || Boolean(source.mapXmlNodeIdShort);
  target.stopLineMissing = target.stopLineMissing || Boolean(source.stopLineMissing);
  return target;
}

function jsonLaneType(lane) {
  return lane?.laneAttributes?.laneType || {};
}

function isNormalVehicleJsonLane(lane) {
  const laneId = Number(lane?.laneId ?? lane?.laneID ?? 0);
  const laneType = jsonLaneType(lane);
  return Object.prototype.hasOwnProperty.call(laneType, "vehicle") && !Object.prototype.hasOwnProperty.call(laneType, "crosswalk") && !lane?.laneTypeAttrVehExt && (!laneId || laneId < 240);
}

function hasJsonStopLine(lane) {
  return (lane?.stopLines || []).some((stop) => {
    const center = stop?.centerPoint || {};
    return center.longitude != null || center.latitude != null || center.long != null || center.lat != null;
  });
}

function hasJsonDownstreamNode(lane) {
  return (lane?.connectsTo || []).some((connection) => hasMapNodeId(connection?.remoteIntersection?.id));
}

function inspectMapJsonContent(content) {
  const quality = mapQualityEmpty();
  try {
    const data = JSON.parse(content);
    const nodes = data?.MessageFrame?.map?.nodes || data?.MessageFrame?.mapFrame?.nodes || data?.map?.nodes || [];
    nodes.forEach((node) => {
      (node?.inLinks || []).forEach((link) => {
        if (!hasMapNodeId(link?.upstreamNodeId?.id)) quality.nodeIdMissing = true;
        const normalLanes = (link?.lanes || []).filter(isNormalVehicleJsonLane);
        if (normalLanes.length && !normalLanes.some(hasJsonStopLine)) quality.stopLineMissing = true;
        normalLanes.forEach((lane) => {
          if (!hasJsonDownstreamNode(lane)) quality.nodeIdMissing = true;
        });
      });
    });
  } catch (error) {
    quality.nodeIdMissing = true;
    quality.stopLineMissing = true;
  }
  return quality;
}

function xmlTagText(xml, tagName) {
  const match = xml.match(new RegExp(`<${tagName}\\b[^>]*>([\\s\\S]*?)</${tagName}\\s*>`));
  return match ? match[1].replace(/<[^>]+>/g, "").trim() : "";
}

function xmlBlocks(xml, tagName) {
  return Array.from(xml.matchAll(new RegExp(`<${tagName}\\b[^>]*>[\\s\\S]*?</${tagName}\\s*>`, "g")), (match) => match[0]);
}

function isNormalVehicleXmlLane(laneXml) {
  const laneId = Number(xmlTagText(laneXml, "laneID") || 0);
  const laneType = laneXml.match(/<laneType\b[^>]*>([\s\S]*?)<\/laneType\s*>/)?.[1] || "";
  return /<vehicle\b/i.test(laneType) && !/<crosswalk\b/i.test(laneType) && !xmlTagText(laneXml, "laneTypeAttrVehExt") && (!laneId || laneId < 240);
}

function hasXmlStopLine(laneXml) {
  return xmlBlocks(laneXml, "StopLine").some((stopLineXml) => {
    const center = stopLineXml.match(/<centerPoint\b[^>]*>([\s\S]*?)<\/centerPoint\s*>/)?.[1] || "";
    return hasMapNodeId(xmlTagText(center, "lat")) || hasMapNodeId(xmlTagText(center, "long"));
  });
}

function hasXmlDownstreamNode(laneXml) {
  return xmlBlocks(laneXml, "Connection").some((connectionXml) => {
    const remote = connectionXml.match(/<remoteIntersection\b[^>]*>([\s\S]*?)<\/remoteIntersection\s*>/)?.[1] || "";
    return hasMapNodeId(xmlTagText(remote, "id"));
  });
}

function hasShortXmlDownstreamNode(laneXml) {
  const connections = xmlBlocks(laneXml, "Connection");
  return connections.some((connectionXml) => {
    const remote = connectionXml.match(/<remoteIntersection\b[^>]*>([\s\S]*?)<\/remoteIntersection\s*>/)?.[1] || "";
    return isShortPresentMapNodeId(xmlTagText(remote, "id"));
  });
}

function inspectMapXmlContent(content) {
  const quality = mapQualityEmpty();
  xmlBlocks(content, "Link").forEach((linkXml) => {
    const upstream = linkXml.match(/<upstreamNodeId\b[^>]*>([\s\S]*?)<\/upstreamNodeId\s*>/)?.[1] || "";
    if (!hasMapNodeId(xmlTagText(upstream, "id"))) quality.nodeIdMissing = true;
    if (isShortPresentMapNodeId(xmlTagText(upstream, "id"))) quality.mapXmlNodeIdShort = true;
    const normalLanes = xmlBlocks(linkXml, "Lane").filter(isNormalVehicleXmlLane);
    if (normalLanes.length && !normalLanes.some(hasXmlStopLine)) quality.stopLineMissing = true;
    normalLanes.forEach((laneXml) => {
      if (!hasXmlDownstreamNode(laneXml)) quality.nodeIdMissing = true;
      if (hasShortXmlDownstreamNode(laneXml)) quality.mapXmlNodeIdShort = true;
    });
  });
  return quality;
}

async function inspectMapAssetQuality(asset, fullPath) {
  if (!["map_xml", "map_json"].includes(asset.category)) return null;
  if (!["xml", "json"].includes(asset.ext)) return null;
  try {
    const content = await fs.readFile(fullPath, "utf8");
    if (asset.ext === "json") return inspectMapJsonContent(content);
    return inspectMapXmlContent(content);
  } catch (error) {
    return { nodeIdMissing: true, stopLineMissing: true };
  }
}

async function walkMapAssetFiles(directory, baseDirectory = directory) {
  const entries = await fs.readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    if (entry.name.startsWith(".")) continue;
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await walkMapAssetFiles(fullPath, baseDirectory)));
      continue;
    }
    const ext = extensionOf(entry.name);
    if (!["xml", "json", "xlsx", "xls", "svg", "png", "zip"].includes(ext)) continue;
    const stat = await fs.stat(fullPath);
    const relativePath = path.relative(baseDirectory, fullPath);
    const asset = { ...parseAssetFile(relativePath), size: stat.size, modifiedAt: stat.mtime.toISOString() };
    const quality = await inspectMapAssetQuality(asset, fullPath);
    files.push(quality ? { ...asset, quality } : asset);
  }
  return files;
}

function emptyIntersection(asset) {
  return {
    id: `${asset.district}-${asset.nodeId || asset.intersectionName}`,
    district: asset.district,
    nodeId: asset.nodeId || "-",
    name: asset.intersectionName,
    latestVersion: asset.version || asset.modifiedAt,
    files: [],
    counts: {
      map_xml: 0,
      map_json: 0,
      map_svg: 0,
      rsi_xml: 0,
      rsi_json: 0,
      signal_excel: 0,
      other: 0,
    },
    quality: mapQualityEmpty(),
  };
}

function buildMapAssetIndex(files) {
  const groups = new Map();
  const categoryTotals = {};
  files.forEach((asset) => {
    categoryTotals[asset.category] = (categoryTotals[asset.category] || 0) + 1;
    const key = `${asset.district}::${asset.nodeId || asset.intersectionName}`;
    if (!groups.has(key)) groups.set(key, emptyIntersection(asset));
    const intersection = groups.get(key);
    intersection.files.push(asset);
    if (Object.prototype.hasOwnProperty.call(intersection.counts, asset.category)) {
      intersection.counts[asset.category] += 1;
    } else {
      intersection.counts.other += 1;
    }
    if ((asset.version || asset.modifiedAt) > intersection.latestVersion) intersection.latestVersion = asset.version || asset.modifiedAt;
    mergeMapQuality(intersection.quality, asset.quality);
  });
  const intersections = Array.from(groups.values()).map((intersection) => {
    const hasMap = intersection.counts.map_xml > 0 || intersection.counts.map_json > 0;
    const hasRsi = intersection.counts.rsi_xml > 0 || intersection.counts.rsi_json > 0;
    const hasSignal = intersection.counts.signal_excel > 0;
    const hasPreview = intersection.counts.map_svg > 0;
    const missing = [
      !hasMap && "MAP",
      !hasRsi && "RSI",
      !hasSignal && "信号机",
      !hasPreview && "SVG预览",
      intersection.quality.nodeIdMissing && "上下游 NodeID",
      intersection.quality.mapXmlNodeIdShort && "MAP-XML上下游 NodeID<5位",
      intersection.quality.stopLineMissing && "某方向车道停止线",
    ].filter(Boolean);
    const mapXmlPreview = intersection.files.find((file) => file.category === "map_svg" && file.folder === "map_xml");
    const mapJsonPreview = intersection.files.find((file) => file.category === "map_svg" && file.folder === "map_json");
    const preview = mapXmlPreview || mapJsonPreview || intersection.files.find((file) => file.category === "map_svg");
    return {
      ...intersection,
      completeness: 6 - missing.length,
      status: missing.length ? "待补全" : "完整",
      missing,
      previewPath: preview?.relativePath || "",
      mapXmlSvgPath: mapXmlPreview?.relativePath || "",
      mapJsonSvgPath: mapJsonPreview?.relativePath || "",
      fileTotal: intersection.files.length,
      files: intersection.files.sort((a, b) => a.category.localeCompare(b.category) || b.modifiedAt.localeCompare(a.modifiedAt)),
    };
  });
  intersections.sort((a, b) => a.status.localeCompare(b.status, "zh-CN") || a.district.localeCompare(b.district, "zh-CN") || String(a.nodeId).localeCompare(String(b.nodeId), "zh-CN"));
  return {
    root: MAP_ASSET_DIR,
    generatedAt: new Date().toISOString(),
    summary: {
      fileTotal: files.length,
      intersectionTotal: intersections.length,
      districtTotals: intersections.reduce((acc, item) => {
        acc[item.district] = (acc[item.district] || 0) + 1;
        return acc;
      }, {}),
      categoryTotals,
      completeTotal: intersections.filter((item) => item.status === "完整").length,
      incompleteTotal: intersections.filter((item) => item.status !== "完整").length,
    },
    intersections,
  };
}

async function loadMapAssetIndex() {
  const now = Date.now();
  if (mapAssetCache && now - mapAssetCache.createdAt < MAP_ASSET_CACHE_TTL_MS) return mapAssetCache.payload;
  const files = await walkMapAssetFiles(MAP_ASSET_DIR);
  const payload = buildMapAssetIndex(files);
  mapAssetCache = { createdAt: now, payload };
  return payload;
}

async function handleMapAssets(request, response, requestUrl) {
  if (request.method !== "GET") {
    response.writeHead(405, { Allow: "GET" });
    response.end();
    return;
  }
  try {
    const payload = await loadMapAssetIndex();
    return sendJson(response, 200, payload);
  } catch (error) {
    console.error(error);
    return sendJson(response, 500, { error: "map_asset_scan_failed", message: error.message });
  }
}

async function handleMapAssetFile(request, response, requestUrl) {
  if (request.method !== "GET") {
    response.writeHead(405, { Allow: "GET" });
    response.end();
    return;
  }
  const relativePath = requestUrl.searchParams.get("path") || "";
  const filePath = mapAssetFilePath(relativePath);
  if (!filePath) return sendJson(response, 403, { error: "forbidden" });
  try {
    const body = await fs.readFile(filePath);
    response.writeHead(200, { "Content-Type": MIME_TYPES[path.extname(filePath)] || "application/octet-stream" });
    response.end(body);
  } catch (error) {
    if (error.code === "ENOENT") return sendJson(response, 404, { error: "not_found" });
    console.error(error);
    return sendJson(response, 500, { error: "map_asset_file_read_failed" });
  }
}

async function handleContractFile(request, response, requestUrl) {
  if (request.method !== "GET") {
    response.writeHead(405, { Allow: "GET" });
    response.end();
    return;
  }
  const rawPath = requestUrl.searchParams.get("path") || "";
  const filePath = contractFilePath(rawPath);
  if (!filePath) return sendJson(response, 403, { error: "forbidden" });
  try {
    const stat = await fs.stat(filePath);
    if (!stat.isFile()) return sendJson(response, 404, { error: "not_found" });
    const body = await fs.readFile(filePath);
    response.writeHead(200, {
      "Content-Type": MIME_TYPES[path.extname(filePath)] || "application/octet-stream",
      "Content-Disposition": contentDispositionInline(filePath),
      "Cache-Control": "private, max-age=60",
    });
    response.end(body);
  } catch (error) {
    if (error.code === "ENOENT") return sendJson(response, 404, { error: "not_found" });
    console.error(error);
    return sendJson(response, 500, { error: "contract_file_read_failed" });
  }
}

async function handleContractFileOpen(request, response, requestUrl) {
  if (request.method !== "POST" && request.method !== "GET") {
    response.writeHead(405, { Allow: "POST, GET" });
    response.end();
    return;
  }
  const payload = request.method === "POST" ? await readRequestJson(request).catch(() => null) : null;
  const rawPath = payload?.path || requestUrl.searchParams.get("path") || "";
  const filePath = contractFilePath(rawPath);
  if (!filePath) return sendJson(response, 403, { error: "forbidden" });
  try {
    const stat = await fs.stat(filePath);
    if (!stat.isFile()) return sendJson(response, 404, { error: "not_found" });
    await openWithDefaultApp(filePath);
    return sendJson(response, 200, { ok: true, fileName: path.basename(filePath) });
  } catch (error) {
    if (error.code === "ENOENT") return sendJson(response, 404, { error: "not_found" });
    console.error(error);
    return sendJson(response, 500, { error: "contract_file_open_failed", message: error.message });
  }
}

async function handleMapAssetExport(request, response, requestUrl) {
  if (request.method !== "GET") {
    response.writeHead(405, { Allow: "GET" });
    response.end();
    return;
  }
  const assetId = requestUrl.searchParams.get("id") || "";
  try {
    const payload = await loadMapAssetIndex();
    const asset = payload.intersections.find((item) => item.id === assetId);
    if (!asset) return sendJson(response, 404, { error: "not_found" });
    const assetBaseName = `${sanitizeArchiveName(asset.nodeId)}_${sanitizeArchiveName(asset.name)}`;
    const archiveRoot = `${assetBaseName} 地图资产文件`;
    const entries = [];
    for (const file of asset.files) {
      const filePath = mapAssetFilePath(file.relativePath);
      if (!filePath) continue;
      entries.push({
        name: `${archiveRoot}/${file.relativePath.split(path.sep).join("/")}`,
        body: await fs.readFile(filePath),
        modifiedAt: file.modifiedAt,
      });
    }
    if (!entries.length) return sendJson(response, 404, { error: "empty_asset" });
    const archiveName = `${assetBaseName}.zip`;
    const body = buildZip(entries);
    response.writeHead(200, {
      "Content-Type": "application/zip",
      "Content-Length": body.length,
      "Content-Disposition": `attachment; filename="map-assets.zip"; filename*=UTF-8''${encodeURIComponent(archiveName)}`,
    });
    response.end(body);
  } catch (error) {
    console.error(error);
    return sendJson(response, 500, { error: "map_asset_export_failed", message: error.message });
  }
}

async function handleRoadsideState(request, response) {
  if (request.method === "GET") {
    try {
      const data = await fs.readFile(STATE_FILE, "utf8");
      response.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      response.end(data);
    } catch (error) {
      if (error.code === "ENOENT") return sendJson(response, 404, { error: "not_found" });
      console.error(error);
      return sendJson(response, 500, { error: "read_failed" });
    }
    return;
  }

  if (request.method === "POST") {
    try {
      const payload = await readRequestJson(request);
      if (!validateStatePayload(payload)) return sendJson(response, 400, { error: "invalid_payload" });
      await fs.mkdir(STORAGE_DIR, { recursive: true });
      const saved = {
        currentDate: payload.currentDate,
        currentRows: payload.currentRows,
        archives: payload.archives,
        updatedAt: new Date().toISOString(),
      };
      await fs.writeFile(STATE_FILE, `${JSON.stringify(saved, null, 2)}\n`, "utf8");
      return sendJson(response, 200, { ok: true, updatedAt: saved.updatedAt });
    } catch (error) {
      console.error(error);
      return sendJson(response, 500, { error: "write_failed" });
    }
  }

  response.writeHead(405, { Allow: "GET, POST" });
  response.end();
}


function sanitizeRelativeUploadPath(value) {
  const parts = String(value || "")
    .split(/[\\/]+/)
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => part.replace(/[<>:"|?*\x00-\x1f]/g, "_").slice(0, 120));
  return parts.join(path.sep) || "contract-file";
}

function contractImportDirection(value) {
  return value === "back" || value === "back_procurement" ? "back_procurement" : "front_sales";
}

function inferContractImportDirection(requestedDirection, relativePath) {
  const text = String(relativePath || "");
  if (/后向采购合同|后向|采购合同|采购/.test(text) && !/前向销售合同|前向销售|销售合同/.test(text)) return "back_procurement";
  if (/前向销售合同|前向销售|前向|销售合同/.test(text) && !/后向采购合同|后向/.test(text)) return "front_sales";
  return requestedDirection;
}

function contractImportSubdir(direction) {
  return direction === "back_procurement" ? "后向采购合同" : "前向销售合同";
}

async function handleContractImportUpload(request, response, requestUrl) {
  if (request.method !== "POST") {
    response.writeHead(405, { Allow: "POST" });
    response.end();
    return;
  }
  try {
    const requestedDirection = contractImportDirection(requestUrl.searchParams.get("direction"));
    const fileName = sanitizeRelativeUploadPath(requestUrl.searchParams.get("fileName") || "contract-file");
    const direction = inferContractImportDirection(requestedDirection, fileName);
    const targetPath = path.normalize(path.join(CONTRACT_MANUAL_IMPORT_DIR, contractImportSubdir(direction), fileName));
    const root = path.normalize(path.join(CONTRACT_MANUAL_IMPORT_DIR, contractImportSubdir(direction)));
    if (!targetPath.startsWith(root)) return sendJson(response, 400, { error: "invalid_file_path" });
    const chunks = [];
    for await (const chunk of request) chunks.push(chunk);
    const body = Buffer.concat(chunks);
    if (!body.length) return sendJson(response, 400, { error: "empty_file" });
    await fs.mkdir(path.dirname(targetPath), { recursive: true });
    await fs.writeFile(targetPath, body);
    const sha256 = crypto.createHash("sha256").update(body).digest("hex");
    const restoreResult = await restoreContractExclusionForImport({ sourcePath: targetPath, fileName, sha256 });
    sendJson(response, 200, {
      ok: true,
      direction,
      fileName,
      path: targetPath,
      size: body.length,
      sha256,
      restoredExcludedContractCount: restoreResult.restoredCount,
      restoredExcludedContracts: restoreResult.restored.map((item) => ({
        contractId: item.contractId || "",
        fileName: item.fileName || "",
        removedAt: item.removedAt || "",
      })),
    });
  } catch (error) {
    console.error(error);
    sendJson(response, 500, { error: "contract_import_upload_failed", message: error.message });
  }
}

function runExecFile(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    execFile(command, args, { ...options, maxBuffer: 20 * 1024 * 1024 }, (error, stdout, stderr) => {
      if (error) {
        error.stdout = stdout;
        error.stderr = stderr;
        reject(error);
      } else {
        resolve({ stdout, stderr });
      }
    });
  });
}

async function handleContractRelationshipsRebuild(request, response) {
  if (request.method !== "POST") {
    response.writeHead(405, { Allow: "POST" });
    response.end();
    return;
  }
  try {
    const candidateSources = [
      "/Users/tt2000/Documents/天安智联/0000无锡市车路云一体化项目/0.合同/盖章版合同 WORD 版",
      "/Users/tt2000/Documents/天安智联/0000无锡市车路云一体化项目/0.合同/盖章版合同",
      path.join(CONTRACT_MANUAL_IMPORT_DIR, "前向销售合同"),
      path.join(CONTRACT_MANUAL_IMPORT_DIR, "后向采购合同"),
    ];
    const rebuildArgs = [CONTRACT_REBUILD_SCRIPT, "--exclude-file", CONTRACT_EXCLUSIONS_FILE];
    for (const source of candidateSources) {
      try {
        await fs.access(source);
        rebuildArgs.push("--source", source);
      } catch (error) {
        // Ignore import folders that do not exist yet.
      }
    }
    const result = await runExecFile(PYTHON_BIN, rebuildArgs, { cwd: path.join(ROOT_DIR, ".."), env: { ...process.env, PYTHON_BIN } });
    let payload = { ok: true, stdout: result.stdout.trim(), stderr: result.stderr.trim() };
    try {
      payload = { ok: true, ...JSON.parse(result.stdout) };
      const exclusions = await readContractExclusions();
      await applyContractExclusionsToGeneratedData(exclusions.exclusions);
      payload.excludedContractCount = exclusions.exclusions.length;
      if (result.stderr.trim()) payload.stderr = result.stderr.trim();
    } catch (error) {
      // Keep raw stdout when the script prints non-JSON diagnostics.
    }
    sendJson(response, 200, payload);
  } catch (error) {
    console.error(error);
    sendJson(response, 500, {
      error: "contract_relationships_rebuild_failed",
      message: error.message,
      stdout: String(error.stdout || "").slice(0, 4000),
      stderr: String(error.stderr || "").slice(0, 4000),
    });
  }
}

async function readDocumentAssetReport() {
  const candidates = [DOCUMENT_ASSET_REPORT, DOCUMENT_ASSET_SAMPLE_REPORT];
  for (const candidate of candidates) {
    try {
      const body = await fs.readFile(candidate, "utf8");
      return { ...JSON.parse(body), reportPath: candidate };
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
    }
  }
  return {
    batchNo: "EMPTY",
    projectCode: "wuxi-cv2x",
    generatedAt: new Date().toISOString(),
    reportPath: "",
    summary: {
      fileCount: 0,
      totalSizeBytes: 0,
      duplicateFileCount: 0,
      unlinkedFileCount: 0,
      categoryCounts: {},
      extensionCounts: {},
      errorCount: 0,
    },
    records: [],
    errors: [],
  };
}

async function handleDocumentAssets(request, response) {
  if (request.method !== "GET") {
    response.writeHead(405, { Allow: "GET" });
    response.end();
    return;
  }
  try {
    const report = await readDocumentAssetReport();
    sendJson(response, 200, report);
  } catch (error) {
    console.error(error);
    sendJson(response, 500, { error: "document_assets_read_failed" });
  }
}

async function serveStatic(request, response, pathname) {
  const requestedPath = pathname === "/" ? "/index.html" : pathname;
  const filePath = path.normalize(path.join(ROOT_DIR, fileURLToPath(`file://${requestedPath}`)));
  if (!filePath.startsWith(ROOT_DIR)) return sendJson(response, 403, { error: "forbidden" });
  try {
    const body = await fs.readFile(filePath);
    response.writeHead(200, { "Content-Type": MIME_TYPES[path.extname(filePath)] || "application/octet-stream" });
    response.end(body);
  } catch (error) {
    if (error.code === "ENOENT") return sendJson(response, 404, { error: "not_found" });
    console.error(error);
    return sendJson(response, 500, { error: "static_read_failed" });
  }
}

const server = http.createServer(async (request, response) => {
  const requestUrl = new URL(request.url, `http://${request.headers.host || `127.0.0.1:${PORT}`}`);
  if (requestUrl.pathname === "/api/health") {
    sendNoContent(response);
    return;
  }
  if (requestUrl.pathname === "/api/roadside-status-state") {
    await handleRoadsideState(request, response);
    return;
  }
  if (requestUrl.pathname === "/api/map-assets") {
    await handleMapAssets(request, response, requestUrl);
    return;
  }
  if (requestUrl.pathname === "/api/map-assets/file") {
    await handleMapAssetFile(request, response, requestUrl);
    return;
  }
  if (requestUrl.pathname === "/api/map-assets/export") {
    await handleMapAssetExport(request, response, requestUrl);
    return;
  }
  if (requestUrl.pathname === "/api/contract-file") {
    await handleContractFile(request, response, requestUrl);
    return;
  }
  if (requestUrl.pathname === "/api/contract-file/open") {
    await handleContractFileOpen(request, response, requestUrl);
    return;
  }
  if (requestUrl.pathname === "/api/contract-import") {
    await handleContractImportUpload(request, response, requestUrl);
    return;
  }
  if (requestUrl.pathname === "/api/contract-system-remove") {
    await handleContractSystemRemove(request, response);
    return;
  }
  if (requestUrl.pathname === "/api/contract-relationships/rebuild") {
    await handleContractRelationshipsRebuild(request, response);
    return;
  }
  if (requestUrl.pathname === "/api/document-assets") {
    await handleDocumentAssets(request, response);
    return;
  }
  await serveStatic(request, response, decodeURIComponent(requestUrl.pathname));
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`Prototype server running at http://127.0.0.1:${PORT}`);
  console.log(`Roadside status state file: ${STATE_FILE}`);
});
