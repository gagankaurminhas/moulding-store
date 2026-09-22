const excelColumnCache = new Map();
let workbookInfoCache = null;

async function graphRequest(pathOrUrl, options = {}, retryCount = 0) {
  const token = await getAccessToken();
  const url = pathOrUrl.startsWith("http")
    ? pathOrUrl
    : `${APP_CONFIG.graphBaseUrl}${pathOrUrl}`;

  const headers = new Headers(options.headers || {});
  headers.set("Authorization", `Bearer ${token}`);
  if (options.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }

  const response = await fetch(url, { ...options, headers });
  const contentType = response.headers.get("content-type") || "";
  const text = await response.text();
  let data = null;

  if (text) {
    data = contentType.includes("application/json") ? JSON.parse(text) : text;
  }

  if (!response.ok) {
    if ((response.status === 429 || response.status === 503 || response.status === 504) && retryCount < 2) {
      const retryAfter = Number(response.headers.get("Retry-After")) || (retryCount + 1) * 2;
      await new Promise(resolve => setTimeout(resolve, retryAfter * 1000));
      return graphRequest(pathOrUrl, options, retryCount + 1);
    }

    const message = data?.error?.message || `Microsoft Graph returned HTTP ${response.status}.`;
    throw new Error(message);
  }

  return data;
}

function workbookConfigured() {
  return Boolean(APP_CONFIG.workbookDriveId && APP_CONFIG.workbookItemId);
}

function workbookBasePath() {
  if (!workbookConfigured()) {
    throw new Error("Workbook Drive ID and Item ID are not configured yet.");
  }

  return `/drives/${encodeURIComponent(APP_CONFIG.workbookDriveId)}/items/${encodeURIComponent(APP_CONFIG.workbookItemId)}/workbook`;
}

async function resolveOwnerWorkbook() {
  const path = `/me/drive/root:/${encodeURIComponent(APP_CONFIG.workbookFileName)}`;
  const info = await graphRequest(path);

  const driveId = info?.parentReference?.driveId;
  const itemId = info?.id;

  if (!driveId || !itemId) {
    throw new Error("The workbook was found, but Microsoft Graph did not return its Drive ID and Item ID.");
  }

  workbookInfoCache = {
    driveId,
    itemId,
    name: info.name,
    webUrl: info.webUrl || ""
  };

  return workbookInfoCache;
}

async function getWorkbookInfo() {
  if (workbookInfoCache) return workbookInfoCache;

  if (!workbookConfigured()) {
    return null;
  }

  const path = `/drives/${encodeURIComponent(APP_CONFIG.workbookDriveId)}/items/${encodeURIComponent(APP_CONFIG.workbookItemId)}?$select=id,name,webUrl,parentReference,file`;
  const info = await graphRequest(path);

  workbookInfoCache = {
    driveId: APP_CONFIG.workbookDriveId,
    itemId: APP_CONFIG.workbookItemId,
    name: info.name,
    webUrl: info.webUrl || ""
  };

  return workbookInfoCache;
}

async function fetchAll(url) {
  let next = url;
  const all = [];

  while (next) {
    const data = await graphRequest(next);
    if (Array.isArray(data?.value)) all.push(...data.value);
    next = data?.["@odata.nextLink"] || null;
  }

  return all;
}

async function getTableColumns(tableName) {
  if (excelColumnCache.has(tableName)) {
    return excelColumnCache.get(tableName);
  }

  const table = encodeURIComponent(tableName);
  const url = `${workbookBasePath()}/tables/${table}/columns?$select=id,name,index&$top=100`;
  const columns = await fetchAll(url);
  columns.sort((a, b) => Number(a.index ?? 0) - Number(b.index ?? 0));

  const names = columns.map(c => c.name);
  excelColumnCache.set(tableName, names);
  return names;
}

async function getTableRows(tableName) {
  const table = encodeURIComponent(tableName);
  const url = `${workbookBasePath()}/tables/${table}/rows?$top=5000`;
  const rows = await fetchAll(url);
  const columns = await getTableColumns(tableName);

  return rows.map(row => {
    const values = row?.values?.[0] || [];
    const object = {};
    columns.forEach((column, index) => {
      object[column] = values[index] ?? "";
    });
    object.__rowIndex = Number(row.index ?? 0);
    return object;
  });
}

async function addTableRow(tableName, rowObject) {
  const columns = await getTableColumns(tableName);
  const values = [columns.map(column => rowObject[column] ?? "")];
  const table = encodeURIComponent(tableName);

  return graphRequest(`${workbookBasePath()}/tables/${table}/rows/add`, {
    method: "POST",
    body: JSON.stringify({ index: null, values })
  });
}

async function updateTableRow(tableName, rowIndex, rowObject) {
  const columns = await getTableColumns(tableName);
  const values = [columns.map(column => rowObject[column] ?? "")];
  const table = encodeURIComponent(tableName);

  return graphRequest(`${workbookBasePath()}/tables/${table}/rows/${encodeURIComponent(rowIndex)}`, {
    method: "PATCH",
    body: JSON.stringify({ values })
  });
}

async function testWorkbookConnection() {
  if (!workbookConfigured()) {
    throw new Error("Configure the workbook Drive ID and Item ID first.");
  }

  const info = await getWorkbookInfo();
  const tables = await fetchAll(`${workbookBasePath()}/tables?$select=id,name&$top=100`);

  return {
    info,
    tables: tables.map(t => t.name)
  };
}
