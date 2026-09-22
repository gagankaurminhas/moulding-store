const ExcelAPI = (() => {
  const DRIVE_KEY = "mouldingWorkbookDriveId";
  const ITEM_KEY = "mouldingWorkbookItemId";

  function workbookIds() {
    return {
      driveId: localStorage.getItem(DRIVE_KEY) || APP_CONFIG.workbookDriveId,
      itemId: localStorage.getItem(ITEM_KEY) || APP_CONFIG.workbookItemId
    };
  }

  async function request(path) {
    const token = await getAccessToken();
    const response = await fetch(APP_CONFIG.graphBaseUrl + path, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json"
      }
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`${response.status}: ${text || response.statusText}`);
    }
    return response.json();
  }

  function workbookBase() {
    const {driveId, itemId} = workbookIds();
    if (!driveId || !itemId) {
      throw new Error("Workbook is not connected. Open Settings and choose Find Workbook.");
    }
    return `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook`;
  }

  async function allPages(path) {
    let url = path;
    const output = [];
    while (url) {
      const data = await request(url);
      if (Array.isArray(data.value)) output.push(...data.value);
      url = data["@odata.nextLink"]
        ? data["@odata.nextLink"].replace(APP_CONFIG.graphBaseUrl, "")
        : null;
    }
    return output;
  }

  async function tableData(tableName) {
    const base = workbookBase();
    const range = await request(`${base}/tables/${encodeURIComponent(tableName)}/range`);
    const values = range.values || [];
    const header = (values[0] || []).map(v => String(v));
    const rows = await allPages(`${base}/tables/${encodeURIComponent(tableName)}/rows`);
    return rows.map(item => {
      const row = item.values?.[0] || [];
      const object = {};
      header.forEach((key, i) => object[key] = row[i] ?? "");
      return object;
    });
  }

  async function load() {
    const t = APP_CONFIG.tables;
    const [orders, deliveryStatus, drivers, trucks, schedules, routeStops] =
      await Promise.all([
        tableData(t.orders),
        tableData(t.deliveryStatus),
        tableData(t.drivers),
        tableData(t.trucks),
        tableData(t.schedules),
        tableData(t.routeStops)
      ]);

    return {orders, deliveryStatus, drivers, trucks, schedules, routeStops};
  }

  async function findWorkbook() {
    const token = await getAccessToken();
    const q = encodeURIComponent(APP_CONFIG.workbookFileName);
    const response = await fetch(
      `${APP_CONFIG.graphBaseUrl}/me/drive/root/search(q='${q}')`,
      {headers: {Authorization: `Bearer ${token}`, Accept: "application/json"}}
    );

    if (!response.ok) throw new Error("Could not search OneDrive for the workbook.");
    const data = await response.json();
    const item = (data.value || []).find(x => x.name === APP_CONFIG.workbookFileName);
    if (!item) throw new Error(`Could not find ${APP_CONFIG.workbookFileName}.`);

    const driveId = item.parentReference?.driveId;
    const itemId = item.id;
    if (!driveId || !itemId) throw new Error("Workbook was found but its IDs were unavailable.");

    localStorage.setItem(DRIVE_KEY, driveId);
    localStorage.setItem(ITEM_KEY, itemId);
    return {driveId, itemId};
  }

  return {load, findWorkbook};
})();
