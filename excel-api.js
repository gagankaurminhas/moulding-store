const ExcelAPI = (() => {
  const keys = {
    drive: "mouldingWorkbookDriveId",
    item: "mouldingWorkbookItemId"
  };

  function ids() {
    return {
      driveId: localStorage.getItem(keys.drive) || APP_CONFIG.workbookDriveId,
      itemId: localStorage.getItem(keys.item) || APP_CONFIG.workbookItemId
    };
  }

  async function request(path, options = {}) {
    const token = await getAccessToken();
    const res = await fetch(APP_CONFIG.graphBaseUrl + path, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(options.headers || {})
      }
    });
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`${res.status}: ${body || res.statusText}`);
    }
    if (res.status === 204) return null;
    return res.json();
  }

  function base() {
    const {driveId, itemId} = ids();
    if (!driveId || !itemId) {
      throw new Error("Workbook is not connected. Open Settings and use Find My Workbook.");
    }
    return `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook`;
  }

  async function allPages(path) {
    let url = path;
    const out = [];
    while (url) {
      const data = await request(url);
      if (Array.isArray(data.value)) out.push(...data.value);
      url = data["@odata.nextLink"]
        ? data["@odata.nextLink"].replace(APP_CONFIG.graphBaseUrl, "")
        : null;
    }
    return out;
  }

  async function rows(tableName) {
    const list = await allPages(`${base()}/tables/${encodeURIComponent(tableName)}/rows`);
    return list.map(x => x.values?.[0] || []);
  }

  async function headers(tableName) {
    const data = await request(`${base()}/tables/${encodeURIComponent(tableName)}/range`);
    return (data.values?.[0] || []).map(String);
  }

  function objects(header, rowValues) {
    return rowValues.map(row => {
      const o = {};
      header.forEach((h, i) => o[h] = row[i] ?? "");
      return o;
    });
  }

  async function tableObjects(tableName) {
    const [h, r] = await Promise.all([headers(tableName), rows(tableName)]);
    return objects(h, r);
  }

  async function addRow(tableName, obj) {
    const h = await headers(tableName);
    const row = h.map(k => obj[k] ?? "");
    return request(`${base()}/tables/${encodeURIComponent(tableName)}/rows/add`, {
      method: "POST",
      body: JSON.stringify({values: [row]})
    });
  }

  async function updateRow(tableName, index, obj) {
    const h = await headers(tableName);
    const row = h.map(k => obj[k] ?? "");
    return request(
      `${base()}/tables/${encodeURIComponent(tableName)}/rows/${index}`,
      {method: "PATCH", body: JSON.stringify({values: [row])}
    );
  }

  async function findWorkbook() {
    const token = await getAccessToken();
    const res = await fetch(
      `${APP_CONFIG.graphBaseUrl}/me/drive/root/search(q='${encodeURIComponent(APP_CONFIG.workbookFileName)}')`,
      {headers: {Authorization: `Bearer ${token}`}}
    );
    if (!res.ok) throw new Error("Could not search OneDrive.");
    const data = await res.json();
    const item = (data.value || []).find(x => x.name === APP_CONFIG.workbookFileName);
    if (!item) throw new Error(`Could not find ${APP_CONFIG.workbookFileName} in OneDrive.`);
    const driveId = item.parentReference?.driveId;
    const itemId = item.id;
    if (!driveId || !itemId) throw new Error("Workbook was found but its OneDrive IDs were unavailable.");
    localStorage.setItem(keys.drive, driveId);
    localStorage.setItem(keys.item, itemId);
    return {driveId, itemId};
  }

  async function load() {
    const names = APP_CONFIG.tables;
    const [
      drivers, trucks, orders, schedules, routeStops, deliveryStatus, loading, issues
    ] = await Promise.all([
      tableObjects(names.drivers),
      tableObjects(names.trucks),
      tableObjects(names.orders),
      tableObjects(names.schedules),
      tableObjects(names.routeStops),
      tableObjects(names.deliveryStatus),
      tableObjects(names.loading),
      tableObjects(names.issues)
    ]);
    return {drivers, trucks, orders, schedules, routeStops, deliveryStatus, loading, issues};
  }

  async function addOrder(order) {
    return addRow(APP_CONFIG.tables.orders, order);
  }

  async function addSchedule(schedule) {
    return addRow(APP_CONFIG.tables.schedules, schedule);
  }

  async function addStop(stop) {
    return addRow(APP_CONFIG.tables.routeStops, stop);
  }

  async function updateOrderById(orderId, patch) {
    const h = await headers(APP_CONFIG.tables.orders);
    const r = await rows(APP_CONFIG.tables.orders);
    const idIndex = h.indexOf("OrderID");
    if (idIndex < 0) throw new Error("OrdersTable is missing OrderID.");
    const idx = r.findIndex(x => String(x[idIndex]) === String(orderId));
    if (idx < 0) throw new Error("Order not found.");
    const current = {};
    h.forEach((k, i) => current[k] = r[idx][i] ?? "");
    Object.assign(current, patch);
    return updateRow(APP_CONFIG.tables.orders, idx, current);
  }

  async function updateScheduleById(scheduleId, patch) {
    const h = await headers(APP_CONFIG.tables.schedules);
    const r = await rows(APP_CONFIG.tables.schedules);
    const idIndex = h.indexOf("ScheduleID");
    if (idIndex < 0) throw new Error("DailyScheduleTable is missing ScheduleID.");
    const idx = r.findIndex(x => String(x[idIndex]) === String(scheduleId));
    if (idx < 0) throw new Error("Schedule not found.");
    const current = {};
    h.forEach((k, i) => current[k] = r[idx][i] ?? "");
    Object.assign(current, patch);
    return updateRow(APP_CONFIG.tables.schedules, idx, current);
  }

  async function updateStopById(stopId, patch) {
    const h = await headers(APP_CONFIG.tables.routeStops);
    const r = await rows(APP_CONFIG.tables.routeStops);
    const idIndex = h.indexOf("StopID");
    if (idIndex < 0) throw new Error("RouteStopsTable is missing StopID.");
    const idx = r.findIndex(x => String(x[idIndex]) === String(stopId));
    if (idx < 0) throw new Error("Stop not found.");
    const current = {};
    h.forEach((k, i) => current[k] = r[idx][i] ?? "");
    Object.assign(current, patch);
    return updateRow(APP_CONFIG.tables.routeStops, idx, current);
  }

  async function setStatus(status) {
    const s = {
      StatusID: `DS-${Date.now()}`,
      ...status
    };
    return addRow(APP_CONFIG.tables.deliveryStatus, s);
  }

  async function findAndSaveWorkbook() {
    return findWorkbook();
  }

  return {
    load, addOrder, addSchedule, addStop, updateOrderById,
    updateScheduleById, updateStopById, setStatus, findAndSaveWorkbook
  };
})();
