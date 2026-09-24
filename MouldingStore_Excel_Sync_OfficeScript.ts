/**
 * Moulding Store Fleet Command -> Excel Online sync
 *
 * Power Automate passes one string parameter named payloadJson.
 * The script upserts the operational record instead of deleting history.
 * Sheets created/updated:
 *   Deliveries
 *   Trucks
 *   Drivers
 *   DeliveryEvents
 *
 * Recommended workbook: MouldingStore_Delivery_Control_Template.xlsx
 */
function main(workbook: ExcelScript.Workbook, payloadJson: string) {
  const payload = JSON.parse(payloadJson || "{}");
  const deliveries = Array.isArray(payload.deliveries) ? payload.deliveries : [];
  const trucks = Array.isArray(payload.trucks) ? payload.trucks : [];
  const drivers = Array.isArray(payload.drivers) ? payload.drivers : [];
  const events = Array.isArray(payload.events) ? payload.events : [];

  upsertSheet(workbook, "Deliveries", [
    "DeliveryID","Date","Customer","Phone","Address","City","Province","PostalCode",
    "VehicleType","WeightTons","ServiceMinutes","WindowStart","WindowEnd","Priority",
    "Truck","Driver","Stop","Status","Latitude","Longitude","Notes","CreatedAt","UpdatedAt"
  ], deliveries.map((d: any) => [
    text(d.id), text(d.date), text(d.customer), text(d.phone), text(d.address), text(d.city), "Alberta", text(d.postal),
    text(d.vehicleType), num(d.weight), num(d.serviceMin), text(d.windowStart), text(d.windowEnd), num(d.priority),
    text(d.truck), text(d.driver), num(d.stop), text(d.status), num(d.lat), num(d.lng), text(d.notes), text(d.createdAt), text(d.updatedAt)
  ]), "DeliveriesTable", "DeliveryID");

  replaceSheet(workbook, "Trucks", ["TruckID","Type","CapacityTons","DriverID","DriverName","Active","Color"],
    trucks.map((t: any) => [text(t.id), text(t.type), num(t.capacity), text(t.driverId), text(t.driver), bool(t.active), text(t.color)]), "TrucksTable");

  replaceSheet(workbook, "Drivers", ["DriverID","DriverName","Phone","DefaultTruck","Active"],
    drivers.map((d: any) => [text(d.id), text(d.name), text(d.phone), text(d.defaultTruckId), bool(d.active)]), "DriversTable");

  appendUniqueEvents(workbook, events);

  return {
    syncedAt: new Date().toISOString(),
    deliveries: deliveries.length,
    trucks: trucks.length,
    drivers: drivers.length,
    events: events.length
  };
}

function upsertSheet(workbook: ExcelScript.Workbook, name: string, headers: string[], incoming: (string|number|boolean)[][], tableName: string, keyHeader: string) {
  let sheet = workbook.getWorksheet(name);
  if (!sheet) sheet = workbook.addWorksheet(name);

  const existing = readTable(sheet, tableName, headers);
  const keyIndex = headers.indexOf(keyHeader);
  const map = new Map<string, (string|number|boolean)[]>();
  for (const row of existing) {
    const key = String(row[keyIndex] ?? "");
    if (key) map.set(key, row);
  }
  for (const row of incoming) {
    const key = String(row[keyIndex] ?? "");
    if (key) map.set(key, row);
  }

  writeTable(sheet, headers, Array.from(map.values()), tableName);
}

function replaceSheet(workbook: ExcelScript.Workbook, name: string, headers: string[], rows: (string|number|boolean)[][], tableName: string) {
  let sheet = workbook.getWorksheet(name);
  if (!sheet) sheet = workbook.addWorksheet(name);
  writeTable(sheet, headers, rows, tableName);
}

function appendUniqueEvents(workbook: ExcelScript.Workbook, incoming: any[]) {
  const headers = ["EventID","Timestamp","Date","Type","DeliveryID","Truck","Driver","Status","Details"];
  let sheet = workbook.getWorksheet("DeliveryEvents");
  if (!sheet) sheet = workbook.addWorksheet("DeliveryEvents");
  const table = sheet.getTables()[0];
  const existing: (string|number|boolean)[][] = table ? (table.getRange().getValues().slice(1) as (string|number|boolean)[][]) : [];
  const seen = new Set(existing.map(r => String(r[0] ?? "")));
  for (const e of incoming) {
    const id = text(e.id);
    if (!id || seen.has(id)) continue;
    existing.push([id,text(e.time),text(e.time).slice(0,10),text(e.type),text(e.jobId),text(e.data?.truckId),text(e.data?.driverId),text(e.data?.status),JSON.stringify(e.data || {})]);
    seen.add(id);
  }
  writeTable(sheet, headers, existing, "DeliveryEventsTable");
}

function readTable(sheet: ExcelScript.Worksheet, tableName: string, headers: string[]): (string|number|boolean)[][] {
  const table = sheet.getTables().find(t => t.getName() === tableName) || sheet.getTables()[0];
  if (!table) return [];
  return table.getRange().getValues().slice(1) as (string|number|boolean)[][];
}

function writeTable(sheet: ExcelScript.Worksheet, headers: string[], rows: (string|number|boolean)[][], tableName: string) {
  sheet.getTables().forEach(t => t.delete());
  const used = sheet.getUsedRange();
  if (used) used.clear(ExcelScript.ClearApplyTo.all);

  const values = [headers, ...rows];
  const range = sheet.getRangeByIndexes(0, 0, Math.max(values.length, 1), headers.length);
  range.setValues(values.length ? values : [headers]);
  const header = sheet.getRangeByIndexes(0, 0, 1, headers.length);
  header.getFormat().getFont().setBold(true);
  header.getFormat().getFill().setColor("#202522");
  header.getFormat().getFont().setColor("#FFFFFF");
  const table = sheet.addTable(range, true);
  table.setName(tableName);
  range.getFormat().autofitColumns();
  sheet.getFreezePanes().freezeRows(1);
}

function text(v: any): string { return v === null || v === undefined ? "" : String(v); }
function num(v: any): string|number { if (v === null || v === undefined || v === "") return ""; const n = Number(v); return Number.isFinite(n) ? n : text(v); }
function bool(v: any): boolean { return v === true || v === "true" || v === 1; }
