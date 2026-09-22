const PAGE_META = {
  dashboard: ["Dashboard", "Today's delivery operations"],
  orders: ["Orders", "Customer orders and delivery readiness"],
  backorders: ["Backorders", "Items waiting for purchasing or receiving"],
  purchasing: ["Purchasing / PO", "Purchase orders and suppliers"],
  receiving: ["Receiving", "Incoming purchase orders"],
  schedule: ["Delivery Schedule", "Assign ready orders to routes"],
  drivers: ["Drivers & Trucks", "Driver and truck assignments"],
  routes: ["Routes", "Route planning and balancing"],
  issues: ["Delivery Issues", "Service problems and accountability"],
  settings: ["Settings", "Microsoft 365 and Excel connection"]
};

let cache = {
  orders: [],
  backorders: [],
  purchaseOrders: [],
  receiving: [],
  schedules: [],
  drivers: [],
  issues: []
};

function esc(value) {
  return String(value ?? "").replace(/[&<>'"]/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[ch]));
}

function statusClass(value) {
  const s = String(value || "").toUpperCase();
  if (["DELIVERED","COMPLETED","RELEASED","RECEIVED","READY FOR DELIVERY","READY","ACTIVE","CLOSED"].includes(s)) return "s-green";
  if (["SCHEDULED","LOADED","OUT FOR DELIVERY","PO ORDERED","PARTIALLY RECEIVED"].includes(s)) return "s-blue";
  if (["PROCESSING","EXPECTED","PO REQUIRED","BACKORDER","PO PARTIALLY RECEIVED"].includes(s)) return "s-orange";
  if (["FAILED","CANCELLED","HIGH","URGENT"].includes(s)) return "s-red";
  return "s-gray";
}

function badge(value) {
  return `<span class="status ${statusClass(value)}">${esc(value || "—")}</span>`;
}

function setAlert(id, message, type = "error") {
  const el = document.getElementById(id);
  if (!el) return;
  el.innerHTML = `<div class="${type}">${esc(message)}</div>`;
}

function clearAlert(id) {
  const el = document.getElementById(id);
  if (el) el.innerHTML = "";
}

function showToast(message) {
  const el = document.getElementById("toast");
  el.textContent = message;
  el.classList.add("show");
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(() => el.classList.remove("show"), 2600);
}

function formatDate(value) {
  if (!value) return "—";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return esc(value);
  return d.toLocaleDateString("en-CA", { year: "numeric", month: "short", day: "numeric" });
}

function toggleSidebar() {
  document.getElementById("sidebar").classList.toggle("open");
}

function navigate(page) {
  document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
  document.getElementById(`page-${page}`).classList.add("active");
  document.querySelectorAll(".nav-item[data-page]").forEach(n => n.classList.remove("active"));
  const nav = document.querySelector(`.nav-item[data-page="${page}"]`);
  if (nav) nav.classList.add("active");
  document.getElementById("pageTitle").textContent = PAGE_META[page][0];
  document.getElementById("pageSubtitle").textContent = PAGE_META[page][1];
  document.getElementById("sidebar").classList.remove("open");

  const loaders = {
    dashboard: loadDashboard,
    orders: loadOrders,
    backorders: loadBackorders,
    purchasing: loadPurchasing,
    receiving: loadReceiving,
    schedule: loadSchedule,
    drivers: loadDrivers,
    routes: loadRoutesPlaceholder,
    issues: loadIssues,
    settings: loadSettings
  };
  loaders[page]?.();
}

function bindNavigation() {
  document.querySelectorAll(".nav-item[data-page]").forEach(item => item.addEventListener("click", () => navigate(item.dataset.page)));
}

async function handleLogin() {
  try {
    await signIn();
    showSignedInState();
    await loadDashboard();
  } catch (error) {
    console.error(error);
    alert(`Microsoft sign-in failed:\n\n${error.message}`);
  }
}

async function handleLogout() {
  try {
    await signOut();
    document.getElementById("appShell").style.display = "none";
    document.getElementById("loginScreen").style.display = "grid";
  } catch (error) {
    alert(`Sign out failed:\n\n${error.message}`);
  }
}

function showSignedInState() {
  const account = getAccount();
  if (!account) return;
  document.getElementById("loginScreen").style.display = "none";
  document.getElementById("appShell").style.display = "flex";
  const name = account.name || account.username || "Microsoft User";
  document.getElementById("userName").textContent = name;
  document.getElementById("userMail").textContent = account.username || "";
  document.getElementById("avatar").textContent = name.split(/\s+/).slice(0,2).map(x => x[0]).join("").toUpperCase();
}

async function bootstrap() {
  bindNavigation();
  try {
    await initAuth();
    if (getAccount()) {
      showSignedInState();
      await loadDashboard();
    }
  } catch (error) {
    console.error(error);
  }
}

async function loadDashboard() {
  clearAlert("dashboardAlert");
  document.getElementById("routeSummary").innerHTML = `<div class="empty">Loading live Excel data...</div>`;
  document.getElementById("connectionSummary").innerHTML = `<div class="empty">Connecting...</div>`;

  if (!workbookConfigured()) {
    document.getElementById("connectionSummary").innerHTML = `<div class="panel-sub">Workbook IDs are not configured yet.</div><div style="margin-top:10px"><button class="btn btn-primary btn-small" onclick="navigate('settings')">Open Setup</button></div>`;
    ["mReady","mScheduled","mBackorders","mPORequired","mPOOrdered","mReceiving","mShipper"].forEach(id => document.getElementById(id).textContent = "—");
    document.getElementById("mWorkbook").textContent = "Setup";
    document.getElementById("routeSummary").innerHTML = `<div class="empty"><div class="empty-big">☁️</div><div>Finish the one-time workbook setup in Settings.</div></div>`;
    return;
  }

  try {
    const [orders, backorders, pos, receiving, schedules] = await Promise.all([
      getTableRows("OrdersTable"),
      getTableRows("BackordersTable"),
      getTableRows("PurchaseOrdersTable"),
      getTableRows("ReceivingTable"),
      getTableRows("DailyScheduleTable")
    ]);

    cache.orders = orders;
    cache.backorders = backorders;
    cache.purchaseOrders = pos;
    cache.receiving = receiving;
    cache.schedules = schedules;

    const ready = orders.filter(r => String(r.OrderStatus).toUpperCase() === "READY FOR DELIVERY").length;
    const scheduled = orders.filter(r => String(r.DeliveryStatus).toUpperCase() === "SCHEDULED").length;
    const openBackorders = backorders.filter(r => !["DELIVERED","CLOSED"].includes(String(r.BackorderStatus).toUpperCase())).length;
    const poRequired = pos.filter(r => String(r.POStatus).toUpperCase() === "PO REQUIRED").length;
    const poOrdered = pos.filter(r => String(r.POStatus).toUpperCase() === "PO ORDERED").length;
    const awaitingReceiving = receiving.filter(r => ["EXPECTED","PARTIALLY RECEIVED"].includes(String(r.ReceivingStatus).toUpperCase())).length;
    const awaitingShipper = receiving.filter(r => String(r.ReceivingStatus).toUpperCase() === "RECEIVED" && String(r.ReleasedToShipper).toLowerCase() !== "true").length;

    document.getElementById("mReady").textContent = ready;
    document.getElementById("mScheduled").textContent = scheduled;
    document.getElementById("mBackorders").textContent = openBackorders;
    document.getElementById("mPORequired").textContent = poRequired;
    document.getElementById("mPOOrdered").textContent = poOrdered;
    document.getElementById("mReceiving").textContent = awaitingReceiving;
    document.getElementById("mShipper").textContent = awaitingShipper;
    document.getElementById("mWorkbook").textContent = "Live";

    document.getElementById("routeSummary").innerHTML = schedules.length
      ? schedules.slice(0,8).map(r => `
          <div class="route">
            <div class="route-head"><div class="driver-line"><div class="driver-icon">🚚</div><div><div class="driver-name">${esc(r.DriverID || r.DriverName || "Driver")}</div><div class="driver-truck">${esc(r.TruckID || "Truck")} · ${esc(r.RouteID || "Route")}</div></div></div>${badge(r.RouteStatus)}</div>
            <div class="route-info"><span>📍 ${esc(r.TotalStops || 0)} stops</span><span>🛣 ${esc(r.EstimatedDistance || 0)}</span><span>⏱ ${esc(r.EstimatedTotalMinutes || 0)} min</span><span>📦 ${esc(r.LoadPercentage || 0)}%</span></div>
            <div class="progress"><div class="progress-bar" style="width:${Math.min(100, Number(r.LoadPercentage) || 0)}%"></div></div>
          </div>`).join("")
      : `<div class="empty">No daily routes found.</div>`;

    const info = await getWorkbookInfo();
    document.getElementById("connectionSummary").innerHTML = `
      <div><div class="panel-sub">Workbook</div><div style="font-weight:850;font-size:12px;margin-top:3px">${esc(info?.name || APP_CONFIG.workbookFileName)}</div></div>
      <div style="margin-top:12px">${badge("LIVE")}</div>
      <div class="panel-sub" style="margin-top:12px">Excel data is being read through Microsoft Graph.</div>`;
  } catch (error) {
    setAlert("dashboardAlert", error.message);
    document.getElementById("mWorkbook").textContent = "Error";
    document.getElementById("connectionSummary").innerHTML = `<div class="error">${esc(error.message)}</div><button class="btn btn-light btn-small" onclick="navigate('settings')">Check Setup</button>`;
  }
}

async function loadOrders() {
  clearAlert("ordersAlert");
  try {
    cache.orders = await getTableRows("OrdersTable");
    renderOrders();
  } catch (error) {
    setAlert("ordersAlert", error.message);
  }
}

function renderOrders() {
  const search = (document.getElementById("orderSearch")?.value || "").toLowerCase();
  const status = document.getElementById("orderFilter")?.value || "";
  const filtered = cache.orders.filter(r => {
    const text = [r.OrderNumber,r.CustomerName,r.DeliveryAddress,r.City,r.PostalCode].join(" ").toLowerCase();
    return text.includes(search) && (!status || String(r.OrderStatus).toUpperCase() === status);
  });
  document.getElementById("ordersCount").textContent = `${filtered.length} rows`;
  document.getElementById("ordersBody").innerHTML = filtered.length ? filtered.map(r => `
    <tr><td><strong>#${esc(r.OrderNumber)}</strong></td><td>${esc(r.CustomerName)}</td><td>${esc(r.DeliveryAddress)}</td><td>${formatDate(r.RequestedDeliveryDate)}</td><td>${badge(r.OrderStatus)}</td><td>${badge(r.DeliveryStatus)}</td></tr>`).join("") : `<tr><td colspan="6"><div class="empty">No matching orders.</div></td></tr>`;
}

async function loadBackorders() {
  clearAlert("backordersAlert");
  try {
    cache.backorders = await getTableRows("BackordersTable");
    document.getElementById("backordersBody").innerHTML = cache.backorders.length ? cache.backorders.map(r => `
      <tr><td>${esc(r.BackorderID)}</td><td>#${esc(r.OrderNumber)}</td><td>${esc(r.CustomerName)}</td><td>${esc(r.ItemDescription)}</td><td>${esc(r.QuantityBackordered)}</td><td>${esc(r.PONumber || "—")}</td><td>${badge(r.BackorderStatus)}</td></tr>`).join("") : `<tr><td colspan="7"><div class="empty">No backorders found.</div></td></tr>`;
  } catch (error) { setAlert("backordersAlert", error.message); }
}

async function loadPurchasing() {
  clearAlert("purchasingAlert");
  try {
    cache.purchaseOrders = await getTableRows("PurchaseOrdersTable");
    document.getElementById("poBody").innerHTML = cache.purchaseOrders.length ? cache.purchaseOrders.map(r => `
      <tr><td><strong>${esc(r.PONumber)}</strong></td><td>${esc(r.Supplier)}</td><td>${formatDate(r.PODate)}</td><td>${formatDate(r.ExpectedDate)}</td><td>${esc(r.TotalItems)}</td><td>${esc(r.ReceivedQuantity)}</td><td>${badge(r.POStatus)}</td></tr>`).join("") : `<tr><td colspan="7"><div class="empty">No purchase orders found.</div></td></tr>`;
  } catch (error) { setAlert("purchasingAlert", error.message); }
}

async function loadReceiving() {
  clearAlert("receivingAlert");
  try {
    cache.receiving = await getTableRows("ReceivingTable");
    document.getElementById("receivingBody").innerHTML = cache.receiving.length ? cache.receiving.map(r => `
      <tr><td><strong>${esc(r.PONumber)}</strong></td><td>${esc(r.ItemDescription)}</td><td>${esc(r.QuantityExpected)}</td><td>${esc(r.QuantityReceived)}</td><td>${esc(r.QuantityShort)}</td><td>${badge(r.ReceivingStatus)}</td><td>${esc(r.Receiver)}</td></tr>`).join("") : `<tr><td colspan="7"><div class="empty">No receiving rows found.</div></td></tr>`;
  } catch (error) { setAlert("receivingAlert", error.message); }
}

async function loadSchedule() {
  clearAlert("scheduleAlert");
  try {
    cache.schedules = await getTableRows("DailyScheduleTable");
    document.getElementById("scheduleBody").innerHTML = cache.schedules.length ? cache.schedules.map(r => `
      <tr><td>${formatDate(r.DeliveryDate)}</td><td>${esc(r.DriverID)}</td><td>${esc(r.TruckID)}</td><td>${esc(r.RouteID)}</td><td>${esc(r.TotalStops)}</td><td>${esc(r.EstimatedDistance)}</td><td>${esc(r.EstimatedTotalMinutes)}</td><td>${esc(r.LoadPercentage)}%</td><td>${badge(r.RouteStatus)}</td></tr>`).join("") : `<tr><td colspan="9"><div class="empty">No daily schedule rows found.</div></td></tr>`;
  } catch (error) { setAlert("scheduleAlert", error.message); }
}

async function loadDrivers() {
  clearAlert("driversAlert");
  try {
    const drivers = await getTableRows("DriversTable");
    document.getElementById("driversBody").innerHTML = drivers.length ? drivers.map(r => `
      <tr><td><strong>${esc(r.DriverName)}</strong></td><td>${esc(r.Phone)}</td><td>${esc(r.DefaultTruckID)}</td><td>${badge(String(r.Active).toUpperCase() === "TRUE" ? "ACTIVE" : "INACTIVE")}</td></tr>`).join("") : `<tr><td colspan="4"><div class="empty">No drivers found.</div></td></tr>`;
  } catch (error) { setAlert("driversAlert", error.message); }
}

async function loadRoutesPlaceholder() {}

async function loadIssues() {
  clearAlert("issuesAlert");
  try {
    cache.issues = await getTableRows("DeliveryIssuesTable");
    document.getElementById("issuesBody").innerHTML = cache.issues.length ? cache.issues.map(r => `
      <tr><td>${formatDate(r.Date)}</td><td>#${esc(r.OrderNumber)}</td><td>${esc(r.CustomerName)}</td><td>${esc(r.IssueType)}</td><td>${esc(r.ResponsibleArea)}</td><td>${badge(r.Severity)}</td><td>${esc(r.Resolved)}</td></tr>`).join("") : `<tr><td colspan="7"><div class="empty">No issues found.</div></td></tr>`;
  } catch (error) { setAlert("issuesAlert", error.message); }
}

function loadSettings() {
  document.getElementById("settingsDriveId").value = APP_CONFIG.workbookDriveId;
  document.getElementById("settingsItemId").value = APP_CONFIG.workbookItemId;
}

async function findWorkbook() {
  clearAlert("settingsAlert");
  document.getElementById("setupResult").innerHTML = `<div class="panel-sub">Searching your OneDrive root for ${esc(APP_CONFIG.workbookFileName)}...</div>`;
  try {
    const info = await resolveOwnerWorkbook();
    document.getElementById("setupResult").innerHTML = `
      <div class="success">Workbook found: ${esc(info.name)}</div>
      <div class="panel-sub">Copy these IDs into config.js in GitHub. They are file identifiers, not passwords.</div>
      <div class="copy-row"><input value="${esc(info.driveId)}" readonly><button class="btn btn-light btn-small" onclick="copyText('${esc(info.driveId)}')">Copy Drive ID</button></div>
      <div class="copy-row"><input value="${esc(info.itemId)}" readonly><button class="btn btn-light btn-small" onclick="copyText('${esc(info.itemId)}')">Copy Item ID</button></div>`;
  } catch (error) {
    setAlert("settingsAlert", error.message);
    document.getElementById("setupResult").innerHTML = `<div class="panel-sub">Make sure the file is named exactly <b>${esc(APP_CONFIG.workbookFileName)}</b> and is in OneDrive → My files.</div>`;
  }
}

function saveWorkbookIds() {
  const driveId = document.getElementById("settingsDriveId").value.trim();
  const itemId = document.getElementById("settingsItemId").value.trim();
  if (!driveId || !itemId) {
    setAlert("settingsAlert", "Both Drive ID and Item ID are required.");
    return;
  }
  APP_CONFIG.workbookDriveId = driveId;
  APP_CONFIG.workbookItemId = itemId;
  try { localStorage.setItem("mouldingWorkbookDriveId", driveId); localStorage.setItem("mouldingWorkbookItemId", itemId); } catch {}
  showToast("Workbook IDs saved in this browser.");
  testConnection();
}

function loadSavedIds() {
  try {
    const driveId = localStorage.getItem("mouldingWorkbookDriveId");
    const itemId = localStorage.getItem("mouldingWorkbookItemId");
    if (!APP_CONFIG.workbookDriveId && driveId) APP_CONFIG.workbookDriveId = driveId;
    if (!APP_CONFIG.workbookItemId && itemId) APP_CONFIG.workbookItemId = itemId;
  } catch {}
}

async function testConnection() {
  clearAlert("settingsAlert");
  if (!workbookConfigured()) {
    setAlert("settingsAlert", "Configure the workbook IDs first.");
    return;
  }
  document.getElementById("setupResult").innerHTML = `<div class="panel-sub">Testing Microsoft Graph → Excel...</div>`;
  try {
    const result = await testWorkbookConnection();
    document.getElementById("setupResult").innerHTML = `<div class="success">Connection successful: ${esc(result.info.name)}</div><div class="panel-sub">Found ${result.tables.length} Excel tables.</div><div class="code-box" style="margin-top:10px">${esc(result.tables.join("\n"))}</div>`;
    showToast("Excel connection successful");
  } catch (error) {
    setAlert("settingsAlert", error.message);
  }
}

async function refreshCurrentPage() {
  const active = document.querySelector(".page.active")?.id?.replace("page-", "dashboard") || "dashboard";
  const page = document.querySelector(".page.active")?.id?.replace("page-", "") || "dashboard";
  await (page === "dashboard" ? loadDashboard() : page === "orders" ? loadOrders() : page === "backorders" ? loadBackorders() : page === "purchasing" ? loadPurchasing() : page === "receiving" ? loadReceiving() : page === "schedule" ? loadSchedule() : page === "drivers" ? loadDrivers() : page === "issues" ? loadIssues() : loadSettings());
}

function openSimpleOrderModal() {
  document.getElementById("orderModal").classList.add("open");
}

function closeModal(id) {
  document.getElementById(id).classList.remove("open");
}

async function createSimpleOrder() {
  const orderNumber = document.getElementById("newOrderNumber").value.trim();
  const customer = document.getElementById("newCustomerName").value.trim();
  const address = document.getElementById("newDeliveryAddress").value.trim();
  const requestedDate = document.getElementById("newRequestedDate").value;
  const priority = document.getElementById("newPriority").value;
  const instructions = document.getElementById("newInstructions").value.trim();

  if (!orderNumber || !customer || !address) {
    showToast("Order number, customer and address are required.");
    return;
  }

  const today = new Date().toISOString();
  const row = {
    OrderID: `O-${Date.now()}`,
    OrderNumber: orderNumber,
    CustomerID: "",
    CustomerName: customer,
    AddressID: "",
    DeliveryAddress: address,
    City: "",
    PostalCode: "",
    Phone: "",
    OrderDate: today.substring(0,10),
    RequestedDeliveryDate: requestedDate,
    Priority: priority,
    OrderStatus: "PROCESSING",
    DeliveryStatus: "NOT SCHEDULED",
    BackorderPresent: "FALSE",
    DeliveryOption: "DELIVERY",
    ReadyForDelivery: "FALSE",
    SpecialInstructions: instructions,
    CreatedBy: getAccount()?.username || "",
    CreatedDate: today,
    UpdatedDate: today
  };

  try {
    await addTableRow("OrdersTable", row);
    closeModal("orderModal");
    showToast(`Order #${orderNumber} added to Excel.`);
    await loadOrders();
  } catch (error) {
    alert(`Could not create the order:\n\n${error.message}`);
  }
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
    showToast("Copied");
  } catch {
    showToast("Copy failed — select the value manually.");
  }
}

document.addEventListener("DOMContentLoaded", () => {
  loadSavedIds();
  bootstrap();
});
