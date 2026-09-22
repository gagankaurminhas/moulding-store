(() => {
  const state = {
    data: {drivers:[], trucks:[], orders:[], schedules:[], routeStops:[], deliveryStatus:[], loading:[], issues:[]},
    view: "dispatch",
    selectedDriver: "",
    selectedDate: dateKey(new Date()),
    busy: false,
    error: ""
  };

  const root = document.getElementById("app");

  function dateKey(d) {
    const x = new Date(d);
    return `${x.getFullYear()}-${String(x.getMonth()+1).padStart(2,"0")}-${String(x.getDate()).padStart(2,"0")}`;
  }
  function esc(v) {
    return String(v ?? "").replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }
  function id(prefix) { return `${prefix}-${Date.now()}-${Math.floor(Math.random()*1000)}`; }
  function num(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; }
  function active(v) { return String(v).toLowerCase() !== "false" && String(v).toLowerCase() !== "no"; }
  function driver(id) { return state.data.drivers.find(x => String(x.DriverID) === String(id)); }
  function truck(id) { return state.data.trucks.find(x => String(x.TruckID) === String(id)); }
  function ordersFor(driverId) {
    return state.data.orders.filter(o =>
      String(o.AssignedDriverID) === String(driverId) &&
      String(o.RequestedDeliveryDate || "").slice(0,10) === state.selectedDate &&
      !["CANCELLED","DELIVERED"].includes(String(o.OrderStatus).toUpperCase())
    );
  }
  function stopsFor(driverId) {
    return state.data.routeStops
      .filter(s => String(s.DriverID) === String(driverId) && String(s.DeliveryDate).slice(0,10) === state.selectedDate)
      .sort((a,b) => num(a.StopNumber)-num(b.StopNumber));
  }
  function routeStatus(driverId) {
    const sched = state.data.schedules.find(s =>
      String(s.DriverID) === String(driverId) &&
      String(s.DeliveryDate).slice(0,10) === state.selectedDate
    );
    return sched?.RouteStatus || "UNASSIGNED";
  }
  function latestStatus(orderId) {
    const rows = state.data.deliveryStatus.filter(x => String(x.OrderID) === String(orderId));
    return rows.length ? rows[rows.length-1].Status : "";
  }
  function capacity(driverId) {
    const d = driver(driverId);
    const t = truck(d?.DefaultTruckID);
    const assigned = ordersFor(driverId).reduce((s,o) => s + num(o.EstimatedLoadUnits), 0);
    const cap = num(t?.Capacity);
    return {used:assigned, cap, pct:cap ? Math.min(100, assigned/cap*100) : 0};
  }

  async function refresh(silent=false) {
    if (state.busy) return;
    state.busy = true;
    if (!silent) render();
    try {
      state.data = await ExcelAPI.load();
      state.error = "";
    } catch(e) {
      state.error = e.message || "Unable to load Excel.";
    } finally {
      state.busy = false;
      render();
    }
  }

  function shell() {
    root.innerHTML = `
      <div class="app-shell">
        <aside class="sidebar">
          <div class="brand"><span class="brand-mark">M</span><div><strong>MOULDING</strong><small>LIVE DISPATCH</small></div></div>
          <nav>
            <button class="${state.view==='dispatch'?'active':''}" data-view="dispatch">Dispatch</button>
            <button class="${state.view==='routes'?'active':''}" data-view="routes">Routes</button>
            <button class="${state.view==='orders'?'active':''}" data-view="orders">Deliveries</button>
            <button class="${state.view==='drivers'?'active':''}" data-view="drivers">Drivers</button>
            <button class="${state.view==='trucks'?'active':''}" data-view="trucks">Trucks</button>
            <button class="${state.view==='settings'?'active':''}" data-view="settings">Settings</button>
          </nav>
          <div class="side-bottom">
            <div class="live-dot"><i></i> Live data</div>
            <small>Auto refresh: ${APP_CONFIG.refreshSeconds}s</small>
          </div>
        </aside>
        <main class="main">
          <header class="topbar">
            <div>
              <div class="eyebrow">OPERATIONS</div>
              <h1>${state.view==='dispatch'?'Live Dispatch':state.view==='routes'?'Routes':state.view==='orders'?'Deliveries':state.view==='drivers'?'Drivers':state.view==='trucks'?'Trucks':'Settings'}</h1>
            </div>
            <div class="top-actions">
              <input id="datePick" type="date" value="${esc(state.selectedDate)}" aria-label="Dispatch date">
              <button id="refreshBtn" class="btn ghost">${state.busy?'Refreshing…':'Refresh'}</button>
              <button id="newOrderBtn" class="btn primary">+ Delivery</button>
            </div>
          </header>
          ${state.error ? `<div class="error">${esc(state.error)} <button id="findWorkbookBtn">Find workbook</button></div>` : ""}
          <section id="content"></section>
        </main>
      </div>
    `;

    root.querySelectorAll("[data-view]").forEach(b => b.addEventListener("click", () => {state.view=b.dataset.view; render();}));
    document.getElementById("refreshBtn")?.addEventListener("click", () => refresh());
    document.getElementById("datePick")?.addEventListener("change", e => {state.selectedDate=e.target.value; render();});
    document.getElementById("newOrderBtn")?.addEventListener("click", newOrder);
    document.getElementById("findWorkbookBtn")?.addEventListener("click", async () => {
      try { await ExcelAPI.findAndSaveWorkbook(); await refresh(); } catch(e) { state.error=e.message; render(); }
    });
  }

  function render() {
    shell();
    const c = document.getElementById("content");
    if (!c) return;
    if (state.view==="dispatch") c.innerHTML = dispatchView();
    if (state.view==="routes") c.innerHTML = routesView();
    if (state.view==="orders") c.innerHTML = ordersView();
    if (state.view==="drivers") c.innerHTML = driversView();
    if (state.view==="trucks") c.innerHTML = trucksView();
    if (state.view==="settings") c.innerHTML = settingsView();
    bindContent();
  }

  function dispatchView() {
    const ds = state.data.drivers.filter(d => active(d.Active));
    const unassigned = state.data.orders.filter(o =>
      String(o.RequestedDeliveryDate||"").slice(0,10)===state.selectedDate &&
      !o.AssignedDriverID &&
      !["DELIVERED","CANCELLED"].includes(String(o.OrderStatus).toUpperCase())
    );
    const total = ds.length;
    const out = ds.filter(d => ["OUT FOR DELIVERY","OUT","IN PROGRESS","ON ROUTE"].includes(routeStatus(d.DriverID).toUpperCase())).length;
    const delivered = state.data.orders.filter(o => String(o.RequestedDeliveryDate||"").slice(0,10)===state.selectedDate && String(o.OrderStatus).toUpperCase()==="DELIVERED").length;
    return `
      <div class="metrics">
        <div class="metric"><span>Drivers</span><b>${total}</b></div>
        <div class="metric"><span>On route</span><b>${out}</b></div>
        <div class="metric"><span>Unassigned</span><b>${unassigned.length}</b></div>
        <div class="metric"><span>Delivered</span><b>${delivered}</b></div>
      </div>
      <div class="dispatch-layout">
        <section class="panel unassigned">
          <div class="panel-head"><div><strong>Unassigned deliveries</strong><small>Drag with your mouse or use Assign</small></div><span>${unassigned.length}</span></div>
          <div class="order-stack" id="unassignedList">
            ${unassigned.length ? unassigned.map(orderCard).join("") : `<div class="empty">Nothing waiting for assignment.</div>`}
          </div>
        </section>
        <section class="driver-grid">
          ${ds.map(driverCard).join("")}
        </section>
      </div>
    `;
  }

  function orderCard(o) {
    const status = latestStatus(o.OrderID) || o.DeliveryStatus || o.OrderStatus || "NEW";
    return `<article class="order-card" draggable="true" data-order="${esc(o.OrderID)}">
      <div class="order-top"><b>${esc(o.OrderNumber||o.OrderID)}</b><span class="pill">${esc(status)}</span></div>
      <strong>${esc(o.CustomerName||"Customer")}</strong>
      <span>${esc(o.DeliveryAddress||"No address")}</span>
      <div class="order-meta"><span>${num(o.EstimatedLoadUnits)} load</span><span>${esc(o.Priority||"Normal")}</span></div>
      <button class="assign-btn" data-assign-order="${esc(o.OrderID)}">Assign</button>
    </article>`;
  }

  function driverCard(d) {
    const t = truck(d.DefaultTruckID);
    const cap = capacity(d.DriverID);
    const stops = stopsFor(d.DriverID);
    const os = ordersFor(d.DriverID);
    const status = routeStatus(d.DriverID);
    return `<section class="driver-card" data-driver="${esc(d.DriverID)}">
      <div class="driver-head">
        <div class="avatar">${esc((d.DriverName||"?").slice(0,1).toUpperCase())}</div>
        <div><strong>${esc(d.DriverName)}</strong><small>${esc(t?.TruckName||d.DefaultTruckID||"No truck")}</small></div>
        <span class="route-pill ${String(status).toLowerCase().replaceAll(" ","-")}">${esc(status)}</span>
      </div>
      <div class="capacity"><div><span>Capacity</span><b>${cap.used} / ${cap.cap || "—"}</b></div><div class="bar"><i style="width:${cap.pct}%"></i></div></div>
      <div class="mini-stats"><span>${os.length} deliveries</span><span>${stops.length} stops</span></div>
      <div class="stop-list" data-drop-driver="${esc(d.DriverID)}">
        ${stops.length ? stops.map((s,i)=>stopCard(s,i)).join("") : os.map((o,i)=>fallbackStop(o,i)).join("") || `<div class="drop-hint">Drop a delivery here</div>`}
      </div>
      <button class="route-open" data-route="${esc(d.DriverID)}">Open route →</button>
    </section>`;
  }

  function stopCard(s,i) {
    return `<div class="stop-row" draggable="true" data-stop="${esc(s.StopID)}">
      <b>${num(s.StopNumber)||i+1}</b><div><strong>${esc(s.CustomerName||s.OrderNumber)}</strong><small>${esc(s.Address||"")}</small></div><span class="stop-status">${esc(s.StopStatus||"PENDING")}</span>
    </div>`;
  }
  function fallbackStop(o,i) {
    return `<div class="stop-row"><b>${i+1}</b><div><strong>${esc(o.CustomerName||o.OrderNumber)}</strong><small>${esc(o.DeliveryAddress||"")}</small></div><span class="stop-status">${esc(latestStatus(o.OrderID)||"PENDING")}</span></div>`;
  }

  function routesView() {
    const ds = state.data.drivers.filter(d=>active(d.Active));
    return `<div class="section-title"><div><h2>Today's routes</h2><p>Live route status and stop sequence.</p></div></div>
      <div class="route-table">
        ${ds.map(d=>{
          const cap=capacity(d.DriverID), s=stopsFor(d.DriverID);
          return `<div class="route-line"><div><b>${esc(d.DriverName)}</b><small>${esc(truck(d.DefaultTruckID)?.TruckName||"")}</small></div><strong>${s.length} stops</strong><span>${cap.used}/${cap.cap||"—"} load</span><em>${esc(routeStatus(d.DriverID))}</em><button data-route="${esc(d.DriverID)}">View</button></div>`
        }).join("")}
      </div>`;
  }

  function ordersView() {
    const rows = state.data.orders.filter(o=>String(o.RequestedDeliveryDate||"").slice(0,10)===state.selectedDate);
    return `<div class="section-title"><div><h2>Deliveries</h2><p>${rows.length} deliveries scheduled for ${esc(state.selectedDate)}.</p></div></div>
      <div class="table-wrap"><table><thead><tr><th>Order</th><th>Customer</th><th>Address</th><th>Driver</th><th>Status</th><th>Load</th></tr></thead><tbody>
      ${rows.map(o=>`<tr><td><b>${esc(o.OrderNumber||o.OrderID)}</b></td><td>${esc(o.CustomerName)}</td><td>${esc(o.DeliveryAddress)}</td><td>${esc(driver(o.AssignedDriverID)?.DriverName||"Unassigned")}</td><td><span class="pill">${esc(latestStatus(o.OrderID)||o.DeliveryStatus||o.OrderStatus||"NEW")}</span></td><td>${num(o.EstimatedLoadUnits)}</td></tr>`).join("")}
      </tbody></table></div>`;
  }

  function driversView() {
    return `<div class="cards">${state.data.drivers.map(d=>`<div class="info-card"><div class="avatar">${esc((d.DriverName||"?").slice(0,1))}</div><h3>${esc(d.DriverName)}</h3><p>${esc(d.Phone||"")}</p><span>${esc(truck(d.DefaultTruckID)?.TruckName||"No truck")}</span></div>`).join("")}</div>`;
  }

  function trucksView() {
    return `<div class="cards">${state.data.trucks.map(t=>{
      const d=state.data.drivers.find(x=>String(x.DefaultTruckID)===String(t.TruckID));
      const used=state.data.orders.filter(o=>String(o.AssignedTruckID)===String(t.TruckID)&&String(o.RequestedDeliveryDate||"").slice(0,10)===state.selectedDate).reduce((s,o)=>s+num(o.EstimatedLoadUnits),0);
      const pct=num(t.Capacity)?Math.min(100,used/num(t.Capacity)*100):0;
      return `<div class="info-card truck-info"><div class="truck-icon">▣</div><h3>${esc(t.TruckName||t.TruckID)}</h3><p>${esc(d?.DriverName||"Unassigned")}</p><strong>${used} / ${num(t.Capacity)||"—"}</strong><div class="bar"><i style="width:${pct}%"></i></div></div>`;
    }).join("")}</div>`;
  }

  function settingsView() {
    return `<section class="panel settings"><h2>Connection</h2><p>Excel workbook: <b>${esc(APP_CONFIG.workbookFileName)}</b></p>
      <button id="findBtn" class="btn primary">Find workbook in OneDrive</button>
      <button id="clearBtn" class="btn ghost">Clear saved workbook IDs</button>
      <p class="muted">The existing auth.js handles Microsoft sign-in. This app reads and writes the Excel tables through Microsoft Graph.</p></section>`;
  }

  function bindContent() {
    root.querySelectorAll("[data-route]").forEach(b=>b.addEventListener("click",()=>{state.view="routes";state.selectedDriver=b.dataset.route;render();}));
    root.querySelectorAll("[data-assign-order]").forEach(b=>b.addEventListener("click",()=>assignDialog(b.dataset.assignOrder)));
    root.querySelectorAll("[data-drop-driver]").forEach(el=>{
      el.addEventListener("dragover",e=>e.preventDefault());
      el.addEventListener("drop",async e=>{
        e.preventDefault();
        const oid=e.dataTransfer.getData("text/order");
        const did=el.dataset.dropDriver;
        if(oid) await assignOrder(oid,did);
      });
    });
    root.querySelectorAll("[draggable='true'][data-order]").forEach(el=>el.addEventListener("dragstart",e=>e.dataTransfer.setData("text/order",el.dataset.order)));
    document.getElementById("findBtn")?.addEventListener("click",async()=>{try{await ExcelAPI.findAndSaveWorkbook();await refresh();}catch(e){state.error=e.message;render();}});
    document.getElementById("clearBtn")?.addEventListener("click",()=>{localStorage.removeItem("mouldingWorkbookDriveId");localStorage.removeItem("mouldingWorkbookItemId");alert("Saved workbook IDs cleared.");});
  }

  async function assignDialog(orderId) {
    const ds=state.data.drivers.filter(d=>active(d.Active));
    const name=prompt("Enter driver name or DriverID:\n\n"+ds.map(d=>`${d.DriverID} — ${d.DriverName}`).join("\n"));
    if (!name) return;
    const d=ds.find(x=>String(x.DriverID).toLowerCase()===name.toLowerCase()||String(x.DriverName).toLowerCase()===name.toLowerCase());
    if (!d) return alert("Driver not found.");
    await assignOrder(orderId,d.DriverID);
  }

  async function assignOrder(orderId, driverId) {
    try {
      const d=driver(driverId);
      const t=truck(d?.DefaultTruckID);
      const o=state.data.orders.find(x=>String(x.OrderID)===String(orderId));
      const current=ordersFor(driverId).filter(x=>String(x.OrderID)!==String(orderId)).reduce((s,x)=>s+num(x.EstimatedLoadUnits),0);
      if(t?.Capacity && current+num(o?.EstimatedLoadUnits)>num(t.Capacity)) {
        alert(`Capacity exceeded. ${t.TruckName} has ${Math.max(0,num(t.Capacity)-current)} load units remaining.`);
        return;
      }
      await ExcelAPI.updateOrderById(orderId,{
        AssignedDriverID:driverId,
        AssignedTruckID:d?.DefaultTruckID||"",
        DeliveryStatus:"SCHEDULED",
        ShippingStatus:"SCHEDULED",
        UpdatedDate:new Date().toISOString()
      });
      await ExcelAPI.setStatus({
        OrderID:orderId,
        OrderNumber:o?.OrderNumber||"",
        DeliveryDate:state.selectedDate,
        DriverID:driverId,
        RouteID:"",
        StopID:"",
        Status:"SCHEDULED",
        Notes:"Assigned from live dispatch"
      });
      await refresh();
    } catch(e) { alert(e.message||"Assignment failed."); }
  }

  async function newOrder() {
    const customer=prompt("Customer name:");
    if(!customer) return;
    const address=prompt("Delivery address:");
    if(!address) return;
    const load=prompt("Estimated load units:","1");
    const units=Math.max(0,num(load));
    const order={
      OrderID:id("O"),
      OrderNumber:`${Math.floor(10000+Math.random()*89999)}`,
      CustomerName:customer,
      DeliveryAddress:address,
      RequestedDeliveryDate:state.selectedDate,
      OrderDate:state.selectedDate,
      Priority:"NORMAL",
      OrderStatus:"READY",
      DeliveryStatus:"READY FOR SCHEDULING",
      ReadyForDelivery:"TRUE",
      EstimatedLoadUnits:units,
      ShippingStatus:"READY",
      CreatedDate:new Date().toISOString(),
      UpdatedDate:new Date().toISOString()
    };
    try { await ExcelAPI.addOrder(order); await refresh(); }
    catch(e) { alert(e.message||"Could not create delivery."); }
  }

  async function boot() {
    render();
    try {
      if(typeof initAuth==="function") await initAuth();
      await refresh();
    } catch(e) { state.error=e.message||"Authentication failed."; render(); }
    setInterval(()=>refresh(true), Math.max(10,APP_CONFIG.refreshSeconds)*1000);
  }

  boot();
})();
