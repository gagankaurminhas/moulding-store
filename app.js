(() => {
  "use strict";

  const state = {
    data: {orders: [], deliveryStatus: [], drivers: [], trucks: [], schedules: [], routeStops: []},
    page: "today",
    dateFrom: "",
    dateTo: "",
    query: "",
    selected: null,
    loading: false,
    error: "",
    lastUpdated: null
  };

  const root = document.getElementById("app");

  const todayKey = () => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
  };

  const esc = value => String(value ?? "").replace(/[&<>"']/g, ch => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  }[ch]));

  const upper = value => String(value ?? "").trim().toUpperCase();
  const n = value => Number.isFinite(Number(value)) ? Number(value) : 0;

  function dateOnly(value) {
    const s = String(value ?? "").slice(0, 10);
    return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : "";
  }

  function formatDate(value) {
    const d = dateOnly(value);
    if (!d) return "—";
    const [y,m,day] = d.split("-").map(Number);
    return new Intl.DateTimeFormat("en-CA", {
      year:"numeric", month:"short", day:"numeric"
    }).format(new Date(y, m-1, day));
  }

  function formatTime(value) {
    if (!value) return "—";
    const s = String(value);
    const d = new Date(s);
    if (!Number.isNaN(d.getTime()) && s.includes("T")) {
      return new Intl.DateTimeFormat("en-CA", {hour:"numeric", minute:"2-digit"}).format(d);
    }
    return s.length >= 5 ? s.slice(0,5) : s;
  }

  function driverById(id) {
    return state.data.drivers.find(x => String(x.DriverID) === String(id));
  }

  function truckById(id) {
    return state.data.trucks.find(x => String(x.TruckID) === String(id));
  }

  function latestStatus(orderId) {
    const rows = state.data.deliveryStatus
      .filter(x => String(x.OrderID) === String(orderId))
      .sort((a,b) => {
        const aa = String(a.CompletionTime || a.ArrivalTime || a.StatusID || "");
        const bb = String(b.CompletionTime || b.ArrivalTime || b.StatusID || "");
        return aa.localeCompare(bb);
      });
    return rows.length ? rows[rows.length - 1] : null;
  }

  function deliveryDate(order) {
    return dateOnly(
      order.RequestedDeliveryDate ||
      order.DeliveryDate ||
      latestStatus(order.OrderID)?.DeliveryDate
    );
  }

  function driverFor(order) {
    const status = latestStatus(order.OrderID);
    const id = order.AssignedDriverID || status?.DriverID;
    return driverById(id);
  }

  function truckFor(order) {
    const status = latestStatus(order.OrderID);
    const id = order.AssignedTruckID || status?.TruckID;
    return truckById(id);
  }

  function statusFor(order) {
    return latestStatus(order.OrderID)?.Status ||
      order.DeliveryStatus ||
      order.ShippingStatus ||
      order.OrderStatus ||
      "SCHEDULED";
  }

  function isCompleted(order) {
    return ["DELIVERED","COMPLETED"].includes(upper(statusFor(order))) ||
      upper(order.OrderStatus) === "DELIVERED";
  }

  function statusClass(status) {
    const s = upper(status);
    if (s.includes("DELIVER")) return "delivered";
    if (s.includes("OUT")) return "out";
    if (s.includes("ARRIV")) return "arrived";
    if (s.includes("FAIL") || s.includes("CANCEL") || s.includes("PROBLEM")) return "problem";
    return "scheduled";
  }

  function matchesQuery(order, query) {
    if (!query.trim()) return true;
    const q = query.trim().toLowerCase();
    const d = driverFor(order)?.DriverName || "";
    const t = truckFor(order)?.TruckName || "";
    return [
      order.OrderID, order.OrderNumber, order.CustomerName,
      order.Company, order.DeliveryAddress, order.City,
      order.PostalCode, order.Phone, d, t
    ].some(v => String(v ?? "").toLowerCase().includes(q));
  }

  function currentRows() {
    let rows = [...state.data.orders];

    if (state.page === "today") {
      const today = todayKey();
      rows = rows.filter(o => deliveryDate(o) === today);
    } else {
      if (state.dateFrom) rows = rows.filter(o => deliveryDate(o) >= state.dateFrom);
      if (state.dateTo) rows = rows.filter(o => deliveryDate(o) <= state.dateTo);
    }

    rows = rows.filter(o => matchesQuery(o, state.query));
    rows.sort((a,b) => {
      const da = deliveryDate(a), db = deliveryDate(b);
      if (da !== db) return db.localeCompare(da);
      return String(a.OrderNumber || a.OrderID).localeCompare(String(b.OrderNumber || b.OrderID));
    });
    return rows;
  }

  function summary(rows) {
    return {
      total: rows.length,
      delivered: rows.filter(isCompleted).length,
      out: rows.filter(o => upper(statusFor(o)).includes("OUT")).length,
      open: rows.filter(o => !isCompleted(o)).length
    };
  }

  async function refresh() {
    if (state.loading) return;
    state.loading = true;
    render();
    try {
      state.data = await ExcelAPI.load();
      state.lastUpdated = new Date();
      state.error = "";
    } catch (error) {
      state.error = error?.message || "Unable to read the Excel workbook.";
    } finally {
      state.loading = false;
      render();
    }
  }

  function renderShell() {
    root.innerHTML = `
      <div class="app">
        <header class="lux-header">
          <div class="brand">
            <div class="monogram">M</div>
            <div>
              <div class="brand-name">MOULDING STORE</div>
              <div class="brand-sub">DELIVERY INTELLIGENCE</div>
            </div>
          </div>

          <div class="header-right">
            <div class="live">
              <span class="live-dot"></span>
              LIVE
            </div>
            <button class="icon-btn" id="refreshBtn" title="Refresh data" aria-label="Refresh data">↻</button>
            <button class="user-chip" id="settingsBtn" aria-label="Open settings">MS</button>
          </div>
        </header>

        <nav class="lux-nav">
          <button class="${state.page === "today" ? "active" : ""}" data-page="today">Today</button>
          <button class="${state.page === "history" ? "active" : ""}" data-page="history">Delivery History</button>
        </nav>

        <main>
          ${state.error ? `
            <div class="error-banner">
              <div><strong>Workbook connection</strong><span>${esc(state.error)}</span></div>
              <button id="findWorkbook">Find Workbook</button>
            </div>` : ""}

          ${state.page === "today" ? todayView() : historyView()}
        </main>

        <footer>
          <span>READ ONLY • SOURCE: MOULDINGSTORE DATABASE</span>
          <span>${state.lastUpdated ? `Updated ${formatTime(state.lastUpdated.toISOString())}` : "Connecting…"}</span>
        </footer>
      </div>
      ${state.selected ? detailModal(state.selected) : ""}
      ${state.page === "settings" ? settingsModal() : ""}
    `;

    bindShell();
  }

  function todayView() {
    const rows = currentRows();
    const s = summary(rows);

    return `
      <section class="hero">
        <div>
          <div class="eyebrow">OPERATIONS / TODAY</div>
          <h1>Today's Deliveries</h1>
          <p>${formatDate(todayKey())} <span class="gold-dot">•</span> Live information from Excel</p>
        </div>
        <div class="hero-date">${new Date().toLocaleDateString("en-CA", {weekday:"long"})}</div>
      </section>

      <section class="stats">
        <div class="stat"><span>Scheduled</span><strong>${s.total}</strong></div>
        <div class="stat"><span>Out for delivery</span><strong>${s.out}</strong></div>
        <div class="stat"><span>Delivered</span><strong>${s.delivered}</strong></div>
        <div class="stat"><span>Open</span><strong>${s.open}</strong></div>
      </section>

      <section class="toolbar">
        <div class="search">
          <span>⌕</span>
          <input id="todaySearch" value="${esc(state.query)}" placeholder="Search order, customer, address, driver…" autocomplete="off">
        </div>
        <div class="read-only">VIEW ONLY</div>
      </section>

      ${deliveryList(rows)}
    `;
  }

  function historyView() {
    const rows = currentRows();
    return `
      <section class="hero compact">
        <div>
          <div class="eyebrow">ARCHIVE</div>
          <h1>Delivery History</h1>
          <p>Search previous deliveries directly from the workbook.</p>
        </div>
      </section>

      <section class="history-tools">
        <div class="search wide">
          <span>⌕</span>
          <input id="historySearch" value="${esc(state.query)}" placeholder="Search order, customer, address, driver…" autocomplete="off">
        </div>
        <label>From<input id="dateFrom" type="date" value="${esc(state.dateFrom)}"></label>
        <label>To<input id="dateTo" type="date" value="${esc(state.dateTo)}"></label>
        <button class="gold-btn" id="clearFilters">Clear</button>
      </section>

      <div class="history-count">${rows.length} matching delivery${rows.length === 1 ? "" : "ies"}</div>
      ${deliveryList(rows)}
    `;
  }

  function deliveryList(rows) {
    if (state.loading) {
      return `<div class="loading-card"><div class="spinner"></div><span>Reading live delivery data…</span></div>`;
    }

    if (!rows.length) {
      return `<div class="empty-card"><div class="empty-mark">—</div><h2>No deliveries found</h2><p>This screen only displays records that exist in the Excel database.</p></div>`;
    }

    return `
      <section class="delivery-list">
        ${rows.map((o, i) => {
          const d = driverFor(o);
          const t = truckFor(o);
          const status = statusFor(o);
          return `
            <button class="delivery-row" data-order="${esc(o.OrderID)}">
              <div class="row-index">${String(i+1).padStart(2,"0")}</div>
              <div class="row-main">
                <div class="row-title">
                  <strong>#${esc(o.OrderNumber || o.OrderID)}</strong>
                  <span class="status ${statusClass(status)}">${esc(status)}</span>
                </div>
                <div class="customer">${esc(o.CustomerName || o.Company || "Customer")}</div>
                <div class="address">${esc(o.DeliveryAddress || "Address not provided")}</div>
              </div>
              <div class="row-meta">
                <div><span>Date</span><strong>${formatDate(deliveryDate(o))}</strong></div>
                <div><span>Driver</span><strong>${esc(d?.DriverName || "—")}</strong></div>
                <div><span>Truck</span><strong>${esc(t?.TruckName || "—")}</strong></div>
              </div>
              <div class="chevron">›</div>
            </button>`;
        }).join("")}
      </section>
    `;
  }

  function detailModal(order) {
    const status = statusFor(order);
    const d = driverFor(order);
    const t = truckFor(order);
    const history = state.data.deliveryStatus
      .filter(x => String(x.OrderID) === String(order.OrderID))
      .sort((a,b) => String(a.ArrivalTime || a.CompletionTime || "").localeCompare(String(b.ArrivalTime || b.CompletionTime || "")));

    return `
      <div class="modal-backdrop" id="modalBackdrop">
        <section class="detail-modal" role="dialog" aria-modal="true">
          <button class="close" id="closeModal" aria-label="Close">×</button>

          <div class="detail-kicker">DELIVERY RECORD</div>
          <div class="detail-title">
            <div>
              <h2>#${esc(order.OrderNumber || order.OrderID)}</h2>
              <span class="status ${statusClass(status)}">${esc(status)}</span>
            </div>
            <div class="detail-date">${formatDate(deliveryDate(order))}</div>
          </div>

          <div class="customer-block">
            <div class="monogram small">${esc((order.CustomerName || "C").slice(0,1).toUpperCase())}</div>
            <div>
              <h3>${esc(order.CustomerName || order.Company || "Customer")}</h3>
              <p>${esc(order.DeliveryAddress || "Address not provided")}</p>
              <p>${esc([order.City, order.Province, order.PostalCode].filter(Boolean).join(", "))}</p>
            </div>
          </div>

          <div class="detail-grid">
            <div><span>Driver</span><strong>${esc(d?.DriverName || "—")}</strong></div>
            <div><span>Truck</span><strong>${esc(t?.TruckName || "—")}</strong></div>
            <div><span>Priority</span><strong>${esc(order.Priority || "—")}</strong></div>
            <div><span>Load</span><strong>${n(order.EstimatedLoadUnits) || "—"}</strong></div>
          </div>

          <div class="timeline-title">Delivery timeline</div>
          <div class="timeline">
            ${history.length ? history.map((h, i) => `
              <div class="timeline-item">
                <div class="timeline-line"><i></i>${i < history.length-1 ? "<b></b>" : ""}</div>
                <div>
                  <strong>${esc(h.Status || "Update")}</strong>
                  <span>${formatTime(h.CompletionTime || h.ArrivalTime || h.CreatedTime)}</span>
                  ${h.Notes ? `<p>${esc(h.Notes)}</p>` : ""}
                  ${h.DeliveryProblem ? `<p class="problem-text">${esc(h.DeliveryProblem)}</p>` : ""}
                </div>
              </div>
            `).join("") : `<div class="no-history">No delivery-status history has been recorded.</div>`}
          </div>

          ${order.SpecialInstructions ? `
            <div class="notes">
              <span>SPECIAL INSTRUCTIONS</span>
              <p>${esc(order.SpecialInstructions)}</p>
            </div>` : ""}
        </section>
      </div>
    `;
  }

  function settingsModal() {
    return `
      <div class="modal-backdrop" id="settingsBackdrop">
        <section class="settings-modal">
          <button class="close" id="closeSettings">×</button>
          <div class="detail-kicker">SYSTEM</div>
          <h2>Connection & Access</h2>
          <p>This portal is <strong>read only</strong>. Delivery data is controlled in the Excel workbook.</p>
          <div class="security-box">
            <span>ACCESS MODEL</span>
            <strong>View only</strong>
            <p>No add, edit, assignment, status-change or delete actions are available from this portal.</p>
          </div>
          <button class="gold-btn full" id="findWorkbook">Find Workbook in OneDrive</button>
          <button class="dark-btn full" id="clearWorkbook">Clear saved workbook connection</button>
        </section>
      </div>
    `;
  }

  function bindShell() {
    root.querySelectorAll("[data-page]").forEach(btn => {
      btn.addEventListener("click", () => {
        state.page = btn.dataset.page;
        state.query = "";
        state.dateFrom = "";
        state.dateTo = "";
        render();
      });
    });

    document.getElementById("refreshBtn")?.addEventListener("click", refresh);

    document.getElementById("settingsBtn")?.addEventListener("click", () => {
      state.page = "settings";
      render();
    });

    document.querySelectorAll("[data-order]").forEach(btn => {
      btn.addEventListener("click", () => {
        state.selected = state.data.orders.find(o => String(o.OrderID) === String(btn.dataset.order)) || null;
        renderShell();
      });
    });

    document.getElementById("closeModal")?.addEventListener("click", () => {
      state.selected = null;
      render();
    });

    document.getElementById("modalBackdrop")?.addEventListener("click", e => {
      if (e.target.id === "modalBackdrop") {
        state.selected = null;
        render();
      }
    });

    document.getElementById("closeSettings")?.addEventListener("click", () => {
      state.page = "today";
      render();
    });

    document.getElementById("settingsBackdrop")?.addEventListener("click", e => {
      if (e.target.id === "settingsBackdrop") {
        state.page = "today";
        render();
      }
    });

    document.getElementById("findWorkbook")?.addEventListener("click", async () => {
      try {
        await ExcelAPI.findWorkbook();
        state.page = "today";
        await refresh();
      } catch (e) {
        state.error = e.message || "Workbook could not be found.";
        render();
      }
    });

    document.getElementById("clearWorkbook")?.addEventListener("click", () => {
      localStorage.removeItem("mouldingWorkbookDriveId");
      localStorage.removeItem("mouldingWorkbookItemId");
      state.page = "today";
      state.error = "Workbook connection cleared. Open Settings to reconnect.";
      render();
    });

    const todaySearch = document.getElementById("todaySearch");
    todaySearch?.addEventListener("input", e => {
      state.query = e.target.value;
      const pos = e.target.selectionStart;
      render();
      const next = document.getElementById("todaySearch");
      if (next) { next.focus(); next.setSelectionRange(pos, pos); }
    });

    const historySearch = document.getElementById("historySearch");
    historySearch?.addEventListener("input", e => {
      state.query = e.target.value;
      const pos = e.target.selectionStart;
      render();
      const next = document.getElementById("historySearch");
      if (next) { next.focus(); next.setSelectionRange(pos, pos); }
    });

    document.getElementById("dateFrom")?.addEventListener("change", e => {
      state.dateFrom = e.target.value;
      render();
    });

    document.getElementById("dateTo")?.addEventListener("change", e => {
      state.dateTo = e.target.value;
      render();
    });

    document.getElementById("clearFilters")?.addEventListener("click", () => {
      state.query = "";
      state.dateFrom = "";
      state.dateTo = "";
      render();
    });
  }

  function render() {
    renderShell();
  }

  async function boot() {
    render();
    try {
      if (typeof initAuth === "function") await initAuth();
      await refresh();
    } catch (e) {
      state.error = e?.message || "Authentication failed.";
      render();
    }

    setInterval(() => refresh(), Math.max(15, Number(APP_CONFIG.refreshSeconds) || 20) * 1000);
  }

  boot();
})();
