(function () {
    'use strict';

    const API_URL = "https://script.google.com/macros/s/AKfycbwQfabXYdBdQHi8NuzgOCGkt-7rFXTlVH9IQ1VXxDrM4-wMOHU07B8ioeXUIjosp8PeRA/exec";
    const AUTO_REFRESH_MS = 60000;
    const AUTO_REFRESH_MIN_CHANGES = 2;
    const CHART_MIN_HEIGHT = 480;
    const CHART_ROW_HEIGHT = 42;
    const CHART_VERTICAL_PADDING = 160;

    const LOST_STAGES = [
        "Calls",
        "Contact Forms",
        "New Lead No Answer",
        "Answered but not Booked",
        "📞 Calls",
        "📋 Contact Forms",
        "❌ New Lead No Answer",
        "😕 Answered but not Booked"
    ];

    const WON_STAGES = [
        "Estimate Scheduled",
        "Estimate Sent",
        "Estimate Approved",
        "Job Scheduled",
        "Job Complete",
        "Declined",
        "✅ Estimate Scheduled",
        "📤 Estimate Sent",
        "👍 Estimate Approved",
        "🛠️ Job Scheduled",
        "🏁 Job Complete",
        "🚫 Declined"
    ];

    const BAR_COLORS = [
        "#0EA5E9",
        "#F97316",
        "#22C55E",
        "#EAB308",
        "#14B8A6",
        "#EF4444",
        "#8B5CF6",
        "#F43F5E",
        "#10B981",
        "#3B82F6"
    ];

    function getStageStatus(stage) {
        const cleanStage = normalizeStageName(stage);
        const lostSet = getStageSet(LOST_STAGES);
        const wonSet = getStageSet(WON_STAGES);

        if (lostSet.has(cleanStage)) return "lost";
        if (wonSet.has(cleanStage)) return "won";
        return "unknown";
    }

    function normalizeStageName(value) {
        return String(value || "")
            .replace(/[\u{1F300}-\u{1FAFF}]/gu, "")
            .replace(/[\u2600-\u27BF]/g, "")
            .trim()
            .toLowerCase();
    }

    function getStageSet(list) {
        return new Set(list.map(normalizeStageName));
    }

    function getUniqueDisplayStages(list) {
        const map = new Map();

        list.forEach(function (name) {
            const clean = normalizeStageName(name);
            if (!clean) return;
            if (!map.has(clean)) map.set(clean, clean.replace(/\b\w/g, function (c) { return c.toUpperCase(); }));
        });

        return Array.from(map.values());
    }

    function getLogicHtml() {
        const lostNames = getUniqueDisplayStages(LOST_STAGES).join(", ");
        const wonNames = getUniqueDisplayStages(WON_STAGES).join(", ");

        return `<div><b>Lost Stages:</b> ${lostNames} | <b>Won Stages:</b> ${wonNames} | <b>Formula:</b> Conversion % = (Won Opportunities ÷ Total Opportunities) × 100</div>`;
    }

    let state = {
        raw: [],
        filtered: [],
        dateStart: '',
        dateEnd: '',
        autoRefreshTimer: null,
        isLoading: false
    };

    let chartInstance = null;

    const valueLabelPlugin = {
        id: "valueLabelPlugin",
        afterDatasetsDraw: function (chart) {
            const ctx = chart.ctx;
            const meta = chart.getDatasetMeta(0);
            const values = chart.data.datasets[0].data;

            ctx.save();
            ctx.font = "600 11px Inter, Segoe UI, sans-serif";
            ctx.fillStyle = "#ffffff";
            ctx.textAlign = "center";
            ctx.textBaseline = "middle";

            meta.data.forEach(function (bar, i) {
                const val = Number(values[i] || 0).toFixed(1) + "%";
                const centerX = (bar.x + bar.base) / 2;
                const centerY = bar.y;
                ctx.fillText(val, centerX, centerY);
            });

            ctx.restore();
        }
    };

    // =========================
    // LOAD CHART
    // =========================
    function loadChartJS() {
        return new Promise((resolve) => {
            if (window.Chart) return resolve();
            const s = document.createElement("script");
            s.src = "https://cdn.jsdelivr.net/npm/chart.js";
            s.onload = resolve;
            document.head.appendChild(s);
        });
    }

    // =========================
    // JSONP
    // =========================
    function loadData() {
        return new Promise((resolve) => {
            const cb = "cb_" + Date.now();

            window[cb] = function (data) {
                delete window[cb];
                document.body.removeChild(script);
                resolve(data);
            };

            const script = document.createElement("script");
            script.src = API_URL + "?callback=" + cb;
            document.body.appendChild(script);
        });
    }

    function normalizeData(raw) {
        return raw.map(function (r) {
            const stage = r.pipeline_stage || "";

            return {
                id: r.opportunity_id || "",
                updatedAt: r.last_updated || r.date_created || "",
                agent: r.assigned_user_name || "Unassigned",
                stage: stage,
                status: getStageStatus(stage),
                date: new Date(r.date_created),
                name: r.opportunity_name,
                source: r.source,
                email: r.email,
                phone: r.phone,
                opportunityType: r.opportunity_type
            };
        });
    }

    function getRowKey(row) {
        return (row.id || "") + "|" + (row.updatedAt || "");
    }

    function countChanges(previousRows, nextRows) {
        const prevKeys = new Set(previousRows.map(getRowKey));
        const nextKeys = new Set(nextRows.map(getRowKey));

        let changed = 0;

        nextKeys.forEach(function (k) {
            if (!prevKeys.has(k)) changed++;
        });

        prevKeys.forEach(function (k) {
            if (!nextKeys.has(k)) changed++;
        });

        return changed;
    }

    function setLoader(visible, text) {
        const loader = document.getElementById("dashLoader");
        if (!loader) return;

        loader.style.display = visible ? "flex" : "none";
        loader.textContent = text || "Loading data...";
    }

    function updateLiveStatus() {
        const status = document.getElementById("liveStatus");
        if (!status) return;

        status.textContent = "Updated: " + new Date().toLocaleTimeString();
    }

    async function fetchAndApplyData(showLoader) {
        if (state.isLoading) return;

        state.isLoading = true;
        if (showLoader) setLoader(true, "Loading dashboard data...");

        try {
            const raw = await loadData();
            if (!Array.isArray(raw)) return;
            const normalized = normalizeData(raw);

            state.raw = normalized;
            filter();
            updateLiveStatus();
        } finally {
            state.isLoading = false;
            if (showLoader) setLoader(false);
        }
    }

    async function autoRefreshTick() {
        if (state.isLoading) return;

        try {
            const raw = await loadData();
            if (!Array.isArray(raw)) return;
            const normalized = normalizeData(raw);
            const changed = countChanges(state.raw, normalized);

            if (changed >= AUTO_REFRESH_MIN_CHANGES) {
                setLoader(true, "New entries found. Refreshing...");
                state.raw = normalized;
                filter();
                updateLiveStatus();
                setLoader(false);
            }
        } catch (err) {
            console.error("Auto refresh failed", err);
            setLoader(false);
        }
    }

    function startAutoRefresh() {
        if (state.autoRefreshTimer) clearInterval(state.autoRefreshTimer);

        state.autoRefreshTimer = setInterval(autoRefreshTick, AUTO_REFRESH_MS);
    }

    // =========================
    // INIT
    // =========================
    async function init() {
        const parent = document.querySelector('.dashboard-divider');
        if (!parent) return;

        await loadChartJS();

        const wrapper = document.createElement('div');

        wrapper.innerHTML = `
      <div style="background:linear-gradient(135deg,#f8fafc,#eef2ff);padding:24px;border-radius:16px;box-shadow:0 12px 38px rgba(2,6,23,0.12);font-family:Inter,Segoe UI,sans-serif;border:1px solid #e2e8f0;position:relative;">
        <div id="dashLoader" style="display:none;position:absolute;inset:0;background:rgba(255,255,255,0.72);backdrop-filter:blur(2px);align-items:center;justify-content:center;font-weight:600;color:#1e293b;border-radius:16px;z-index:3;">Loading data...</div>
        
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
          <h2 style="margin:0;font-size:24px;color:#0f172a;letter-spacing:0.2px;">Agent Performance</h2>
          <div style="display:flex;gap:8px;align-items:center;">
            <input type="date" id="startDate" style="border:1px solid #cbd5e1;border-radius:8px;padding:7px 10px;background:#fff;">
            <input type="date" id="endDate" style="border:1px solid #cbd5e1;border-radius:8px;padding:7px 10px;background:#fff;">
            <button id="refreshNow" style="border:none;border-radius:8px;padding:8px 12px;background:#0ea5e9;color:#fff;font-weight:600;cursor:pointer;">Refresh</button>
            <span id="liveStatus" style="font-size:12px;color:#475569;min-width:130px;">Updated: --</span>
          </div>
        </div>

        <div style="background:#eff6ff;padding:12px;border-radius:10px;margin-bottom:18px;border:1px solid #bfdbfe;color:#1e3a8a;line-height:1.45;">
          ${getLogicHtml()}
        </div>

        <div style="display:flex;gap:18px;flex-wrap:wrap;align-items:stretch;">
          <div id="agentChartCard" style="flex:1;min-width:360px;background:#fff;border:1px solid #dbeafe;border-radius:14px;padding:14px;box-shadow:0 6px 16px rgba(14,116,144,0.08);height:420px;overflow:hidden;">
            <div style="font-size:13px;font-weight:600;color:#0c4a6e;margin-bottom:10px;">Conversion by Agent</div>
            <canvas id="agentChart"></canvas>
          </div>

          <div style="flex:1;min-width:320px;background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:10px 12px;box-shadow:0 8px 20px rgba(15,23,42,0.06);">
            <div style="font-size:13px;font-weight:600;color:#334155;margin:4px 4px 10px;">Agent Summary Table</div>
            <table style="width:100%;border-collapse:separate;border-spacing:0 6px;">
              <thead style="font-size:12px;color:#475569;text-transform:uppercase;letter-spacing:.5px;">
                <tr>
                  <th style="text-align:left;padding:6px 8px;">Agent</th>
                  <th style="text-align:right;padding:6px 8px;">Total</th>
                  <th style="text-align:right;padding:6px 8px;">Won</th>
                  <th style="text-align:right;padding:6px 8px;">Lost</th>
                  <th style="text-align:right;padding:6px 8px;">Conversion(%)</th>
                </tr>
              </thead>
              <tbody id="agentTable"></tbody>
            </table>
          </div>
        </div>
      </div>
    `;

        parent.appendChild(wrapper);

        attachEvents();
        await fetchAndApplyData(true);
        startAutoRefresh();
    }

    // =========================
    // FILTER
    // =========================
    function filter() {
        let data = state.raw;

        // 🔥 Skip delete events
        data = data.filter(d => d.opportunityType !== 'OpportunityDelete');

        console.log(data, 'res');

        if (state.dateStart) {
            const s = new Date(state.dateStart);
            data = data.filter(d => d.date >= s);
        }

        if (state.dateEnd) {
            const e = new Date(state.dateEnd);
            e.setHours(23, 59, 59);
            data = data.filter(d => d.date <= e);
        }

        state.filtered = data;
        render();
    }

    // =========================
    // STATS
    // =========================
    function getStats() {
        let map = {};

        state.filtered.forEach(o => {
            if (!map[o.agent]) map[o.agent] = { total: 0, won: 0, lost: 0 };

            map[o.agent].total++;

            if (o.status === "won") map[o.agent].won++;
            else map[o.agent].lost++;
        });

        return Object.keys(map).map(agent => {
            const t = map[agent].total;
            const w = map[agent].won;

            return {
                agent,
                total: t,
                won: w,
                lost: map[agent].lost,
                conversion: t ? (w / t * 100) : 0
            };
        }).sort((a, b) => b.conversion - a.conversion);
    }

    // =========================
    // RENDER
    // =========================
    function render() {
        const data = getStats();

        const tbody = document.getElementById("agentTable");

        tbody.innerHTML = data.map(d => `
      <tr style="cursor:pointer;" onclick="showAgentModal('${d.agent}')">
        <td style="padding:10px 8px;background:#f8fafc;border-top-left-radius:10px;border-bottom-left-radius:10px;font-weight:600;color:#0f172a;">${d.agent}</td>
        <td style="padding:10px 8px;text-align:right;background:#f8fafc;color:#334155;">${d.total}</td>
        <td style="padding:10px 8px;text-align:right;background:#ecfdf5;color:#15803d;font-weight:600;">${d.won}</td>
        <td style="padding:10px 8px;text-align:right;background:#fff1f2;color:#be123c;font-weight:600;">${d.lost}</td>
        <td style="padding:10px 8px;text-align:right;background:#f8fafc;border-top-right-radius:10px;border-bottom-right-radius:10px;"><b style="color:#1d4ed8;">${d.conversion.toFixed(1)}%</b></td>
      </tr>
    `).join('');

        renderChart(data);
    }

    function renderChart(data) {
        const ctx = document.getElementById("agentChart");
        const chartCard = document.getElementById("agentChartCard");
        if (chartCard) {
            var dynamicHeight = Math.max(CHART_MIN_HEIGHT, (data.length * CHART_ROW_HEIGHT) + CHART_VERTICAL_PADDING);
            chartCard.style.minHeight = CHART_MIN_HEIGHT + "px";
            chartCard.style.height = dynamicHeight + "px";
        }

        if (chartInstance) chartInstance.destroy();

        chartInstance = new Chart(ctx, {
            type: 'bar',
            plugins: [valueLabelPlugin],
            data: {
                labels: data.map(d => d.agent),
                datasets: [{
                    label: 'Conversion %',
                    data: data.map(d => d.conversion),
                    backgroundColor: data.map((_, i) => BAR_COLORS[i % BAR_COLORS.length]),
                    borderColor: data.map((_, i) => BAR_COLORS[i % BAR_COLORS.length]),
                    borderWidth: 1,
                    borderRadius: 8
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                layout: {
                    padding: {
                        right: 8,
                        bottom: 18
                    }
                },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function (context) {
                                return "Conversion: " + Number(context.parsed.x).toFixed(1) + "%";
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        beginAtZero: true,
                        max: 100,
                        ticks: {
                            stepSize: 10,
                            callback: function (v) { return v + "%"; }
                        },
                        grid: { color: "rgba(148,163,184,0.25)" }
                    },
                    y: {
                        grid: { display: false },
                        ticks: {
                            autoSkip: false,
                            callback: function (value) {
                                const label = this.getLabelForValue(value) || "";
                                return label.length > 24 ? (label.slice(0, 24) + "...") : label;
                            }
                        }
                    }
                }
            }
        });
    }

    // =========================
    // MODAL (ADVANCED)
    // =========================
    function showAgentModal(agent) {
        const opps = state.filtered.filter(o => o.agent === agent);

        const total = opps.length;
        const won = opps.filter(o => o.status === "won").length;
        const lost = total - won;
        const conv = total ? (won / total * 100) : 0;

        const old = document.getElementById("agent-modal");
        if (old) old.remove();

        const modal = document.createElement("div");
        modal.id = "agent-modal";

        modal.style = `
    position:fixed;
    top:0; left:0;
    width:100%; height:100%;
    background:rgba(15,23,42,0.55);
    backdrop-filter: blur(4px);
    display:flex;
    align-items:center;
    justify-content:center;
    z-index:9999;
  `;

        const box = document.createElement("div");

        box.style = `
    background:#fff;
    width:95%;
    max-width:1000px;
    max-height:85vh;
    border-radius:14px;
    overflow:hidden;
    display:flex;
    flex-direction:column;
    box-shadow:0 20px 60px rgba(0,0,0,0.2);
    font-family: Inter, sans-serif;
  `;

        box.innerHTML = `
    
    <!-- HEADER -->
    <div style="
      padding:16px 20px;
      border-bottom:1px solid #eee;
      display:flex;
      justify-content:space-between;
      align-items:center;
    ">
      <div>
        <h3 style="margin:0;font-size:18px;">${agent}</h3>
        <span style="font-size:13px;color:#6b7280;">
          ${total} Opportunities
        </span>
      </div>

      <button id="closeModal" style="
        border:none;
        background:#f3f4f6;
        border-radius:8px;
        padding:6px 10px;
        cursor:pointer;
        font-size:14px;
      ">✕</button>
    </div>

    <!-- STATS -->
    <div style="
      display:flex;
      gap:12px;
      padding:15px 20px;
      border-bottom:1px solid #eee;
    ">
      ${statCard("Total", total, "#111")}
      ${statCard("Won", won, "#16a34a")}
      ${statCard("Lost", lost, "#dc2626")}
      ${statCard("Conversion", conv.toFixed(1) + "%", "#4f46e5")}
    </div>

    <!-- TABLE -->
    <div style="flex:1; overflow:auto;">
      <table style="
        width:100%;
        border-collapse:collapse;
        font-size:14px;
      ">
        
        <thead style="
          position:sticky;
          top:0;
          background:#f9fafb;
          z-index:2;
        ">
          <tr>
            <th style="padding:12px;text-align:left;">Name</th>
            <th>Stage</th>
            <th>Status</th>
            <th>Date</th>
            <th>Email</th>
            <th>Phone</th>
          </tr>
        </thead>

        <tbody>
          ${opps.map(o => `
            <tr style="border-bottom:1px solid #f1f5f9;">
              
              <td style="padding:12px;">
                <div style="font-weight:500;">${o.name || '-'}</div>
                <div style="font-size:12px;color:#6b7280;">
                  ${o.source || ''}
                </div>
              </td>

              <td style="color:#374151;">
                ${o.stage}
              </td>

              <td>
                <span style="
                  padding:4px 10px;
                  border-radius:999px;
                  font-size:12px;
                  font-weight:500;
                  color:#fff;
                  background:${o.status === 'won' ? '#16a34a' : '#dc2626'};
                ">
                  ${o.status}
                </span>
              </td>

              <td style="color:#6b7280;">
                ${o.date.toLocaleDateString()}
              </td>

              <td style="color:#374151;">
                ${o.email || '-'}
              </td>

              <td style="color:#374151;">
                ${o.phone || '-'}
              </td>

            </tr>
          `).join('')}
        </tbody>

      </table>
    </div>
  `;

        modal.appendChild(box);
        document.body.appendChild(modal);

        document.getElementById("closeModal").onclick = () => modal.remove();

        modal.onclick = (e) => {
            if (e.target === modal) modal.remove();
        };
    }

    // =========================
    // SMALL STAT CARD
    // =========================
    function statCard(label, value, color) {
        return `
    <div style="
      flex:1;
      background:#f9fafb;
      padding:10px;
      border-radius:10px;
      text-align:center;
    ">
      <div style="font-size:12px;color:#6b7280;">
        ${label}
      </div>
      <div style="
        font-size:18px;
        font-weight:600;
        color:${color};
      ">
        ${value}
      </div>
    </div>
  `;
    }

    window.showAgentModal = showAgentModal;

    function attachEvents() {
        document.addEventListener("change", (e) => {
            if (e.target.id === "startDate") {
                state.dateStart = e.target.value;
                filter();
            }
            if (e.target.id === "endDate") {
                state.dateEnd = e.target.value;
                filter();
            }
        });

        document.addEventListener("click", function (e) {
            if (e.target && e.target.id === "refreshNow") {
                fetchAndApplyData(true);
            }
        });
    }

    init();

})();