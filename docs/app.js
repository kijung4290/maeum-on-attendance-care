const state = { data: null, filter: "all", query: "" };
const $ = (selector) => document.querySelector(selector);
const statusName = { present: "출석", absent: "결석", late: "지각", excused: "인정결석" };

async function loadDashboard() {
  try {
    const isStatic = location.hostname.endsWith("github.io");
    let response = await fetch(isStatic ? "./demo-dashboard.json" : "/api/dashboard");
    if (!response.ok && !isStatic) response = await fetch("./demo-dashboard.json");
    if (!response.ok) throw new Error("데이터를 불러오지 못했습니다.");
    state.data = await response.json();
    state.data.staticMode = isStatic || response.url.includes("demo-dashboard.json");
    render();
  } catch (error) {
    $("#member-list").innerHTML = `<tr><td colspan="6" class="loading-row">${escapeHtml(error.message)}</td></tr>`;
    showToast(error.message);
  }
}

function render() {
  const { summary, mode, meta } = state.data;
  $("#sum-high").textContent = summary.high;
  $("#sum-watch").textContent = summary.watch;
  $("#sum-stable").textContent = summary.stable;
  $("#sum-rate").textContent = summary.avgAttendance;
  $("#sum-total").textContent = `총 ${summary.total}명 기준`;
  $("#reference-date").textContent = formatDate(meta.referenceDate, true);
  renderMode(mode);
  renderBrief();
  renderMembers();
  renderReports();
}

function renderMode(mode) {
  const live = mode === "jev";
  $("#mode-banner").classList.toggle("demo", !live);
  $("#mode-title").textContent = live ? "TypeSafe AI · JEV 연결됨" : "데모 분석 모드";
  $("#mode-copy").textContent = live ? "실제 JEV 확률 판정을 사용할 준비가 되었습니다." : "API 키 없이 로컬 규칙으로 안전하게 기능을 시험하고 있습니다.";
}

function renderBrief() {
  const high = state.data.members.filter((m) => m.risk.choice === "high").sort((a,b) => b.risk.confidence-a.risk.confidence);
  if (high.length) {
    $("#brief-title").textContent = `${high[0].name} 어르신 외 ${Math.max(0, high.length - 1)}명을 먼저 살펴봐 주세요`;
    $("#brief-copy").textContent = `${high[0].risk.reasons[0]}. 기록만으로 단정하지 말고 안부와 상황을 확인해 주세요.`;
  } else {
    $("#brief-title").textContent = "현재 긴급하게 확인할 출석 신호는 없어요";
    $("#brief-copy").textContent = "관심 관찰 대상의 다음 출석도 가볍게 살펴봐 주세요.";
  }
}

function renderMembers() {
  const query = state.query.trim().toLowerCase();
  const members = state.data.members.filter((m) => (state.filter === "all" || m.risk.choice === state.filter) && (!query || m.name.includes(query) || m.id.toLowerCase().includes(query)));
  $("#empty-state").hidden = members.length > 0;
  $(".table-wrap").hidden = members.length === 0;
  $("#member-list").innerHTML = members.map((m) => `
    <tr>
      <td><div class="person"><span class="person-avatar">${m.name[0]}</span><div><b>${escapeHtml(m.name)}</b><small>${m.id} · 만 ${m.age}세</small></div></div></td>
      <td><span class="risk-pill risk-${m.risk.choice}">${m.risk.meta.label}</span><small class="risk-confidence">확신도 ${toPercent(m.risk.confidence)}</small></td>
      <td><div class="attendance-strip">${m.features.records.map((r) => `<i class="attendance-dot ${r.status}" data-tip="${formatDate(r.date)} · ${statusName[r.status]}"></i>`).join("")}</div></td>
      <td><span class="rate">${m.features.attendanceRate}%</span><small class="rate"><br>${m.features.counts.present + m.features.counts.late}/${m.features.records.length - m.features.counts.excused}회</small></td>
      <td class="reason">${escapeHtml(m.risk.reasons[0])}</td>
      <td><button class="detail-button" data-id="${m.id}">자세히 보기</button></td>
    </tr>`).join("");
  document.querySelectorAll(".detail-button").forEach((button) => button.addEventListener("click", () => openDrawer(button.dataset.id)));
}

function renderReports() {
  const { members, summary } = state.data;
  const total = summary.total || 1;
  $("#risk-chart").innerHTML = `
    <span class="high" style="width:${summary.high / total * 100}%"></span>
    <span class="watch" style="width:${summary.watch / total * 100}%"></span>
    <span class="stable" style="width:${summary.stable / total * 100}%"></span>`;
  $("#risk-legend").innerHTML = [
    ["high", "집중 확인", summary.high], ["watch", "관심 관찰", summary.watch], ["stable", "안정", summary.stable]
  ].map(([type, label, count]) => `<div class="legend-item ${type}"><small><i></i>${label}</small><b>${count}명</b></div>`).join("");
  $("#report-rate").textContent = summary.avgAttendance;
  $("#report-absence").textContent = `${members.filter((m) => m.features.recentAbsences > 0).length}명`;
  $("#report-streak").textContent = `${members.filter((m) => m.features.consecutiveAbsences >= 2).length}명`;
  const priority = members.filter((m) => m.risk.choice !== "stable").sort((a, b) => {
    const rank = { high: 2, watch: 1 };
    return rank[b.risk.choice] - rank[a.risk.choice] || b.risk.confidence - a.risk.confidence;
  }).slice(0, 10);
  $("#priority-list").innerHTML = priority.map((m) => `<div class="priority-item"><span class="person-avatar">${m.name[0]}</span><div class="priority-info"><b>${escapeHtml(m.name)} 어르신</b><p>${escapeHtml(m.risk.reasons[0])}</p></div><button class="detail-button" data-report-id="${m.id}">확인</button><span class="risk-pill risk-${m.risk.choice}">${m.risk.meta.label}</span></div>`).join("");
  document.querySelectorAll("[data-report-id]").forEach((button) => button.addEventListener("click", () => openDrawer(button.dataset.reportId)));
}

function navigate(view, updateHash = true) {
  const target = ["dashboard", "members", "reports"].includes(view) ? view : "dashboard";
  document.querySelectorAll(".view-page").forEach((page) => { page.hidden = page.id !== target; });
  document.querySelectorAll("[data-view]").forEach((link) => link.classList.toggle("active", link.dataset.view === target));
  if (updateHash) history.replaceState(null, "", `#${target}`);
  $(".sidebar").classList.remove("open");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function openDrawer(id) {
  const m = state.data.members.find((item) => item.id === id);
  if (!m) return;
  $("#drawer-content").innerHTML = `
    <div class="drawer-profile"><div class="drawer-avatar">${m.name[0]}</div><div><h2>${escapeHtml(m.name)} 어르신</h2><p>${m.id} · 만 ${m.age}세 · ${m.gender} · 등록 ${formatDate(m.joinedAt, true)}</p></div></div>
    <section class="drawer-risk ${m.risk.choice}">
      <div class="drawer-risk-header"><h3><span class="risk-pill risk-${m.risk.choice}">${m.risk.meta.label}</span></h3><span class="confidence-ring">확신도 ${toPercent(m.risk.confidence)}</span></div>
      <ul>${m.risk.reasons.map((reason) => `<li>${escapeHtml(reason)}</li>`).join("")}</ul>
    </section>
    <section class="detail-block"><h3>기본 연락 정보</h3><div class="info-grid"><div class="info-item"><small>본인 연락처</small><b>${m.phone}</b></div><div class="info-item"><small>보호자 / 관계</small><b>${escapeHtml(m.guardian)}</b></div><div class="info-item"><small>출석률</small><b>${m.features.attendanceRate}%</b></div><div class="info-item"><small>최근 연속 결석</small><b>${m.features.consecutiveAbsences}회</b></div></div></section>
    <section class="detail-block"><h3>담당자 메모</h3><div class="memo">${escapeHtml(m.memo)}</div></section>
    <section class="detail-block"><h3>최근 12회 출석 기록</h3><div class="history">${m.features.records.map((r) => `<div class="history-item"><i class="${r.status}"></i><small>${r.date.slice(5).replace("-",".")}<br>${statusName[r.status]}</small></div>`).join("")}</div></section>
    <button class="drawer-action" data-analyze="${m.id}">✦ 이 참여자 다시 분석</button>
    <p class="source-line">${m.risk.source === "jev" ? "TypeSafe AI JEV 분석 결과" : "로컬 데모 규칙 분석 결과"} · 최종 판단은 담당자에게 있습니다.</p>`;
  $("#detail-drawer").classList.add("open");
  $("#drawer-backdrop").classList.add("open");
  $("#detail-drawer").setAttribute("aria-hidden", "false");
  $("[data-analyze]").addEventListener("click", (event) => analyzeOne(event.currentTarget.dataset.analyze, event.currentTarget));
}

function closeDrawer() {
  $("#detail-drawer").classList.remove("open");
  $("#drawer-backdrop").classList.remove("open");
  $("#detail-drawer").setAttribute("aria-hidden", "true");
}

async function analyzeOne(id, button) {
  if (state.data.staticMode) {
    showToast("공개 데모에서는 저장된 분석 결과를 사용합니다. 실제 JEV 분석은 서버 실행 시 이용할 수 있어요.");
    return;
  }
  const before = button.textContent;
  button.disabled = true; button.textContent = "분석 중...";
  try {
    const response = await fetch(`/api/analyze/${encodeURIComponent(id)}`, { method: "POST" });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "분석하지 못했습니다.");
    const index = state.data.members.findIndex((m) => m.id === id);
    state.data.members[index] = result;
    recalculateSummary(); render(); openDrawer(id);
    showToast(result.risk.source === "jev" ? "JEV 분석을 완료했습니다." : "데모 분석을 완료했습니다.");
  } catch (error) { showToast(error.message); }
  finally { button.disabled = false; button.textContent = before; }
}

async function analyzeAll() {
  if (state.data.staticMode) {
    showToast("100명의 데모 분석 결과를 새로 불러왔습니다.");
    await loadDashboard();
    return;
  }
  const button = $("#analyze-all");
  button.disabled = true; button.innerHTML = '<span class="loader"></span> 전체 분석 중...';
  try {
    const response = await fetch("/api/analyze-all", { method: "POST" });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "분석하지 못했습니다.");
    state.data.members = result.members; state.data.mode = result.mode;
    recalculateSummary(); render();
    showToast(result.mode === "jev" ? "전체 JEV 분석이 완료되었습니다." : "전체 데모 분석이 완료되었습니다.");
  } catch (error) { showToast(error.message); }
  finally { button.disabled = false; button.innerHTML = "<span>✦</span> AI로 전체 다시 분석"; }
}

function recalculateSummary() {
  const members = state.data.members;
  state.data.summary = { total: members.length, high: members.filter(m=>m.risk.choice==="high").length, watch: members.filter(m=>m.risk.choice==="watch").length, stable: members.filter(m=>m.risk.choice==="stable").length, avgAttendance: Math.round(members.reduce((s,m)=>s+m.features.attendanceRate,0)/members.length) };
}

function showToast(message) {
  const toast = $("#toast"); toast.textContent = message; toast.classList.add("show");
  clearTimeout(showToast.timer); showToast.timer = setTimeout(() => toast.classList.remove("show"), 2800);
}

function formatDate(value, long = false) { const date = new Date(`${value}T00:00:00`); return long ? `${date.getFullYear()}. ${date.getMonth()+1}. ${date.getDate()}.` : `${date.getMonth()+1}/${date.getDate()}`; }
function toPercent(value) { return `${Math.round(Number(value || 0) * 100)}%`; }
function escapeHtml(value) { return String(value).replace(/[&<>'"]/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[char])); }

$("#today").textContent = new Intl.DateTimeFormat("ko-KR", { year:"numeric", month:"long", day:"numeric", weekday:"short" }).format(new Date());
$("#search").addEventListener("input", (event) => { state.query = event.target.value; renderMembers(); });
document.querySelectorAll(".filter-button").forEach((button) => button.addEventListener("click", () => { document.querySelectorAll(".filter-button").forEach(b=>b.classList.remove("active")); button.classList.add("active"); state.filter=button.dataset.filter; renderMembers(); }));
$("#drawer-close").addEventListener("click", closeDrawer); $("#drawer-backdrop").addEventListener("click", closeDrawer);
$("#analyze-all").addEventListener("click", analyzeAll);
$("#show-priority").addEventListener("click", () => { state.filter="high"; document.querySelectorAll(".filter-button").forEach(b=>b.classList.toggle("active",b.dataset.filter==="high")); renderMembers(); navigate("members"); });
$("#open-guide").addEventListener("click", () => $("#guide-dialog").showModal());
$(".dialog-close").addEventListener("click", () => $("#guide-dialog").close());
$(".menu-button").addEventListener("click", () => $(".sidebar").classList.toggle("open"));
document.querySelectorAll("[data-view]").forEach((link) => link.addEventListener("click", (event) => { event.preventDefault(); navigate(link.dataset.view); }));
$("#print-report").addEventListener("click", () => window.print());
document.addEventListener("keydown", (event) => { if(event.key === "Escape") closeDrawer(); });
navigate(location.hash.slice(1) || "dashboard", false);
loadDashboard();
