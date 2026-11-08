/* app.js - 버튼 동작과 결과 표시. 계산은 worker.js 의 Python(model.py)이 한다. */
(function () {
  "use strict";

  // ---------------------------------------------------------------- 도우미
  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  };
  var nf = function (v, d) {
    if (v == null || (typeof v === "number" && !isFinite(v))) return "-";
    d = d || 0;
    return Number(v).toLocaleString("ko-KR", { minimumFractionDigits: d, maximumFractionDigits: d });
  };
  var jo = function (v, d) { return nf(v / 1e6, d == null ? 1 : d); };           // 백만원 -> 조원 숫자
  var eok = function (v, d) { return nf(v / 100, d == null ? 1 : d); };           // 백만원 -> 억원 숫자
  var human = function (v) {                                                       // 단위를 붙여 읽기 쉽게
    if (v == null) return "-";
    var a = Math.abs(v);
    if (a >= 1e6) return nf(v / 1e6, 2) + "조원";
    if (a >= 100) return nf(v / 100, 1) + "억원";
    return nf(v, 0) + "백만원";
  };
  var pct = function (v, d) { return v == null ? "-" : nf(v, d == null ? 1 : d) + "%"; };

  var LEVELS = {
    large: { label: "대분류", file: "io-2023-large.xlsx" },
    mid: { label: "중분류", file: "io-2023-mid.xlsx" },
    small: { label: "소분류", file: "io-2023-small.xlsx" },
    basic: { label: "기본부문", file: "io-2023-basic.xlsx" }
  };
  var FMT = {            // 지표별 소수 자릿수
    out: 0, out_share: 2, va_ratio: 1, mid_ratio: 1, imp_dep: 1, dom_ratio: 1, mid_dem_ratio: 1, export_rate: 1,
    import_rate: 1, prod: 3, va_mult: 3, im_mult: 3, infl: 3, sens: 3
  };

  var state = { ready: false, loaded: false, level: "large", info: null, metrics: null, inds: [], labelToCode: {},
    sheet: { key: "total", c0: 0, cn: 10 }, rank: { sortKey: "out", sortDir: -1, type: "" }, sim: [], fdItem: "9190" };

  // ---------------------------------------------------------------- 작업자(Python) 연결
  var worker = null, seq = 0, pending = {};
  function call(cmd, args) {
    return new Promise(function (resolve, reject) {
      var id = ++seq;
      pending[id] = { resolve: resolve, reject: reject };
      worker.postMessage({ id: id, cmd: cmd, args: args || {} });
    });
  }
  function setEngine(text, cls) {
    var el = $("engine");
    el.textContent = text;
    el.className = "engine" + (cls ? " " + cls : "");
  }
  function onWorkerMessage(e) {
    var m = e.data;
    if (m.type === "status") setEngine(m.text);
    else if (m.type === "ready") {
      state.ready = true;
      setEngine("Python 엔진 준비 완료", "ready");
      $("btn-load").disabled = false;
      $("load-msg").textContent = "분류를 고르고 불러오기를 누르세요.";
    } else if (m.type === "fail") {
      setEngine("Python 엔진 오류", "fail");
      $("load-msg").textContent = "엔진을 시작하지 못했습니다: " + m.text + " (인터넷 연결을 확인하고 새로고침하세요)";
    } else if (m.type === "progress") {
      $("load-msg").textContent = m.text;
      if (m.pct != null) { $("load-bar").hidden = false; $("load-bar").value = m.pct; }
    } else if (m.type === "result") {
      var p = pending[m.id];
      delete pending[m.id];
      if (!p) return;
      var r;
      try { r = JSON.parse(m.json); } catch (err) { p.reject(new Error("결과를 해석하지 못했습니다.")); return; }
      if (r.ok) p.resolve(r.data); else p.reject(new Error(r.error || "알 수 없는 오류"));
    }
  }
  function startWorker() {
    try {
      worker = new Worker("worker.js");
      worker.onmessage = onWorkerMessage;
      worker.onerror = function (ev) { setEngine("Python 엔진 오류", "fail"); $("load-msg").textContent = "작업자 오류: " + (ev.message || ""); };
    } catch (err) {
      setEngine("Python 엔진 오류", "fail");
      $("load-msg").textContent = "이 브라우저에서는 Web Worker를 시작할 수 없습니다: " + err.message;
    }
  }

  // ---------------------------------------------------------------- 공통 표시 조각
  function bars(items, o) {
    o = o || {};
    var max = 0;
    items.forEach(function (it) { max = Math.max(max, Math.abs(it.v == null ? 0 : it.v)); });
    if (!max) max = 1;
    return '<div class="bars">' + items.map(function (it) {
      var w = Math.max(0, Math.min(100, (Math.abs(it.v || 0) / max) * 100));
      return '<div class="bar"><em title="' + esc(it.label) + '">' + esc(it.label) + '</em><div class="track"><i class="fill ' +
        (it.v < 0 ? "neg " : "") + (o.cls || "") + '" style="width:' + w.toFixed(1) + '%"></i></div><span>' + esc(it.text) + "</span></div>";
    }).join("") + "</div>";
  }
  function table(heads, rows, o) {
    o = o || {};
    var h = "<table><thead><tr>" + heads.map(function (x, i) {
      var cls = (x.l ? "l " : "") + (o.sortable && x.key ? "sort" : "");
      return '<th class="' + cls + '"' + (x.key ? ' data-k="' + esc(x.key) + '"' : "") + ">" + x.t + "</th>";
    }).join("") + "</tr></thead><tbody>";
    h += rows.map(function (r) {
      return "<tr>" + r.map(function (c, i) {
        var left = heads[i] && heads[i].l;
        return '<td class="' + (left ? "l" : "") + '">' + c + "</td>";
      }).join("") + "</tr>";
    }).join("");
    return h + "</tbody></table>";
  }
  function showErr(el, err) { el.innerHTML = '<p class="note warn">' + esc(err.message || err) + "</p>"; }

  // ---------------------------------------------------------------- 1. 불러오기
  function selectedLevel() {
    var radios = document.querySelectorAll('input[name="level"]');
    for (var i = 0; i < radios.length; i++) if (radios[i].checked) return radios[i].value;
    return "large";
  }
  async function loadData() {
    if (!state.ready) return;
    var lv = selectedLevel();
    var cfg = LEVELS[lv];
    $("btn-load").disabled = true;
    $("load-bar").hidden = false;
    $("load-bar").value = 0;
    $("load-note").hidden = false;
    $("load-note").textContent = lv === "basic"
      ? "기본부문(12MB)은 브라우저에서 Excel을 읽는 데 1~2분 걸릴 수 있습니다. 창을 닫지 말고 기다려 주세요."
      : "Excel을 읽고 지표를 계산하고 있습니다…";
    try {
      var info = await call("load", { level: lv, file: cfg.file, label: cfg.label });
      state.level = lv;
      state.info = info;
      var res = await Promise.all([call("metrics"), call("industries")]);
      state.metrics = res[0];
      state.inds = res[1].items;
      state.labelToCode = {};
      state.inds.forEach(function (x) { state.labelToCode[x.c + " " + x.n] = x.c; });
      state.loaded = true;
      state.rank = { sortKey: "out", sortDir: -1, type: "" };
      state.sheet = { key: "total", c0: 0, cn: 10 };
      state.sim = [{ code: "", label: "", amount: "", unit: "억원" }];
      $("load-msg").textContent = "✅ " + cfg.label + " " + info.n + "개 부문을 불러왔습니다.";
      $("load-note").hidden = true;
      $("app").hidden = false;
      renderAll();
    } catch (err) {
      $("load-msg").textContent = "❌ 불러오기 실패: " + err.message;
      $("load-note").hidden = true;
    } finally {
      $("btn-load").disabled = false;
      $("load-bar").hidden = true;
    }
  }

  function renderAll() {
    renderSummary();
    renderSheetsSetup();
    $("s-q").value = "";
    renderSearchList();
    $("s-profile").innerHTML = "";
    renderRankSetup();
    renderRank();
    renderFd();
    renderSimSetup();
    $("sim-result").innerHTML = "";
    $("tb-deep").innerHTML = "";
    $("deep-msg").textContent = "";
    showTab("summary");
  }

  // ---------------------------------------------------------------- 탭
  function showTab(name) {
    var btns = document.querySelectorAll("#tabs button");
    for (var i = 0; i < btns.length; i++) btns[i].classList.toggle("on", btns[i].getAttribute("data-tab") === name);
    ["summary", "sheets", "search", "rank", "fd", "sim", "guide"].forEach(function (t) { $("pane-" + t).hidden = t !== name; });
    if (name === "sheets" && state.loaded) loadSheet();
  }

  // ---------------------------------------------------------------- 요약·검증
  function renderSummary() {
    var inf = state.info, nat = inf.national, T = nat.T;
    $("sum-level").textContent = "· " + inf.level + " " + inf.n + "개 부문 기준 (전산업 합계는 분류와 관계없이 같음)";
    var k = [
      ["총산출 (국내 공급)", nf(nat.out / T, 1) + "조원"], ["수입 (수입계)", nf(nat.imp_all / T, 1) + "조원"],
      ["총공급 = 총수요", nf(nat.sup / T, 1) + "조원"], ["중간투입률", pct(nat.mid_ratio)], ["부가가치율", pct(nat.va_ratio)],
      ["수출률 (수출 ÷ 총산출)", pct(nat.export_rate)], ["수입률 (수입 ÷ 총공급)", pct(nat.import_rate)],
      ["수입의존도", pct(nat.imp_dep)], ["대외거래 비중", pct(nat.trade_share)]
    ];
    $("kpis").innerHTML = k.map(function (x) { return '<div class="kpi"><b>' + x[1] + "</b><span>" + x[0] + "</span></div>"; }).join("");
    var rowsOf = function (arr) { return arr.map(function (x) { return [esc(x.n), jo(x.v, 1), pct(x.share)]; }); };
    var heads = [{ t: "항목", l: 1 }, { t: "금액(조원)" }, { t: "비중(%)" }];
    $("tb-fd").innerHTML = table(heads, rowsOf(nat.fd_items).concat([["최종수요계", jo(nat.fd, 1), "100.0%"]]));
    $("tb-va").innerHTML = table(heads, rowsOf(nat.va_items).concat([["부가가치계", jo(nat.va, 1), "100.0%"]]));
    $("tb-press").innerHTML = table(
      [{ t: "항목", l: 1 }, { t: "보도자료(조원)" }, { t: "Excel 합계(조원)" }, { t: "차이" }, { t: "판정" }],
      inf.press.map(function (p) {
        return [esc(p.label), nf(p.press, 1), nf(p.calc, 1), nf(p.diff, 2), p.ok ? '<span class="ok">✓ 일치</span>' : '<span class="bad">✗ 차이</span>'];
      }));
    $("tb-acct").innerHTML = checksTable(inf.checks);
  }
  function checksTable(list) {
    return table(
      [{ t: "검증 항목", l: 1 }, { t: "산식", l: 1 }, { t: "최대 차이" }, { t: "허용 오차" }, { t: "판정" }],
      list.map(function (c) {
        return [esc(c.name), esc(c.formula), nf(c.maxdiff, c.unit === "백만원" ? 3 : 10) + " " + esc(c.unit), esc(c.tol),
          c.ok ? '<span class="ok">✓ 통과</span>' : '<span class="bad">✗ 확인 필요</span>'];
      }));
  }
  async function runDeep() {
    $("btn-deep").disabled = true;
    $("deep-msg").textContent = "계산 중… (Excel의 계수표를 추가로 읽습니다)";
    try {
      var d = await call("validate");
      $("tb-deep").innerHTML = checksTable(d.checks) +
        (d.zero_output ? '<p class="muted">총산출이 0인 부문 ' + d.zero_output + "개는 계수 비교에서 제외했습니다.</p>" : "");
      $("deep-msg").textContent = "완료";
    } catch (err) { showErr($("tb-deep"), err); $("deep-msg").textContent = ""; }
    $("btn-deep").disabled = false;
  }

  // ---------------------------------------------------------------- 시트 보기
  function renderSheetsSetup() {
    $("sheet-sel").innerHTML = state.info.sheets.map(function (s) {
      return '<option value="' + esc(s.key) + '"' + (s.present ? "" : " disabled") + ">" + esc(s.label) + " (" + esc(s.unit) + ")</option>";
    }).join("");
    $("sheet-sel").value = state.sheet.key;
    $("sheet-rq").value = "";
    $("sheet-cq").value = "";
  }
  async function loadSheet() {
    var st = state.sheet;
    $("sheet-table").innerHTML = '<p class="muted" style="padding:10px">시트를 읽는 중…</p>';
    try {
      var d = await call("view", { key: st.key, q: $("sheet-rq").value, cq: $("sheet-cq").value, c0: st.c0, cn: st.cn });
      $("sheet-info").innerHTML = "<b>" + esc(d.label) + "</b> · 단위: <b>" + esc(d.unit) + "</b> · 전체 " + nf(d.nrows) + "행 × " + nf(d.ncols) + "열 · " + esc(d.desc);
      var heads = [{ t: "코드", l: 1 }, { t: "산업명", l: 1 }].concat(d.cols.map(function (c) { return { t: esc(c.c) + '<br><small class="muted">' + esc(c.n.length > 14 ? c.n.slice(0, 14) + "…" : c.n) + "</small>" }; }));
      var dec = d.unit === "백만원" ? 0 : 5;
      var rows = d.rows.map(function (r) { return [esc(r.c), esc(r.n)].concat(r.v.map(function (v) { return nf(v, dec); })); });
      $("sheet-table").innerHTML = table(heads, rows) + (d.truncated ? '<p class="muted" style="padding:8px">행이 많아 앞의 ' + d.rows.length + "행만 표시했습니다. 행 검색을 사용하세요.</p>" : "");
      var last = Math.min(d.c0 + d.cols.length, d.mcols);
      $("sheet-pos").textContent = d.mcols ? "열 " + nf(d.c0 + 1) + "–" + nf(last) + " / " + nf(d.mcols) + " (일치 열 기준)" : "일치하는 열이 없습니다";
      $("sheet-prev").disabled = d.c0 <= 0;
      $("sheet-next").disabled = d.c0 + st.cn >= d.mcols;
    } catch (err) { showErr($("sheet-table"), err); }
  }

  // ---------------------------------------------------------------- 산업 검색
  function resultCard(x) {
    return '<button class="res" data-c="' + esc(x.c) + '"><b>[' + esc(x.c) + "] " + esc(x.n) + "</b>" +
      (x.direct ? '<small class="tag-direct">이름·코드 일치</small>' : "") +
      (x.via && x.via.length ? "<small>다른 분류 일치: " + esc(x.via.join(" / ")) + "</small>" : "") + "</button>";
  }
  function renderSearchList() {
    $("s-count").textContent = "전체 " + state.inds.length + "개 산업 중 앞의 80개입니다. 검색어를 입력해 좁혀 보세요.";
    $("s-results").innerHTML = state.inds.slice(0, 80).map(function (x) { return resultCard({ c: x.c, n: x.n, direct: false, via: [] }); }).join("");
  }
  async function doSearch() {
    var q = $("s-q").value.trim();
    if (!q) { renderSearchList(); return; }
    try {
      var d = await call("search", { q: q });
      $("s-count").textContent = "검색 결과 " + d.total + "건" + (d.total > d.items.length ? " (앞의 " + d.items.length + "건 표시)" : "");
      $("s-results").innerHTML = d.items.length ? d.items.map(resultCard).join("") : '<p class="muted">일치하는 산업이 없습니다.</p>';
    } catch (err) { showErr($("s-results"), err); }
  }
  async function showProfile(code) {
    $("s-profile").innerHTML = '<p class="muted">분석 중…</p>';
    try {
      var p = await call("profile", { code: code });
      $("s-profile").innerHTML = profileHtml(p);
      var el = $("s-profile");
      if (el.scrollIntoView) el.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (err) { showErr($("s-profile"), err); }
  }
  function profileHtml(p) {
    var m = p.metrics, s = p.supply, d = p.demand, i = p.input;
    var h = '<div class="prof"><hr><h3>[' + esc(p.c) + "] " + esc(p.n) + '<span class="badge">' + esc(p.level) + '</span><span class="badge">' + esc(p.type) + "</span></h3>";
    if (p.zero) h += '<p class="note">이 부문은 2023년 총산출이 0이라 비율·계수 지표를 계산하지 않습니다.</p>';
    h += '<h3>분류 연결</h3><div class="hier">' + p.hier.map(function (x) {
      var list = x.items.slice(0, 12).map(function (y) { return esc(y.c) + " " + esc(y.n); }).join("<br>");
      return "<div><b>" + esc(x.level) + " · " + x.relation + " (" + x.count + "개)</b>" + (list || "-") + (x.count > 12 ? "<br>…" : "") + "</div>";
    }).join("") + "</div>";
    var kp = [
      ["총산출", human(s.out)], ["부가가치율", pct(m.va_ratio)], ["중간투입률", pct(m.mid_ratio)], ["수입의존도", pct(m.imp_dep)],
      ["수출률", pct(m.export_rate)], ["수입률", pct(m.import_rate)], ["생산유발계수", nf(m.prod, 3)], ["부가가치유발계수", nf(m.va_mult, 3)],
      ["수입유발계수", nf(m.im_mult, 3)], ["영향력계수 (평균=1)", nf(m.infl, 3)], ["감응도계수 (평균=1)", nf(m.sens, 3)], ["총산출 비중", pct(m.out_share, 2)]
    ];
    h += '<h3>핵심 지표</h3><div class="kpis">' + kp.map(function (x) { return '<div class="kpi"><b>' + x[1] + "</b><span>" + x[0] + "</span></div>"; }).join("") + "</div>";
    h += '<div class="grid2"><div><h3>수요 구조 <small class="muted">(총수요 ' + human(d.dem) + " 대비 %)</small></h3>" +
      bars([{ label: "중간수요", v: d.mid_dem, text: human(d.mid_dem) + " (" + pct(d.dem ? 100 * d.mid_dem / d.dem : null) + ")" }].concat(d.items.map(function (x) {
        return { label: x.n, v: x.v, text: human(x.v) + " (" + pct(x.share) + ")" };
      }))) + "</div><div><h3>투입 구조 <small class=\"muted\">(총투입 " + human(i.tin) + " 대비 %)</small></h3>" +
      bars([{ label: "중간투입", v: i.mid_in, text: human(i.mid_in) + " (" + pct(i.tin ? 100 * i.mid_in / i.tin : null) + ")", cls: "g" }].concat(i.items.map(function (x) {
        return { label: x.n, v: x.v, text: human(x.v) + " (" + pct(x.share) + ")", cls: "g" };
      }))) + "</div></div>";
    var topRows = function (list) { return list.map(function (x) { return { label: "[" + x.c + "] " + x.n, v: x.v, text: human(x.v) + " (" + pct(x.r) + ")", cls: "o" }; }); };
    h += '<div class="grid2"><div><h3>주요 투입 산업 <small class="muted">(총투입 대비 %)</small></h3>' + (p.top_inputs.length ? bars(topRows(p.top_inputs)) : '<p class="muted">-</p>') +
      '</div><div><h3>주요 판매 산업 <small class="muted">(중간수요 대비 %)</small></h3>' + (p.top_buyers.length ? bars(topRows(p.top_buyers)) : '<p class="muted">-</p>') + "</div></div>";
    h += '<h3>공급 구조</h3>' + table([{ t: "항목", l: 1 }, { t: "금액(백만원)" }, { t: "비고" }], [
      ["총산출", nf(s.out), human(s.out)], ["수입", nf(s.imp), human(s.imp)], ["생산물세(수입)", nf(s.impt), human(s.impt)],
      ["수입계", nf(s.imp_all), human(s.imp_all)], ["총공급계", nf(s.sup), human(s.sup)]]);
    return h + "</div>";
  }

  // ---------------------------------------------------------------- 지표·순위
  function renderRankSetup() {
    $("r-metric").innerHTML = state.metrics.meta.map(function (x) { return '<option value="' + x.key + '">' + esc(x.label) + " (" + esc(x.unit) + ")</option>"; }).join("");
    $("r-metric").value = "prod";
    $("r-q").value = "";
    $("r-order").value = "desc";
  }
  function metricMeta(key) { return state.metrics.meta.filter(function (x) { return x.key === key; })[0]; }
  function fmtMetric(key, v) { return v == null ? "-" : nf(v, FMT[key] == null ? 2 : FMT[key]); }
  function renderRank() {
    var key = $("r-metric").value || "prod";
    var meta = metricMeta(key);
    var order = $("r-order").value === "asc" ? 1 : -1;
    var topN = Number($("r-top").value) || 15;
    $("r-desc").innerHTML = "<b>" + esc(meta.label) + "</b> (" + esc(meta.unit) + ") — " + esc(meta.desc);
    var rows = state.metrics.rows.filter(function (r) { return r[key] != null; });
    rows.sort(function (a, b) { return order * (a[key] - b[key]); });
    $("r-chart").innerHTML = bars(rows.slice(0, topN).map(function (r) { return { label: "[" + r.c + "] " + r.n, v: r[key], text: fmtMetric(key, r[key]) }; }));
    // 산업 유형 칩
    var tc = state.metrics.type_count, types = Object.keys(tc).sort();
    $("r-types").innerHTML = '<span class="chip tag' + (state.rank.type === "" ? " on" : "") + '" data-t="">전체 ' + state.metrics.rows.length + "</span>" +
      types.map(function (t) { return '<span class="chip tag' + (state.rank.type === t ? " on" : "") + '" data-t="' + esc(t) + '">' + esc(t) + " " + tc[t] + "</span>"; }).join("");
    renderRankTable();
  }
  function rankRows() {
    var q = $("r-q").value.trim().toLowerCase(), t = state.rank.type, k = state.rank.sortKey, dir = state.rank.sortDir;
    var rows = state.metrics.rows.filter(function (r) {
      return (!t || r.type === t) && (!q || r.c.toLowerCase().indexOf(q) >= 0 || r.n.toLowerCase().indexOf(q) >= 0);
    });
    rows.sort(function (a, b) {
      var x = a[k], y = b[k];
      if (typeof x === "string" || typeof y === "string") return dir * String(x).localeCompare(String(y), "ko");
      if (x == null) return 1;
      if (y == null) return -1;
      return dir * (x - y);
    });
    return rows;
  }
  function renderRankTable() {
    var rows = rankRows();
    var heads = [{ t: "코드", l: 1, key: "c" }, { t: "산업명", l: 1, key: "n" }, { t: "유형", l: 1, key: "type" }].concat(state.metrics.meta.map(function (m) {
      return { t: esc(m.label) + '<br><small class="muted">' + esc(m.unit) + "</small>" + (state.rank.sortKey === m.key ? (state.rank.sortDir > 0 ? " ▲" : " ▼") : ""), key: m.key };
    }));
    var body = rows.map(function (r) {
      return [esc(r.c), esc(r.n), esc(r.type)].concat(state.metrics.meta.map(function (m) { return fmtMetric(m.key, r[m.key]); }));
    });
    $("r-table").innerHTML = table(heads, body, { sortable: true }) + '<p class="muted" style="padding:6px 10px">' + rows.length + "개 산업</p>";
  }
  function downloadCsv() {
    var rows = rankRows();
    var head = ["코드", "산업명", "유형"].concat(state.metrics.meta.map(function (m) { return m.label + "(" + m.unit + ")"; }));
    var lines = [head.join(",")].concat(rows.map(function (r) {
      return [r.c, '"' + String(r.n).replace(/"/g, '""') + '"', '"' + r.type + '"'].concat(state.metrics.meta.map(function (m) { return r[m.key] == null ? "" : r[m.key]; })).join(",");
    }));
    var blob = new Blob(["﻿" + lines.join("\n")], { type: "text/csv;charset=utf-8" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "산업별지표_" + state.info.level + ".csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  // ---------------------------------------------------------------- 최종수요 유발효과
  async function renderFd() {
    try {
      var d = await call("fdview", { item: state.fdItem, top: 15 });
      var heads = [{ t: "최종수요 항목", l: 1 }, { t: "생산유발액<br>(조원)" }, { t: "구성비(%)" }, { t: "생산유발계수" },
        { t: "부가가치유발액<br>(조원)" }, { t: "구성비(%)" }, { t: "부가가치유발계수" }, { t: "수입유발액<br>(조원)" }, { t: "구성비(%)" }, { t: "수입유발계수" }];
      $("fd-table").innerHTML = table(heads, d.table.map(function (r) {
        return [esc(r.n), jo(r.out, 1), nf(r.out_share, 1), nf(r.out_mult, 3), jo(r.va, 1), nf(r.va_share, 1), nf(r.va_mult, 3), jo(r.im, 1), nf(r.im_share, 1), nf(r.im_mult, 3)];
      }));
      $("fd-item").innerHTML = d.table.map(function (r) { return '<option value="' + r.c + '">' + esc(r.n) + "</option>"; }).join("");
      $("fd-item").value = d.item;
      $("fd-title").textContent = d.item_name + " 때문에 생산이 많이 유발되는 산업";
      $("fd-chart").innerHTML = bars(d.top.map(function (r) { return { label: "[" + r.c + "] " + r.n, v: r.out, text: jo(r.out, 1) + "조원" }; }));
      $("fd-top").innerHTML = table([{ t: "코드", l: 1 }, { t: "산업명", l: 1 }, { t: "생산유발액(억원)" }, { t: "부가가치유발액(억원)" }, { t: "수입유발액(억원)" }],
        d.top.map(function (r) { return [esc(r.c), esc(r.n), eok(r.out), eok(r.va), eok(r.im)]; }));
    } catch (err) { showErr($("fd-table"), err); }
  }

  // ---------------------------------------------------------------- 시뮬레이터
  function renderSimSetup() {
    $("ind-list").innerHTML = state.inds.map(function (x) { return '<option value="' + esc(x.c + " " + x.n) + '"></option>'; }).join("");
    renderSimRows();
  }
  function renderSimRows() {
    $("sim-rows").innerHTML = state.sim.map(function (r, i) {
      return '<div class="sim-row" data-i="' + i + '"><input class="ind" list="ind-list" placeholder="산업 코드·이름 (목록에서 선택)" value="' + esc(r.label) + '">' +
        '<input class="amt" type="number" step="any" placeholder="증감 금액" value="' + esc(r.amount) + '">' +
        '<select class="unit">' + ["억원", "조원", "백만원"].map(function (u) { return '<option' + (u === r.unit ? " selected" : "") + ">" + u + "</option>"; }).join("") + "</select>" +
        '<button class="del" title="삭제">✕</button></div>';
    }).join("");
  }
  function readSimRows() {
    var rows = document.querySelectorAll("#sim-rows .sim-row");
    for (var i = 0; i < rows.length; i++) {
      state.sim[i].label = rows[i].querySelector("input.ind").value;
      state.sim[i].amount = rows[i].querySelector("input.amt").value;
      state.sim[i].unit = rows[i].querySelector("select.unit").value;
    }
  }
  function resolveCode(label) {
    label = (label || "").trim();
    if (state.labelToCode[label]) return state.labelToCode[label];
    var hit = state.inds.filter(function (x) { return x.c === label || x.n === label; });
    return hit.length === 1 ? hit[0].c : null;
  }
  async function runSim() {
    readSimRows();
    var FACT = { "백만원": 1, "억원": 100, "조원": 1e6 };
    var shocks = [];
    for (var i = 0; i < state.sim.length; i++) {
      var r = state.sim[i];
      if (!r.label && !r.amount) continue;
      var code = resolveCode(r.label), amt = Number(r.amount);
      if (!code) { $("sim-msg").textContent = (i + 1) + "번째 줄의 산업을 목록에서 선택하세요."; return; }
      if (!isFinite(amt) || r.amount === "") { $("sim-msg").textContent = (i + 1) + "번째 줄의 금액을 입력하세요."; return; }
      shocks.push({ code: code, amount: amt * FACT[r.unit] });
    }
    if (!shocks.length) { $("sim-msg").textContent = "산업과 금액을 입력하세요."; return; }
    $("sim-msg").textContent = "계산 중…";
    $("sim-run").disabled = true;
    try {
      var d = await call("simulate", { shocks: shocks });
      $("sim-msg").textContent = "";
      $("sim-result").innerHTML = simHtml(d);
    } catch (err) { $("sim-msg").textContent = ""; showErr($("sim-result"), err); }
    $("sim-run").disabled = false;
  }
  function simHtml(d) {
    var k = [
      ["입력한 최종수요", human(d.final_demand)], ["유발 생산액", human(d.dx) + " (×" + nf(d.mult_prod, 3) + ")"],
      ["유발 부가가치", human(d.dva) + " (×" + nf(d.mult_va, 3) + ")"], ["유발 수입", human(d.dim) + " (×" + nf(d.mult_im, 3) + ")"]
    ];
    var h = '<div class="kpis">' + k.map(function (x) { return '<div class="kpi"><b>' + x[1] + "</b><span>" + x[0] + "</span></div>"; }).join("") + "</div>";
    h += '<p class="muted">검산: 부가가치 + 수입 − 최종수요 = ' + nf(d.closure, 4) + " 백만원 (0이면 균형). 괄호 안은 최종수요 1단위당 배수입니다.</p>";
    h += "<h3>부가가치 유발액의 구성 <small class=\"muted\">(산업별 부가가치 항목 비율을 적용한 추정)</small></h3>" +
      bars(d.comp.map(function (c) { return { label: c.n, v: c.v, text: human(c.v), cls: "g" }; }));
    h += "<h3>생산유발액 상위 산업 <small class=\"muted\">(억원)</small></h3>" +
      bars(d.rows.slice(0, 10).map(function (r) { return { label: "[" + r.c + "] " + r.n, v: r.dx, text: eok(r.dx) + "억원" }; }));
    h += "<h3>산업별 결과 <small class=\"muted\">(단위: 억원, 총 " + d.nrows + "개 산업" + (d.nrows > d.rows.length ? ", 앞의 " + d.rows.length + "개 표시" : "") + ")</small></h3>" +
      '<div class="scroll tall">' + table(
        [{ t: "코드", l: 1 }, { t: "산업명", l: 1 }, { t: "직접(입력)" }, { t: "간접효과" }, { t: "생산유발액" }, { t: "비중(%)" }, { t: "부가가치유발" }, { t: "수입유발" }],
        d.rows.map(function (r) { return [esc(r.c), esc(r.n), eok(r.direct), eok(r.indirect), eok(r.dx), nf(r.share, 1), eok(r.dva), eok(r.dim)]; })) + "</div>";
    return h;
  }

  // ---------------------------------------------------------------- 이벤트 연결
  function bind() {
    $("btn-load").addEventListener("click", loadData);
    $("tabs").addEventListener("click", function (e) {
      var t = e.target && e.target.getAttribute && e.target.getAttribute("data-tab");
      if (t) showTab(t);
    });
    $("btn-deep").addEventListener("click", runDeep);
    $("sheet-sel").addEventListener("change", function () { state.sheet.key = $("sheet-sel").value; state.sheet.c0 = 0; loadSheet(); });
    var resetSheet = function () { state.sheet.c0 = 0; loadSheet(); };
    $("sheet-rq").addEventListener("change", resetSheet);
    $("sheet-cq").addEventListener("change", resetSheet);
    $("sheet-prev").addEventListener("click", function () { state.sheet.c0 = Math.max(0, state.sheet.c0 - state.sheet.cn); loadSheet(); });
    $("sheet-next").addEventListener("click", function () { state.sheet.c0 += state.sheet.cn; loadSheet(); });
    $("s-btn").addEventListener("click", doSearch);
    $("s-q").addEventListener("keydown", function (e) { if (e.key === "Enter") doSearch(); });
    $("s-results").addEventListener("click", function (e) {
      var el = e.target && e.target.closest ? e.target.closest("button.res") : null;
      if (el) showProfile(el.getAttribute("data-c"));
    });
    ["r-metric", "r-order", "r-top"].forEach(function (id) { $(id).addEventListener("change", renderRank); });
    $("r-q").addEventListener("input", renderRankTable);
    $("r-types").addEventListener("click", function (e) {
      var t = e.target && e.target.getAttribute ? e.target.getAttribute("data-t") : null;
      if (t == null) return;
      state.rank.type = t;
      renderRank();
    });
    $("r-table").addEventListener("click", function (e) {
      var th = e.target && e.target.closest ? e.target.closest("th") : null;
      var k = th && th.getAttribute("data-k");
      if (!k) return;
      if (state.rank.sortKey === k) state.rank.sortDir = -state.rank.sortDir;
      else { state.rank.sortKey = k; state.rank.sortDir = (k === "c" || k === "n" || k === "type") ? 1 : -1; }
      renderRankTable();
    });
    $("r-csv").addEventListener("click", downloadCsv);
    $("fd-item").addEventListener("change", function () { state.fdItem = $("fd-item").value; renderFd(); });
    $("sim-add").addEventListener("click", function () { readSimRows(); state.sim.push({ code: "", label: "", amount: "", unit: "억원" }); renderSimRows(); });
    $("sim-run").addEventListener("click", runSim);
    $("sim-rows").addEventListener("click", function (e) {
      var b = e.target && e.target.closest ? e.target.closest("button.del") : null;
      if (!b) return;
      readSimRows();
      var i = Number(b.parentNode.getAttribute("data-i"));
      state.sim.splice(i, 1);
      if (!state.sim.length) state.sim.push({ code: "", label: "", amount: "", unit: "억원" });
      renderSimRows();
    });
  }

  bind();
  startWorker();
})();
