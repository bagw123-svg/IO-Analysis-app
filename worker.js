/* worker.js - 화면(app.js)과 별도의 작업 공간(Web Worker)에서 Pyodide로 Python(model.py)을 실행한다.
 * 화면 -> 작업자: {id, cmd, args}
 * 작업자 -> 화면: {type:'status'|'ready'|'fail'|'progress'|'result', ...}
 */
const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
importScripts(PYODIDE_URL + "pyodide.js");

let pyodide = null;
let model = null;
let queue = Promise.resolve();
const written = {}; // 이미 가상 파일시스템에 쓴 파일

function post(type, payload) {
  self.postMessage(Object.assign({ type: type }, payload || {}));
}

async function init() {
  try {
    post("status", { text: "Python(Pyodide) 내려받는 중…" });
    pyodide = await loadPyodide({ indexURL: PYODIDE_URL });
    post("status", { text: "numpy 준비 중…" });
    await pyodide.loadPackage(["numpy"]);
    post("status", { text: "model.py 불러오는 중…" });
    const res = await fetch("model.py?v=2");
    if (!res.ok) throw new Error("model.py를 읽지 못했습니다 (" + res.status + ")");
    pyodide.FS.writeFile("/home/pyodide/model.py", await res.text());
    pyodide.runPython("import sys\nif '/home/pyodide' not in sys.path: sys.path.insert(0, '/home/pyodide')");
    model = pyodide.pyimport("model");
    model.set_progress(pyodide.ffi.create_proxy(function (msg) { post("progress", { text: String(msg) }); }));
    post("ready", {});
  } catch (e) {
    post("fail", { text: String(e && e.message ? e.message : e) });
  }
}

/* 파일을 내려받아 Pyodide 가상 파일시스템에 쓴다 (진행률 전달) */
async function stage(url, dest, label) {
  if (written[dest]) return;
  const res = await fetch(url);
  if (!res.ok) throw new Error(label + " 파일을 읽지 못했습니다 (" + res.status + ")");
  const total = Number(res.headers.get("Content-Length")) || 0;
  const reader = res.body && res.body.getReader ? res.body.getReader() : null;
  let bytes;
  if (reader) {
    const chunks = [];
    let got = 0;
    for (;;) {
      const r = await reader.read();
      if (r.done) break;
      chunks.push(r.value);
      got += r.value.length;
      const mb = (got / 1048576).toFixed(1);
      post("progress", { text: label + " 내려받는 중… " + mb + "MB" + (total ? " / " + (total / 1048576).toFixed(1) + "MB" : ""), pct: total ? Math.round((100 * got) / total) : null });
    }
    bytes = new Uint8Array(got);
    let off = 0;
    for (const c of chunks) { bytes.set(c, off); off += c.length; }
  } else {
    bytes = new Uint8Array(await res.arrayBuffer());
  }
  pyodide.FS.writeFile(dest, bytes);
  written[dest] = true;
}

async function run(msg) {
  const id = msg.id, cmd = msg.cmd;
  let args = msg.args || {};
  try {
    if (!model) throw new Error("Python 엔진이 아직 준비되지 않았습니다.");
    if (cmd === "load") {
      const lv = args.level;
      await stage("data/product-class.xlsx", "/home/pyodide/product-class.xlsx", "상품분류표");
      await stage("data/" + args.file, "/home/pyodide/io-" + lv + ".xlsx", args.label + " 투입산출표");
      args = { level: lv, path: "/home/pyodide/io-" + lv + ".xlsx", cls_path: "/home/pyodide/product-class.xlsx" };
      post("progress", { text: "Python이 Excel을 읽는 중입니다…", pct: null });
    }
    const json = model.handle(cmd, JSON.stringify(args));
    post("result", { id: id, json: json });
  } catch (e) {
    post("result", { id: id, json: JSON.stringify({ ok: false, error: String(e && e.message ? e.message : e) }) });
  }
}

const ready = init();

self.onmessage = function (e) {
  queue = queue.then(function () { return ready; }).then(function () { return run(e.data); });
};
