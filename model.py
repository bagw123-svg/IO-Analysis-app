"""model.py - 산업연관표(투입산출표) Excel 읽기, 검증, 산업연관분석 계산.

브라우저의 Pyodide(WebAssembly)에서 실행되며 numpy와 Python 표준 라이브러리(zipfile, xml)만 사용한다.
Excel(.xlsx)은 추가 패키지 없이 직접 읽는다.
worker.js 가 handle(cmd, args_json) 을 호출하고, 결과는 JSON 문자열로 돌려준다.

단위 규칙
  - 거래액·유발액은 모두 '백만원' (2023년 당해년 가격, 생산자가격)
  - 계수(투입계수·유발계수)는 무차원 (생산 또는 최종수요 1단위당 크기)
  - 비율 지표는 % , 영향력·감응도계수는 전산업 평균을 1.0 으로 한 지수
"""
import json
import math
import re
import traceback
import xml.etree.ElementTree as ET
import zipfile

import numpy as np

LEVELS = {"large": "대분류", "mid": "중분류", "small": "소분류", "basic": "기본부문"}
LEVEL_ORDER = ["basic", "small", "mid", "large"]  # 상품분류표 열 순서(세분 -> 광범위)

# key, Excel 시트명, 이름, 단위, 설명
SHEETS = [
    ("total", "A표_총거래표(생산자)", "총거래표", "백만원",
     "산업간 거래(중간수요) + 최종수요 + 총산출·수입 + 부가가치. 모든 분석의 기초 자료"),
    ("import", "A표_수입거래표(생산자)", "수입거래표", "백만원", "총거래표 중 수입품으로 공급된 거래"),
    ("domestic", "A표_국산거래표(생산자)", "국산거래표", "백만원", "총거래표 중 국산품으로 공급된 거래"),
    ("A", "총투입계수(A)", "총투입계수표", "무차원", "a_ij = 거래액 x_ij ÷ 총투입 X_j (산업 j 생산 1단위당 투입)"),
    ("Am", "수입투입계수(Am)", "수입투입계수표", "무차원", "수입 중간재 투입계수"),
    ("Ad", "국산투입계수(Ad)", "국산투입계수표", "무차원", "국산 중간재 투입계수 (Ad = A - Am)"),
    ("B", "생산유발계수", "생산유발계수표", "무차원",
     "B = (I - Ad)^-1. 최종수요 1단위가 전 산업에 유발하는 생산액. 열합=영향력, 행합=감응도"),
    ("M", "수입유발계수", "수입유발계수표", "무차원", "M = Am · B. 최종수요 1단위가 유발하는 수입(중간재)액"),
    ("V", "부가가치유발계수", "부가가치유발계수표", "무차원", "V = v̂ · B (v = 부가가치율). 최종수요 1단위가 유발하는 부가가치"),
    ("fd_out", "최종수요 항목별 생산유발액", "최종수요 항목별 생산유발액", "백만원", "소비·투자·수출 항목별 생산유발액"),
    ("fd_imp", "최종수요 항목별 수입유발액", "최종수요 항목별 수입유발액", "백만원", "소비·투자·수출 항목별 수입유발액"),
    ("fd_va", "최종수요 항목별 부가가치유발액", "최종수요 항목별 부가가치유발액", "백만원", "소비·투자·수출 항목별 부가가치유발액"),
]
SHEET_BY_KEY = {s[0]: s for s in SHEETS}

FD_ITEMS = [("9111", "민간소비지출"), ("9112", "정부소비지출"), ("9121", "민간고정자본형성"),
            ("9122", "정부고정자본형성"), ("9131", "재고증감"), ("9132", "귀중품 순취득"), ("9140", "수출")]
VA_ITEMS = [("9610", "피용자보수"), ("9621", "영업잉여"), ("9622", "고정자본소모"), ("9630", "생산세(보조금 공제)")]

# 한국은행 2025.9.24 보도자료(2023년 산업연관표 연장표) 수록값, 조원
PRESS = [
    ("총산출(국내 공급)", 5786.7), ("수입(해외 공급)", 1016.0), ("총공급 = 총수요", 6802.7),
    ("중간투입(= 중간수요)", 3403.4), ("부가가치", 2383.3), ("  피용자보수", 1170.4), ("  영업잉여", 495.9),
    ("  고정자본소모", 539.2), ("  순생산세", 177.8), ("소비", 1631.3), ("투자", 768.2), ("수출", 999.8),
]

S = {}
_PROGRESS = None


# ----------------------------------------------------------------------------- 공통 도우미
def set_progress(fn):
    global _PROGRESS
    _PROGRESS = fn


def _log(msg):
    if _PROGRESS is not None:
        try:
            _PROGRESS(msg)
        except Exception:
            pass


def _clean(s):
    if s is None:
        return ""
    return " ".join(str(s).replace("　", " ").split())


def _num(x):
    if isinstance(x, bool) or x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    return 0.0


def _sd(a, b):
    """a/b, b==0 이면 nan"""
    a = np.asarray(a, float)
    b = np.asarray(b, float)
    out = np.full(a.shape, np.nan)
    m = b != 0
    out[m] = a[m] / b[m]
    return out


def _f(x, nd=6):
    try:
        x = float(x)
    except Exception:
        return None
    if not math.isfinite(x):
        return None
    return round(x, nd)


def _arr(a, nd=6):
    return [_f(v, nd) for v in np.asarray(a, float).ravel().tolist()]


# ----------------------------------------------------------------------------- 최소 xlsx 읽기 (표준 라이브러리만 사용)
_M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _col_index(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def _to_num(text):
    t = text.strip()
    if "." in t or "E" in t or "e" in t:
        return float(t)
    return int(t)


class Book:
    """.xlsx 파일(zip + XML)에서 시트 이름과 셀 값을 읽는다. 값은 수식의 저장된 결과값이다."""

    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        wb = ET.fromstring(self.z.read("xl/workbook.xml"))
        rels = ET.fromstring(self.z.read("xl/_rels/workbook.xml.rels"))
        target = {r.get("Id"): r.get("Target") for r in rels}
        self.sheets = {}
        for sh in wb.iter(_M + "sheet"):
            t = target[sh.get(_R + "id")].lstrip("/")
            self.sheets[sh.get("name")] = t if t.startswith("xl/") else "xl/" + t
        self.sheetnames = list(self.sheets)
        self._sst = None

    def close(self):
        try:
            self.z.close()
        except Exception:
            pass

    def _shared(self):
        if self._sst is None:
            out = []
            if "xl/sharedStrings.xml" in self.z.namelist():
                with self.z.open("xl/sharedStrings.xml") as f:
                    for _, el in ET.iterparse(f, events=("end",)):
                        if el.tag == _M + "si":
                            parts = []
                            for ch in el:
                                if ch.tag == _M + "t":
                                    parts.append(ch.text or "")
                                elif ch.tag == _M + "r":
                                    parts.extend((x.text or "") for x in ch.findall(_M + "t"))
                            out.append("".join(parts))
                            el.clear()
            self._sst = out
        return self._sst

    def dims(self, name):
        """(행 수, 열 수) - 시트 머리의 dimension 정보만 읽는다."""
        with self.z.open(self.sheets[name]) as f:
            head = f.read(4096).decode("utf-8", "ignore")
        m = re.search(r'<dimension ref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"', head)
        if not m:
            return None, None
        if m.group(3):
            return int(m.group(4)), _col_index(m.group(3)) + 1
        return int(m.group(2)), _col_index(m.group(1)) + 1

    def rows(self, name):
        """시트 전체를 행 목록(list of list)으로 돌려준다. 빈 칸은 None, 행·열 번호는 엑셀과 같다."""
        sst = self._shared()
        out = []
        with self.z.open(self.sheets[name]) as f:
            for _, el in ET.iterparse(f, events=("end",)):
                if el.tag != _M + "row":
                    continue
                rn = int(el.get("r") or (len(out) + 1))
                while len(out) < rn - 1:
                    out.append([])
                vals = []
                for c in el.findall(_M + "c"):
                    ref = c.get("r")
                    ci = _col_index(ref) if ref else len(vals)
                    while len(vals) <= ci:
                        vals.append(None)
                    t = c.get("t")
                    v = c.find(_M + "v")
                    if t == "inlineStr":
                        node = c.find(_M + "is")
                        val = "".join((x.text or "") for x in node.iter(_M + "t")) if node is not None else None
                    elif v is None or v.text is None:
                        val = None
                    elif t == "s":
                        val = sst[int(v.text)]
                    elif t in ("str", "e"):
                        val = v.text
                    elif t == "b":
                        val = v.text.strip() == "1"
                    else:
                        val = _to_num(v.text)
                    vals[ci] = val
                out.append(vals)
                el.clear()
        return out


# ----------------------------------------------------------------------------- Excel 읽기
def _read_grid(sheet_name):
    """투입산출표 시트 한 장을 읽어 {cc, cn, rc, rn, M, ci, ri} 로 돌려준다."""
    rows = S["wb"].rows(sheet_name)
    hi = None
    for i, r in enumerate(rows[:12]):
        if r and r[0] is not None and "상품" in str(r[0]):
            hi = i
            break
    if hi is None:
        raise ValueError("시트 머리글(상품) 행을 찾지 못했습니다: " + sheet_name)
    hdr, nm = rows[hi], rows[hi + 1]
    cols = [j for j in range(2, len(hdr)) if hdr[j] not in (None, "")]
    cc = [str(hdr[j]).strip() for j in cols]
    cn = [_clean(nm[j]) if j < len(nm) else "" for j in cols]
    body = [r for r in rows[hi + 2:] if r and r[0] not in (None, "")]
    rc = [str(r[0]).strip() for r in body]
    rn = [_clean(r[1]) if len(r) > 1 else "" for r in body]
    M = np.array([[_num(r[j]) if j < len(r) else 0.0 for j in cols] for r in body], dtype=float)
    if M.size == 0:
        M = np.zeros((len(body), len(cols)))
    ci, ri = {}, {}
    for k, c in enumerate(cc):
        ci.setdefault(c, k)
    for k, c in enumerate(rc):
        ri.setdefault(c, k)
    return {"cc": cc, "cn": cn, "rc": rc, "rn": rn, "M": M, "ci": ci, "ri": ri}


def G(key):
    """시트를 한 번만 읽어 캐시"""
    if key not in S["grids"]:
        _, name, label, _, _ = SHEET_BY_KEY[key]
        if name not in S["wb"].sheetnames:
            raise KeyError("Excel에 '%s' 시트가 없습니다." % name)
        _log("시트 읽는 중: %s" % label)
        S["grids"][key] = _read_grid(name)
    return S["grids"][key]


def _sub(g, rows, cols):
    """산업 코드 목록으로 부분행렬을 뽑는다(없는 코드는 0)."""
    ri = [g["ri"].get(r, -1) for r in rows]
    ci = [g["ci"].get(c, -1) for c in cols]
    out = np.zeros((len(rows), len(cols)))
    rr = [k for k, v in enumerate(ri) if v >= 0]
    cc = [k for k, v in enumerate(ci) if v >= 0]
    if rr and cc:
        out[np.ix_(rr, cc)] = g["M"][np.ix_([ri[k] for k in rr], [ci[k] for k in cc])]
    return out


def _colv(g, code, rows):
    """총거래표 등에서 특정 열(code) 값을 산업 행 순서로"""
    j = g["ci"].get(code)
    if j is None:
        return np.zeros(len(rows))
    return np.array([g["M"][g["ri"][r], j] if r in g["ri"] else 0.0 for r in rows])


def _rowv(g, code, cols):
    """총거래표 등에서 특정 행(code) 값을 산업 열 순서로"""
    i = g["ri"].get(code)
    if i is None:
        return np.zeros(len(cols))
    return np.array([g["M"][i, g["ci"][c]] if c in g["ci"] else 0.0 for c in cols])


# ----------------------------------------------------------------------------- 상품분류표
def _read_classification(path):
    bk = Book(path)
    rows = bk.rows(bk.sheetnames[0])
    bk.close()
    recs = []
    cur = {"small": None, "mid": None, "large": None}
    for r in rows[3:]:
        if not r or r[0] in (None, ""):
            continue
        code = str(r[0]).strip()
        if code == "9090":
            break
        if len(r) > 2 and r[2]:
            cur["small"] = (str(r[2]).strip(), _clean(r[3]))
        if len(r) > 4 and r[4]:
            cur["mid"] = (str(r[4]).strip(), _clean(r[5]))
        if len(r) > 6 and r[6]:
            cur["large"] = (str(r[6]).strip(), _clean(r[7]))
        recs.append({"basic": (code, _clean(r[1])), "small": cur["small"], "mid": cur["mid"], "large": cur["large"]})
    ent = {lv: {} for lv in LEVEL_ORDER}
    mem = {lv: {} for lv in LEVEL_ORDER}
    for k, rec in enumerate(recs):
        for lv in LEVEL_ORDER:
            if rec[lv]:
                ent[lv].setdefault(rec[lv][0], rec[lv][1])
                mem[lv].setdefault(rec[lv][0], []).append(k)
    S["cls"] = {"recs": recs, "ent": ent, "mem": mem}


def _hier(code):
    """현재 분류의 산업 code 에 대응하는 다른 분류(상위/하위) 목록"""
    cur = S["level"]
    cls = S["cls"]
    out = []
    members = cls["mem"][cur].get(code, [])
    for lv in LEVEL_ORDER[::-1]:  # 대 -> 기본
        if lv == cur:
            continue
        seen, items = set(), []
        for k in members:
            e = cls["recs"][k][lv]
            if e and e[0] not in seen:
                seen.add(e[0])
                items.append({"c": e[0], "n": e[1]})
        out.append({"level": LEVELS[lv], "relation": "상위" if LEVEL_ORDER.index(lv) > LEVEL_ORDER.index(cur) else "하위",
                    "count": len(items), "items": items[:40]})
    return out


# ----------------------------------------------------------------------------- 불러오기
def load(a):
    level, path, cls_path = a["level"], a["path"], a["cls_path"]
    old = S.get("wb")
    S.clear()
    try:
        if old is not None:
            old.close()
    except Exception:
        pass
    S["level"] = level
    _log("Excel 파일 여는 중")
    S["wb"] = Book(path)
    S["grids"] = {}
    S["cache"] = {}
    _log("상품분류표 읽는 중")
    _read_classification(cls_path)

    g = G("total")
    allc = g["cc"]
    if "9090" not in allc:
        raise ValueError("총거래표에서 중간수요계(9090) 열을 찾지 못했습니다.")
    ind = allc[:allc.index("9090")]
    S["ind"] = ind
    S["n"] = len(ind)
    S["ix"] = {c: i for i, c in enumerate(ind)}
    S["names"] = [g["rn"][g["ri"][c]] if c in g["ri"] else g["cn"][g["ci"][c]] for c in ind]
    S["X"] = _sub(g, ind, ind)
    # 총거래표의 열(최종수요·공급)과 행(투입)
    S["col"] = {k: _colv(g, k, ind) for k in
                ["9090", "9111", "9112", "9121", "9122", "9131", "9132", "9140", "9190", "9290", "9310",
                 "9321", "9322", "9329", "9331", "9332", "9390"]}
    S["row"] = {k: _rowv(g, k, ind) for k in ["9590", "9610", "9621", "9622", "9630", "9690", "9790"]}
    # 핵심 계수표를 미리 읽어 둔다(분석 탭 공통 사용)
    for key in ["import", "B", "V", "M", "fd_out", "fd_imp", "fd_va"]:
        G(key)
    _log("지표 계산 중")
    _compute_metrics()
    return {"level": LEVELS[level], "n": S["n"], "sheets": _sheet_list(), "national": _national(),
            "checks": _check_accounts(), "press": _press_check(), "metrics_meta": METRIC_META}


def _sheet_list():
    out = []
    for key, name, label, unit, desc in SHEETS:
        info = {"key": key, "sheet": name, "label": label, "unit": unit, "desc": desc, "rows": None, "cols": None,
                "loaded": key in S["grids"], "present": name in S["wb"].sheetnames}
        if key in S["grids"]:
            gg = S["grids"][key]
            info["rows"], info["cols"] = len(gg["rc"]), len(gg["cc"])
        elif info["present"]:
            info["rows"], info["cols"] = S["wb"].dims(name)
        out.append(info)
    return out


# ----------------------------------------------------------------------------- 지표 계산
METRIC_META = [
    {"key": "out", "label": "총산출", "unit": "백만원", "desc": "국내 생산액(총투입과 같음)"},
    {"key": "out_share", "label": "총산출 비중", "unit": "%", "desc": "전산업 총산출 대비 비중"},
    {"key": "va_ratio", "label": "부가가치율", "unit": "%", "desc": "부가가치 ÷ 총투입"},
    {"key": "mid_ratio", "label": "중간투입률", "unit": "%", "desc": "중간투입 ÷ 총투입"},
    {"key": "imp_dep", "label": "수입의존도", "unit": "%", "desc": "수입 중간재 투입 ÷ 총투입"},
    {"key": "dom_ratio", "label": "국산화율", "unit": "%", "desc": "국산 중간재 투입 ÷ 중간투입"},
    {"key": "mid_dem_ratio", "label": "중간수요율", "unit": "%", "desc": "중간수요 ÷ 총수요"},
    {"key": "export_rate", "label": "수출률", "unit": "%", "desc": "수출 ÷ 총산출"},
    {"key": "import_rate", "label": "수입률", "unit": "%", "desc": "수입계 ÷ 총공급(= 총수요)"},
    {"key": "prod", "label": "생산유발계수", "unit": "무차원", "desc": "최종수요 1단위가 전 산업에 유발하는 생산(열합)"},
    {"key": "va_mult", "label": "부가가치유발계수", "unit": "무차원", "desc": "최종수요 1단위가 유발하는 부가가치(열합)"},
    {"key": "im_mult", "label": "수입유발계수", "unit": "무차원", "desc": "최종수요 1단위가 유발하는 수입(열합)"},
    {"key": "infl", "label": "영향력계수", "unit": "지수(평균=1)", "desc": "후방연쇄효과: 열합 ÷ 열합 평균"},
    {"key": "sens", "label": "감응도계수", "unit": "지수(평균=1)", "desc": "전방연쇄효과: 행합 ÷ 행합 평균"},
]


def _type_of(infl, sens):
    if not (math.isfinite(infl) and math.isfinite(sens)):
        return "-"
    if infl >= 1 and sens >= 1:
        return "중추형(영향력·감응도 모두 큼)"
    if infl >= 1:
        return "최종수요 의존형(영향력 큼)"
    if sens >= 1:
        return "중간재 공급형(감응도 큼)"
    return "독립형(영향력·감응도 모두 작음)"


def _compute_metrics():
    ind, n = S["ind"], S["n"]
    c, r = S["col"], S["row"]
    gi = G("import")
    imp_mid = _sub(gi, ind, ind).sum(axis=0)  # 산업 j 의 수입 중간재 투입 합
    gb, gv, gm = G("B"), G("V"), G("M")
    prod = _sub(gb, ind, ind).sum(axis=0)
    vam = _sub(gv, ind, ind).sum(axis=0)
    imm = _sub(gm, ind, ind).sum(axis=0)
    infl = _rowv(gb, "9801", ind)
    sens = _colv(gb, "9802", ind)
    out = c["9310"]
    tot = float(np.nansum(out))
    m = {
        "out": out,
        "out_share": 100 * out / tot if tot else np.full(n, np.nan),
        "va_ratio": 100 * _sd(r["9690"], r["9790"]),
        "mid_ratio": 100 * _sd(r["9590"], r["9790"]),
        "imp_dep": 100 * _sd(imp_mid, r["9790"]),
        "dom_ratio": 100 * (1 - _sd(imp_mid, r["9590"])),
        "mid_dem_ratio": 100 * _sd(c["9090"], c["9290"]),
        "export_rate": 100 * _sd(c["9140"], out),
        "import_rate": 100 * _sd(c["9329"], c["9390"]),
        "prod": np.where(out != 0, prod, np.nan),
        "va_mult": np.where(out != 0, vam, np.nan),
        "im_mult": np.where(out != 0, imm, np.nan),
        "infl": np.where(out != 0, infl, np.nan),
        "sens": np.where(out != 0, sens, np.nan),
    }
    S["metrics"] = m
    S["type"] = [_type_of(m["infl"][i], m["sens"][i]) for i in range(n)]
    S["imp_mid"] = imp_mid


def metrics(a=None):
    n = S["n"]
    m = S["metrics"]
    rows = []
    for i in range(n):
        row = {"c": S["ind"][i], "n": S["names"][i], "type": S["type"][i]}
        for k, _ in [(x["key"], 0) for x in METRIC_META]:
            row[k] = _f(m[k][i], 6 if k in ("prod", "va_mult", "im_mult", "infl", "sens") else 3)
        rows.append(row)
    cnt = {}
    for t in S["type"]:
        cnt[t] = cnt.get(t, 0) + 1
    return {"meta": METRIC_META, "rows": rows, "type_count": cnt, "level": LEVELS[S["level"]]}


# ----------------------------------------------------------------------------- 국민경제 요약·정합성
def _national():
    c, r = S["col"], S["row"]
    T = 1e6  # 백만원 -> 조원
    tot = {k: float(v.sum()) for k, v in c.items()}
    rt = {k: float(v.sum()) for k, v in r.items()}
    out, dem = tot["9310"], tot["9290"]
    res = {
        "out": out, "imp_all": tot["9329"], "sup": tot["9390"], "mid_dem": tot["9090"],
        "fd": tot["9190"], "mid_in": rt["9590"], "va": rt["9690"], "tin": rt["9790"],
        "fd_items": [{"c": k, "n": n, "v": tot[k], "share": 100 * tot[k] / tot["9190"] if tot["9190"] else None}
                     for k, n in FD_ITEMS],
        "va_items": [{"c": k, "n": n, "v": rt[k], "share": 100 * rt[k] / rt["9690"] if rt["9690"] else None}
                     for k, n in VA_ITEMS],
        "mid_ratio": 100 * rt["9590"] / rt["9790"], "va_ratio": 100 * rt["9690"] / rt["9790"],
        "export_rate": 100 * tot["9140"] / out, "import_rate": 100 * tot["9329"] / tot["9390"],
        "trade_share": 100 * (tot["9140"] + tot["9329"]) / tot["9390"],
        "imp_dep": 100 * float(S["imp_mid"].sum()) / rt["9790"],
        "T": T,
    }
    return res


def _press_check():
    nat = _national()
    T = nat["T"]
    fi = {x["c"]: x["v"] for x in nat["fd_items"]}
    vi = {x["c"]: x["v"] for x in nat["va_items"]}
    mine = [nat["out"], nat["imp_all"], nat["sup"], nat["mid_dem"], nat["va"], vi["9610"], vi["9621"], vi["9622"],
            vi["9630"], fi["9111"] + fi["9112"], fi["9121"] + fi["9122"] + fi["9131"] + fi["9132"], fi["9140"]]
    out = []
    for (label, ref), v in zip(PRESS, mine):
        v = v / T
        out.append({"label": label, "press": ref, "calc": _f(v, 1), "diff": _f(v - ref, 2), "ok": abs(v - ref) <= 0.15})
    return out


def _check_accounts():
    """총거래표 하나로 확인하는 회계 항등식 (불러오기 때마다 자동 실행)"""
    c, r, X = S["col"], S["row"], S["X"]

    def mx(a, b):
        return float(np.max(np.abs(np.asarray(a) - np.asarray(b)))) if len(a) else 0.0

    fd_sum = sum(c[k] for k, _ in FD_ITEMS)
    va_sum = sum(r[k] for k, _ in VA_ITEMS)
    items = [
        ("중간수요계 = 거래행렬 행합", "Σ_j x_ij = 중간수요계_i", mx(X.sum(axis=1), c["9090"]), 1.0),
        ("중간투입계 = 거래행렬 열합", "Σ_i x_ij = 중간투입계_j", mx(X.sum(axis=0), r["9590"]), 1.0),
        ("최종수요계 = 7개 항목 합", "소비+투자+재고+귀중품+수출", mx(fd_sum, c["9190"]), 1.0),
        ("총수요계 = 중간수요계 + 최종수요계", "", mx(c["9090"] + c["9190"], c["9290"]), 1.0),
        ("총수요계 = 총산출 + 수입계", "수요·공급 균형", mx(c["9310"] + c["9329"], c["9290"]), 1.0),
        ("총투입 = 총산출", "산업별 생산액 균형", mx(r["9790"], c["9310"]), 1.0),
        ("총투입 = 중간투입계 + 부가가치계", "", mx(r["9590"] + r["9690"], r["9790"]), 1.0),
        ("부가가치계 = 4개 부가가치 항목 합", "피용자보수+영업잉여+고정자본소모+생산세", mx(va_sum, r["9690"]), 1.0),
    ]
    return [{"name": a, "formula": b, "maxdiff": _f(d, 6), "tol": t, "ok": d <= t, "unit": "백만원"} for a, b, d, t in items]


# ----------------------------------------------------------------------------- 심화 검증(계수표)
def validate(a=None):
    ind, n = S["ind"], S["n"]
    out = S["col"]["9310"]
    live = out != 0
    res = []

    def add(name, formula, d, tol, unit):
        res.append({"name": name, "formula": formula, "maxdiff": _f(d, 8), "tol": tol, "ok": bool(d <= tol), "unit": unit})

    X, tin = S["X"], S["row"]["9790"]
    A = _sub(G("A"), ind, ind)
    Acalc = np.zeros_like(X)
    Acalc[:, live] = X[:, live] / tin[live]
    add("투입계수 재계산", "a_ij = x_ij ÷ 총투입_j  (Excel 총투입계수표와 비교)", np.max(np.abs(A[:, live] - Acalc[:, live])), 1e-9, "무차원")

    Ad, Am = _sub(G("Ad"), ind, ind), _sub(G("Am"), ind, ind)
    add("국산 + 수입 투입계수 = 총투입계수", "Ad + Am = A", np.max(np.abs(Ad + Am - A)), 1e-9, "무차원")
    B = _sub(G("B"), ind, ind)
    Binv = np.linalg.inv(np.eye(n) - Ad)
    add("레온티에프 역행렬", "B = (I - Ad)^-1  (Excel 생산유발계수표와 비교)", np.max(np.abs(B - Binv)), 1e-8, "무차원")
    v = _sd(S["row"]["9690"], tin)
    v = np.where(np.isfinite(v), v, 0.0)
    V = _sub(G("V"), ind, ind)
    add("부가가치유발계수", "V_ij = v_i · B_ij  (v = 부가가치율)", np.max(np.abs(V - v[:, None] * B)), 1e-8, "무차원")
    M = _sub(G("M"), ind, ind)
    add("수입유발계수", "M = Am · B", np.max(np.abs(M - Am @ B)), 1e-8, "무차원")
    add("유발 균형", "열합(부가가치유발) + 열합(수입유발) = 1  (산출이 있는 산업)",
        np.max(np.abs(V.sum(axis=0)[live] + M.sum(axis=0)[live] - 1)), 1e-8, "무차원")
    gd = G("domestic")
    fdom = _colv(gd, "9190", ind)
    add("생산유발 일치", "B · (국산 최종수요) = 총산출", np.max(np.abs(B @ fdom - out)), 1.0, "백만원")
    infl = B.sum(axis=0) / B.sum(axis=0).mean()
    add("영향력계수", "열합 ÷ 열합 평균  (Excel 값과 비교)", np.max(np.abs(infl - _rowv(G("B"), "9801", ind))), 1e-8, "지수")
    return {"level": LEVELS[S["level"]], "checks": res, "zero_output": int((~live).sum())}


# ----------------------------------------------------------------------------- 시트 보기
def view(a):
    key = a["key"]
    g = G(key)
    q = (a.get("q") or "").strip().lower()
    cq = (a.get("cq") or "").strip().lower()
    c0, cn = int(a.get("c0", 0)), int(a.get("cn", 10))
    maxr = int(a.get("maxr", 300))
    unit = SHEET_BY_KEY[key][3]
    nd = 0 if unit == "백만원" else 6
    ridx = [i for i in range(len(g["rc"])) if not q or q in g["rc"][i].lower() or q in g["rn"][i].lower()]
    cidx = [j for j in range(len(g["cc"])) if not cq or cq in g["cc"][j].lower() or cq in g["cn"][j].lower()]
    win = cidx[c0:c0 + cn]
    rows = [{"c": g["rc"][i], "n": g["rn"][i], "v": [_f(g["M"][i, j], nd) for j in win]} for i in ridx[:maxr]]
    return {"key": key, "label": SHEET_BY_KEY[key][2], "unit": unit, "desc": SHEET_BY_KEY[key][4],
            "nrows": len(g["rc"]), "ncols": len(g["cc"]), "mrows": len(ridx), "mcols": len(cidx), "c0": c0,
            "cols": [{"c": g["cc"][j], "n": g["cn"][j]} for j in win], "rows": rows, "truncated": len(ridx) > maxr}


# ----------------------------------------------------------------------------- 산업 검색
def search(a):
    q = (a.get("q") or "").strip().lower()
    cur = S["level"]
    cls = S["cls"]
    if not q:
        return {"total": 0, "items": []}
    hits = {}  # 현재 분류 code -> {direct, via}

    def put(code, via, direct):
        h = hits.setdefault(code, {"direct": False, "via": []})
        h["direct"] = h["direct"] or direct
        if via and len(h["via"]) < 3 and via not in h["via"]:
            h["via"].append(via)

    for i, code in enumerate(S["ind"]):
        if q in code.lower() or q in S["names"][i].lower():
            put(code, "", True)
    for lv in LEVEL_ORDER:
        if lv == cur:
            continue
        for code, name in cls["ent"][lv].items():
            if q in code.lower() or q in name.lower():
                curs = []
                for k in cls["mem"][lv][code]:
                    e = cls["recs"][k][cur]
                    if e and e[0] not in curs:
                        curs.append(e[0])
                for cc in curs:
                    if cc in S["ix"]:
                        put(cc, "%s %s %s" % (LEVELS[lv], code, name), False)
    items = []
    for code in sorted(hits, key=lambda c: (not hits[c]["direct"], S["ix"][c])):
        items.append({"c": code, "n": S["names"][S["ix"][code]], "direct": hits[code]["direct"], "via": hits[code]["via"]})
    return {"total": len(items), "items": items[:80]}


def industries(a=None):
    return {"items": [{"c": c, "n": S["names"][i]} for i, c in enumerate(S["ind"])]}


# ----------------------------------------------------------------------------- 산업 상세
def _top(vec, denom, k=8):
    idx = [i for i in np.argsort(-np.asarray(vec)) if vec[i] > 0][:k]
    out = []
    for i in idx:
        out.append({"c": S["ind"][i], "n": S["names"][i], "v": _f(vec[i], 0), "r": _f(100 * vec[i] / denom, 2) if denom else None})
    return out


def profile(a):
    code = a["code"]
    i = S["ix"][code]
    c, r, X, m = S["col"], S["row"], S["X"], S["metrics"]
    out = float(c["9310"][i])
    tin = float(r["9790"][i])
    dem = float(c["9290"][i])
    res = {
        "c": code, "n": S["names"][i], "level": LEVELS[S["level"]], "type": S["type"][i], "hier": _hier(code),
        "zero": out == 0,
        "supply": {"out": out, "imp": float(c["9321"][i]), "impt": float(c["9322"][i]), "imp_all": float(c["9329"][i]),
                   "sup": float(c["9390"][i])},
        "demand": {"mid_dem": float(c["9090"][i]), "fd": float(c["9190"][i]), "dem": dem,
                   "items": [{"c": k, "n": n, "v": float(c[k][i]), "share": _f(100 * c[k][i] / dem, 2) if dem else None} for k, n in FD_ITEMS]},
        "input": {"tin": tin, "mid_in": float(r["9590"][i]), "va": float(r["9690"][i]),
                  "items": [{"c": k, "n": n, "v": float(r[k][i]), "share": _f(100 * r[k][i] / tin, 2) if tin else None} for k, n in VA_ITEMS]},
        "metrics": {k["key"]: _f(m[k["key"]][i], 4) for k in METRIC_META},
        "top_inputs": _top(X[:, i], tin) if tin else [],
        "top_buyers": _top(X[i, :], float(c["9090"][i])) if c["9090"][i] else [],
    }
    # 가장 영향을 주는 최종수요 항목 (총수요 대비)
    return res


# ----------------------------------------------------------------------------- 최종수요 충격 시뮬레이션
def simulate(a):
    ind, n = S["ind"], S["n"]
    f = np.zeros(n)
    for s in a["shocks"]:
        f[S["ix"][s["code"]]] += float(s["amount"])
    if not np.any(f):
        raise ValueError("충격 금액이 모두 0입니다.")
    B = _sub(G("B"), ind, ind)
    V = _sub(G("V"), ind, ind)
    M = _sub(G("M"), ind, ind)
    dx, dv, dm = B @ f, V @ f, M @ f
    tin = S["row"]["9790"]
    comp = []
    for k, nm in VA_ITEMS:
        ratio = np.where(tin != 0, S["row"][k] / np.where(tin != 0, tin, 1), 0.0)
        comp.append({"c": k, "n": nm, "v": float((ratio * dx).sum())})
    tf, tx, tv, tm = float(f.sum()), float(dx.sum()), float(dv.sum()), float(dm.sum())
    order = np.argsort(-np.abs(dx))
    rows = []
    for i in order:
        if abs(dx[i]) < 1e-9 and abs(f[i]) < 1e-9:
            continue
        rows.append({"c": ind[i], "n": S["names"][i], "direct": _f(f[i], 1), "dx": _f(dx[i], 1), "indirect": _f(dx[i] - f[i], 1),
                     "share": _f(100 * dx[i] / tx, 2) if tx else None, "dva": _f(dv[i], 1), "dim": _f(dm[i], 1)})
    return {"final_demand": tf, "dx": tx, "dva": tv, "dim": tm, "mult_prod": _f(tx / tf, 4), "mult_va": _f(tv / tf, 4),
            "mult_im": _f(tm / tf, 4), "closure": _f(tv + tm - tf, 6), "comp": comp, "rows": rows[:400], "nrows": len(rows)}


# ----------------------------------------------------------------------------- 최종수요 항목별 유발효과
def fdview(a):
    item = a.get("item", "9190")
    topn = int(a.get("top", 15))
    items = FD_ITEMS + [("9190", "최종수요계")]
    gs = {"out": G("fd_out"), "va": G("fd_va"), "im": G("fd_imp")}
    table = []
    for code, nm in items:
        row = {"c": code, "n": nm}
        for k, g in gs.items():
            j = g["ci"].get(code)
            row[k] = _f(g["M"][g["ri"]["9819"], j], 0) if j is not None and "9819" in g["ri"] else None
            row[k + "_share"] = _f(g["M"][g["ri"]["9812"], j], 2) if j is not None and "9812" in g["ri"] else None
            row[k + "_mult"] = _f(g["M"][g["ri"]["9811"], j], 4) if j is not None and "9811" in g["ri"] else None
        table.append(row)
    ind = S["ind"]
    sel = {}
    for k, g in gs.items():
        sel[k] = _colv(g, item, ind)
    order = np.argsort(-sel["out"])[:topn]
    top = [{"c": ind[i], "n": S["names"][i], "out": _f(sel["out"][i], 0), "va": _f(sel["va"][i], 0), "im": _f(sel["im"][i], 0)} for i in order]
    name = dict(items).get(item, item)
    return {"table": table, "item": item, "item_name": name, "top": top}


# ----------------------------------------------------------------------------- 진입점
CMDS = {"load": load, "metrics": metrics, "validate": validate, "view": view, "search": search,
        "industries": industries, "profile": profile, "simulate": simulate, "fdview": fdview}


def handle(cmd, args_json):
    """worker.js 가 호출한다. 항상 JSON 문자열을 돌려준다."""
    try:
        args = json.loads(args_json) if args_json else {}
        res = CMDS[cmd](args)
        return json.dumps({"ok": True, "data": res}, ensure_ascii=False, allow_nan=False)
    except Exception as e:  # noqa: BLE001
        return json.dumps({"ok": False, "error": "%s: %s" % (type(e).__name__, e), "trace": traceback.format_exc()[-1500:]},
                          ensure_ascii=False)
