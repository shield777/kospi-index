# -*- coding: utf-8 -*-
"""
스프레드·기관 지수 자동 계산 → 휴대폰 웹페이지(GitHub Pages) 갱신

평소: 실행.bat 의 자동 수집이 끝난 바로 다음 줄에서
      python spread_index.py
      를 부르면 끝. 따로 손댈 것 없음.

다른 모드
  python spread_index.py --check          수집 파일을 어떻게 읽었는지만 보여 줌 (아무것도 안 씀)
  python spread_index.py --setup 깃허브아이디   GitHub Pages 저장소 첫 설정 (한 번만)

지키는 규칙
  - 0_경기지표.xlsx 는 열지도 쓰지도 않음
  - 수집 파일은 임시 폴더로 복사해서 읽기만 함
  - 오류가 나면 웹페이지·기준이력을 덮어쓰지 않고 로그만 남김
"""
import os, sys, re, json, glob, shutil, tempfile, subprocess, traceback, math, time
from datetime import datetime, date, timedelta

# ────────────────────────────── 설정 (경로가 바뀌면 여기만 고치기) ──────────────────────────────
DATA_DIR = r"C:\Users\Shiel\OneDrive\주식\0.경기지표\data"
REPO_DIR = r"C:\Users\Shiel\kospi-index"
REPO_NAME = "kospi-index"
# 화면 디자인(템플릿) 자동 받기: 매 실행마다 GitHub 의 tpl 가지에서 최신 디자인을 확인. 빈칸("")이면 끔
TEMPLATE_URL = "https://raw.githubusercontent.com/shield777/kospi-index/tpl/spread_index_tpl.html"
# ───────────────────────────────────────────────────────────────────────────────────────────────

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "spread_index_tpl.html")
COLLECT_GLOB = "경기지표_자동수집_*.xlsx"
BASE_CSV = "지수_기준이력.csv"
BASE_TXT = "지수_기준이력_원본.txt"          # 기준이력 csv 가 없을 때만 씀
RECORD_CSV = "지수기록.csv"
LOG_TXT = "자동업데이트_기록.txt"
LOCAL_HTML = "지수_최신.html"
LOCK = "지수계산.lock"
ITEMS = ["코스피", "국채3년", "국채10년", "스프레드", "기관", "외국인"]
PREFER_SHEET = "일단위"                      # 수집 파일에서 먼저 보는 시트
OVERWRITE_DAYS = 7

try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:
    pass

import numpy as np
import pandas as pd


class StopRun(Exception):
    """계산을 멈추고 이전 페이지를 유지해야 하는 오류"""


# ═══════════════════════════════ 로그 / 잠금 / 안전한 쓰기 ═══════════════════════════════
_log_lines = []


def say(msg):
    print(msg)
    _log_lines.append(msg)


def write_log(data_dir, status, summary):
    path = os.path.join(data_dir, LOG_TXT)
    enc = "utf-8"
    if os.path.exists(path):
        try:
            with open(path, "rb") as f:
                f.read().decode("utf-8")
        except UnicodeDecodeError:
            enc = "cp949"
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [f"{stamp} [지수] {status}: {summary}"]
    lines += [f"    {x}" for x in _log_lines if x.strip()]
    try:
        with open(path, "a", encoding=enc, errors="replace") as f:
            f.write("\n".join(lines) + "\n")
    except Exception as e:  # 로그 실패로 작업 전체를 망치지 않음
        print("로그 쓰기 실패:", e)


def acquire_lock(data_dir):
    path = os.path.join(data_dir, LOCK)
    for _ in range(2):
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            return path
        except FileExistsError:
            # 30분 넘은 잠금은 비정상 종료로 남은 것 → 지움
            if time.time() - os.path.getmtime(path) > 1800:
                os.remove(path)
                continue
            raise StopRun("다른 지수 계산이 실행 중입니다 (잠금 파일 있음)")
    raise StopRun("잠금 파일을 만들 수 없습니다")


def atomic_write(path, data, encoding="utf-8"):
    """임시 파일에 다 쓴 뒤 한 번에 바꿔치기 → 중간에 실패해도 원래 파일이 남음"""
    d = os.path.dirname(path) or "."
    fd, tmp = tempfile.mkstemp(dir=d, prefix=".tmp_", suffix=os.path.splitext(path)[1])
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data if isinstance(data, bytes) else data.encode(encoding))
        for i in range(5):  # OneDrive가 잠깐 잡고 있을 때 재시도
            try:
                os.replace(tmp, path)
                return
            except PermissionError:
                time.sleep(1 + i)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


# ═══════════════════════════════ 1. 입력: 기준이력 ═══════════════════════════════
def to_num(s):
    if not pd.api.types.is_numeric_dtype(s):
        s = s.astype(object).where(s.notna(), None).astype(str).str.replace(",", "", regex=False).str.strip()
        s = s.replace({"": None, "nan": None, "None": None, "-": None})
    return pd.to_numeric(s, errors="coerce")


def load_base_txt(path):
    df = pd.read_csv(path, sep="\t", dtype=str, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]
    dt = pd.to_datetime(df["년"].str.strip() + "-" + df["월"].str.strip() + "-" + df["일"].str.strip())
    out = pd.DataFrame({c: to_num(df[c]) if c in df.columns else np.nan for c in ITEMS})
    out.index = pd.DatetimeIndex(dt, name="날짜")
    return fill_spread(out)


def fill_spread(df):
    """스프레드 칸이 비어 있으면 국채10년 − 국채3년으로 채움"""
    if "스프레드" not in df.columns:
        df["스프레드"] = np.nan
    calc = df["국채10년"] - df["국채3년"]
    df["스프레드"] = df["스프레드"].fillna(calc)
    return df[ITEMS]


def load_base(data_dir):
    p = os.path.join(data_dir, BASE_CSV)
    if os.path.exists(p):
        df = pd.read_csv(p, encoding="utf-8-sig", dtype={"날짜": str})
        df.index = pd.DatetimeIndex(pd.to_datetime(df.pop("날짜")), name="날짜")
        for c in ITEMS:
            if c not in df.columns:
                df[c] = np.nan
        return fill_spread(df[ITEMS].apply(to_num)), False
    for cand in [os.path.join(data_dir, BASE_TXT), os.path.join(HERE, BASE_TXT)]:
        if os.path.exists(cand):
            say(f"기준이력 csv가 없어 원본 텍스트로 새로 만듭니다: {os.path.basename(cand)}")
            return load_base_txt(cand), True
    raise StopRun(f"{BASE_CSV} 가 data 폴더에 없습니다 (함께 드린 파일을 data 폴더에 넣어 주세요)")


def save_base(data_dir, df):
    out = df.copy()
    out.index = out.index.strftime("%Y-%m-%d")
    p = os.path.join(data_dir, BASE_CSV)
    if os.path.exists(p):
        shutil.copy2(p, os.path.join(data_dir, "지수_기준이력_백업.csv"))
    atomic_write(p, out.to_csv(index_label="날짜", float_format="%.10g"), encoding="utf-8-sig")


# ═══════════════════════════════ 1. 입력: 자동 수집 파일 ═══════════════════════════════
def norm(h):
    return re.sub(r"[\s()\[\]{}_\-./·:]", "", str(h)).lower()


def item_score(item, h):
    """머리글 h 가 item 과 얼마나 맞는지 (0 = 아님, 클수록 확실)"""
    n = norm(h)
    if not n or n in ("nan", "none"):
        return 0
    bad_common = ["등락", "변동", "%", "비중", "비율", "보유", "지분", "누적", "전일", "대비", "평균", "거래량", "수익률"]
    if any(b in n for b in bad_common):
        return 0
    if item == "코스피":
        if not ("코스피" in n or "kospi" in n):
            return 0
        if any(b in n for b in ["200", "코스닥", "kosdaq", "거래", "선물", "시가", "고가", "저가", "외국", "기관", "개인"]):
            return 0
        return 10 if n in ("코스피", "kospi", "코스피종가", "kospi종가", "코스피지수", "kospiclose", "종가코스피") else 6
    if item in ("국채3년", "국채10년"):
        tenor = "3년" if item == "국채3년" else "10년"
        tenor_en = ["3y", "3yr", "3년물"] if item == "국채3년" else ["10y", "10yr", "10년물"]
        if not ("국채" in n or "국고채" in n or "ktb" in n):
            return 0
        if any(b in n for b in ["미국", "us", "회사채", "통안", "스프레드", "선물", "30년", "20년", "50년", "13년"]):
            return 0
        if item == "국채3년" and "10년" in n:
            return 0
        if not (tenor in n or any(t in n for t in tenor_en)):
            return 0
        return 10 if n in (item, "국고채" + tenor, "국고채" + tenor + "물", "국채" + tenor + "물") else 6
    if item == "스프레드":
        if not ("스프레드" in n or "spread" in n or "장단기" in n):
            return 0
        if any(b in n for b in ["미국", "us", "회사채", "신용", "크레딧", "credit", "bbb", "aa", "하이일드", "통안", "2년", "5년", "20년", "30년"]):
            return 0
        return 10 if n in ("스프레드", "spread", "국채스프레드", "장단기스프레드", "스프레드10년3년", "10년3년스프레드") else 6
    if item == "기관":
        if "기관" not in n or any(b in n for b in ["외국", "개인", "코스닥", "kosdaq"]):
            return 0
        return 10 if n in ("기관", "기관합계", "기관순매수", "기관계", "코스피기관", "기관순매수코스피") else 6
    if item == "외국인":
        if "외국인" not in n or any(b in n for b in ["기관", "개인", "코스닥", "kosdaq"]):
            return 0
        return 10 if n in ("외국인", "외국인합계", "외국인순매수", "외국인계", "코스피외국인", "외국인순매수코스피") else 6
    return 0


def parse_dates(s):
    """엑셀 날짜·'2026-10-30'·20261030 등 여러 형태를 날짜로"""
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    for i, v in s.items():
        if v is None or (isinstance(v, float) and math.isnan(v)):
            continue
        if isinstance(v, (datetime, date, pd.Timestamp)):
            out[i] = pd.Timestamp(v)
            continue
        t = str(v).strip()
        if re.fullmatch(r"\d{8}(\.0)?", t):
            t = t[:8]
            out[i] = pd.to_datetime(t, format="%Y%m%d", errors="coerce")
        elif re.match(r"\d{4}[-./년 ]\s*\d{1,2}[-./월 ]\s*\d{1,2}", t):
            t2 = re.sub(r"[./년월]\s*", "-", t).replace("일", "").strip()
            out[i] = pd.to_datetime(t2.split()[0], errors="coerce")
    ok = out.notna() & (out.dt.year >= 1990) & (out.dt.year <= 2100)
    return out.where(ok)


def read_sheet(raw, where):
    """시트 하나에서 날짜 + 필요 항목 열을 찾아 돌려줌 (머리글 이름으로 찾음)"""
    best = None
    for r in range(min(20, len(raw))):
        row = raw.iloc[r].tolist()
        hits = sum(1 for h in row if any(item_score(it, h) for it in ITEMS))
        if hits and (best is None or hits > best[1]):
            best = (r, hits)
    if best is None:
        return []
    hr = best[0]
    headers = raw.iloc[hr].tolist()
    body = raw.iloc[hr + 1:].reset_index(drop=True)
    # 날짜 열: 이름이 날짜/일자 이거나, 값의 대부분이 날짜인 열
    date_col, ymd = None, None
    names = [norm(h) for h in headers]
    for j, n in enumerate(names):
        if n in ("날짜", "일자", "date", "기준일", "기준일자", "일시", "거래일", "datetime") or n.endswith("날짜") or n.endswith("일자"):
            date_col = j
            break
    if date_col is None and {"년", "월", "일"} <= set(names):
        ymd = (names.index("년"), names.index("월"), names.index("일"))
    if date_col is None and ymd is None:
        for j in range(len(headers)):
            col = body.iloc[:, j]
            nn = col.notna().sum()
            if nn >= 3 and parse_dates(col).notna().sum() >= 0.8 * nn:
                date_col = j
                break
    if date_col is None and ymd is None:
        return []
    if ymd:
        y, m, d = (to_num(body.iloc[:, k]) for k in ymd)
        dt = pd.to_datetime(dict(year=y, month=m, day=d), errors="coerce")
    else:
        dt = parse_dates(body.iloc[:, date_col])
    found = []
    for j, h in enumerate(headers):
        for it in ITEMS:
            sc = item_score(it, h)
            if sc:
                vals = to_num(body.iloc[:, j])
                s = pd.Series(vals.values, index=dt).loc[lambda x: x.index.notna()]
                s = s[~s.index.duplicated(keep="last")].dropna()
                if len(s):
                    pref = norm(where.split(" / ")[-1]) == norm(PREFER_SHEET)
                    found.append(dict(item=it, score=sc, n=len(s), header=str(h).strip(), where=where, pref=pref, s=s))
    return found


def read_collect_file(path, tmpdir):
    dst = os.path.join(tmpdir, "c_" + os.path.basename(path))
    shutil.copy2(path, dst)  # 열려 있어도 되도록 복사본을 읽음
    sheets = pd.read_excel(dst, sheet_name=None, header=None)
    found = []
    for name, raw in sheets.items():
        found += read_sheet(raw, f"{os.path.basename(path)} / {name}")
    return found


def load_collect(data_dir, check=False, need_from=None):
    files = sorted(glob.glob(os.path.join(data_dir, COLLECT_GLOB)))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        raise StopRun(f"수집 파일({COLLECT_GLOB})이 data 폴더에 없습니다")
    tmpdir = tempfile.mkdtemp(prefix="spread_idx_")
    per_file = []
    try:
        # 최신 파일부터 읽다가, 필요한 기간(기준이력 끝 − 14일 이후)을 이미 덮으면 나머지는 안 읽음 → 빨리 끝남
        for f in reversed(files):
            try:
                fl = read_collect_file(f, tmpdir)
            except Exception as e:
                say(f"수집 파일 읽기 실패(건너뜀): {os.path.basename(f)} - {e}")
                continue
            per_file.append((f, fl))
            starts = [x["s"].index.min() for x in fl if x["item"] == "코스피"]
            if need_from is not None and starts and min(starts) <= need_from:
                break
        per_file.reverse()  # 날짜 순 → 나중 파일이 같은 날짜를 덮어씀
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    if not per_file:
        raise StopRun("수집 파일을 하나도 읽지 못했습니다")

    # 항목별로 가장 잘 맞는 열 하나 고르기 (최신 파일 기준: 점수 → 데이터 수)
    latest = per_file[-1][1]
    choice = {}
    for it in ITEMS:
        c = [x for x in latest if x["item"] == it]
        if not c:
            for _, fl in reversed(per_file):
                c = [x for x in fl if x["item"] == it]
                if c:
                    break
        if c:
            c.sort(key=lambda x: (x["pref"], x["score"], x["n"]), reverse=True)
            choice[it] = (c[0]["header"], c[0]["where"].split(" / ")[1])
            if check:
                for x in c:
                    say(f"  {it:5s} 후보: '{x['header']}' ({x['where']}, {x['n']}행, 점수 {x['score']})")
    missing = [it for it in ["코스피", "기관", "외국인"] if it not in choice]
    if "스프레드" not in choice and not ("국채3년" in choice and "국채10년" in choice):
        missing.append("스프레드(또는 국채3년·국채10년)")
    if missing:
        raise StopRun("수집 파일에서 이 항목의 열을 찾지 못했습니다: " + ", ".join(missing))

    cols = {}
    for it in [x for x in ITEMS if x in choice]:
        hdr, sheet = choice[it]
        parts = []
        for f, fl in per_file:
            m = [x for x in fl if x["item"] == it and x["header"] == hdr and x["where"].endswith(" / " + sheet)]
            if m:
                parts.append(m[0]["s"])
        s = pd.concat(parts)
        cols[it] = s[~s.index.duplicated(keep="last")]
    df = pd.DataFrame(cols).sort_index().reindex(columns=ITEMS)
    df.index = pd.DatetimeIndex(df.index.normalize(), name="날짜")
    df = df[~df.index.duplicated(keep="last")]
    say(f"수집 파일 {len(per_file)}/{len(files)}개 읽음, 열: " + ", ".join(f"{it}='{choice[it][0]}'" for it in ITEMS if it in choice)
        + f", 기간 {df.index.min():%Y-%m-%d}~{df.index.max():%Y-%m-%d}")
    return clean_collect(df)


def clean_collect(df):
    # 주말 행 제거
    wk = df.index.dayofweek >= 5
    if wk.any():
        say(f"주말 행 {int(wk.sum())}개 제외")
        df = df[~wk]
    # 휴장일에 앞 날짜 값을 그대로 복사해 넣은 행(코스피·기관·외국인이 전날과 완전히 같음) → 수급 중복 합산 방지
    k = ["코스피", "기관", "외국인"]
    same = (df[k] == df[k].shift(1)).all(axis=1) & df[k].notna().all(axis=1)
    if same.any():
        say(f"전날과 똑같은 값(휴장일 복사로 보임) {int(same.sum())}행의 코스피·수급 제외")
        df.loc[same, k] = np.nan
    # 상식 범위 밖 값 제거
    for c, lo, hi in [("코스피", 100, 50000), ("국채3년", -2, 25), ("국채10년", -2, 25), ("스프레드", -500, 500)]:
        bad = df[c].notna() & ((df[c] < lo) | (df[c] > hi))
        if bad.any():
            say(f"{c} 범위 밖 값 {int(bad.sum())}개 제외")
            df.loc[bad, c] = np.nan
    return df.dropna(how="all")


def unit_factor(base, coll, col):
    """수집 파일 금액 단위를 기준이력(백만원)에 맞추는 나눗수"""
    both = pd.concat([base[col], coll[col]], axis=1, keys=["b", "c"]).dropna()
    both = both[(both.b.abs() > 0) & (both.c.abs() > 0)]
    if len(both) >= 5:
        r = (both.c / both.b)
        r = r[r > 0]
        if len(r) >= 3:
            e = math.log10(r.median())
            f = 10 ** round(e)
            corr = both.c.corr(both.b)
            note = f"겹치는 {len(both)}일 비교, 배율 {r.median():.4g}, 상관 {corr:.3f}"
            if abs(e - round(e)) > 0.15 or corr < 0.9:
                say(f"⚠ {col} 값이 기준이력과 잘 맞지 않습니다 ({note}) — 다른 종류의 값일 수 있음")
            return f, note
    bm = base[col].dropna().abs().tail(250).median()
    cm = coll[col].dropna().abs().median()
    if not (bm > 0 and cm > 0):
        return 1, "비교 불가(1로 둠)"
    e = math.log10(cm / bm)
    opts = [-2, 0, 3, 6]          # 억원, 백만원, 천원, 원
    k = min(opts, key=lambda o: abs(o - e))
    return 10 ** k, f"겹치는 날짜 부족 → 크기로 판단 (배율 약 10^{e:.1f})"


def spread_factor(base, coll):
    """수집 스프레드가 %p 인지 bp(×100) 인지 기준이력과 비교"""
    both = pd.concat([base["스프레드"], coll["스프레드"]], axis=1, keys=["b", "c"]).dropna()
    both = both[both.b.abs() > 0.05]
    if len(both) >= 5:
        r = (both.c.abs() / both.b.abs()).median()
        diff = (both.c - both.b).abs().max()
        if abs(r - 100) < abs(r - 1):
            return 100, f"겹치는 {len(both)}일 비교, 배율 {r:.3g} → bp로 보고 ÷100"
        return 1, f"겹치는 {len(both)}일 비교, 최대 차이 {diff:.4f}%p"
    med = coll["스프레드"].dropna().abs().median()
    return (100, "크기로 판단 → bp로 보고 ÷100") if med > 5 else (1, "크기로 판단 → %p")


def merge(base, coll):
    coll = coll.copy()
    if coll["스프레드"].notna().any():
        f, note = spread_factor(base, coll)
        if f != 1:
            coll["스프레드"] = coll["스프레드"] / f
        say(f"스프레드: 수집 파일의 스프레드 열 사용 ({note})")
        both = coll[["국채3년", "국채10년", "스프레드"]].dropna()
        if len(both):
            gap = (both["스프레드"] - (both["국채10년"] - both["국채3년"])).abs().max()
            if gap > 0.002:
                say(f"⚠ 수집 파일의 스프레드가 국채10년−3년과 최대 {gap:.4f}%p 다릅니다 (스프레드 열 값을 사용)")
    else:
        say("스프레드: 수집 파일에 스프레드 열이 없어 국채10년−국채3년으로 계산")
    coll = fill_spread(coll)
    for col in ["기관", "외국인"]:
        f, note = unit_factor(base, coll, col)
        if f != 1:
            coll[col] = coll[col] / f
        say(f"{col} 단위: ÷{f:g} ({note})")
    cutoff = coll.index.max() - pd.Timedelta(days=OVERWRITE_DAYS)
    new_dates = coll.index.difference(base.index)
    recent_old = coll.index.intersection(base.index)
    recent_old = recent_old[recent_old >= cutoff]
    out = base.copy()
    if len(recent_old):
        upd = coll.loc[recent_old]
        out.update(upd)              # 비어 있는 값(NaN)은 덮어쓰지 않음
    if len(new_dates):
        out = pd.concat([out, coll.loc[new_dates]])
    out = out.sort_index()
    say(f"기준이력 {len(base)}행 → {len(out)}행 (새 날짜 {len(new_dates)}, 최근 {OVERWRITE_DAYS}일 갱신 {len(recent_old)})")
    if len(out) < len(base):
        raise StopRun("합친 결과가 기존보다 짧습니다")
    return out


# ═══════════════════════════════ 2~3. 월 묶기 + 지수 ═══════════════════════════════
def monthly(d):
    p = d.index.to_period("M")
    g = d.assign(sp=d["스프레드"].fillna(d["국채10년"] - d["국채3년"])).groupby(p)
    m = pd.DataFrame({
        "k": g["코스피"].last(),                    # 마지막 거래일 종가
        "sp": g["sp"].mean(),                       # 일별 스프레드 평균
        "기관": g["기관"].sum(min_count=1),          # 합계
        "외국인": g["외국인"].sum(min_count=1),
    })
    return m.sort_index()


def zflow(s):
    raw = s.rolling(3).sum() / s.abs().rolling(36, min_periods=12).mean()
    return (raw / raw.rolling(60, min_periods=24).std()).clip(-3, 3)


def compute(m):
    m = m.copy()
    m["sp12"] = m.sp.diff(12)
    m["z_sp"] = (m.sp12 / m.sp12.rolling(60, min_periods=24).std()).clip(-3, 3)
    m["z_ins"] = zflow(m["기관"])
    m["z_fr"] = zflow(m["외국인"])
    m["c1"] = 100 + 10 * (0.7 * m.z_sp + 0.3 * m.z_ins)
    m["c2"] = 100 + 10 * (0.6 * m.z_sp + 0.3 * m.z_ins + 0.1 * m.z_fr)
    m = m[m.c1.first_valid_index():]
    # 끝쪽에 값이 비는 달(이번 달 수급이 아직 안 들어온 경우 등)은 잘라냄
    while len(m) and (pd.isna(m.c1.iloc[-1]) or pd.isna(m.c2.iloc[-1]) or pd.isna(m.k.iloc[-1])):
        say(f"{m.index[-1]} 은 값이 비어 제외")
        m = m.iloc[:-1]
    if m[["c1", "c2", "k"]].isna().any().any():
        bad = m.index[m[["c1", "c2", "k"]].isna().any(axis=1)]
        raise StopRun("중간에 지수가 비는 달이 있습니다: " + ", ".join(str(x) for x in bad[:5]))
    if len(m) < 15:
        raise StopRun("지수 이력이 너무 짧습니다")
    # 페이지와 같은 값(소수 첫째 자리)으로 판정
    m["c1r"] = m.c1.round(1)
    m["c2r"] = m.c2.round(1)
    m["slope"] = (m.c1 - m.c1.shift(3)).round(1) + 0.0     # 명세 6장과 같은 방식(반올림 전 값의 차이), -0.0 → 0.0
    return m


def judge(v, s):
    if v >= 105:
        return "상승 우세"
    if v >= 100:
        return "약한 상승 · 동력 약화 주의" if (s is not None and not pd.isna(s) and s < 0) else "약한 상승"
    if v >= 95:
        return "하락 경고" if (s is not None and not pd.isna(s) and s < 0) else "약세"
    return "바닥권"


def warnings_at(m, i):
    c1 = m.c1r.iloc[max(0, i - 11):i + 1]
    s = m.slope.iloc[i]
    w1 = bool(c1.max() >= 105 and not pd.isna(s) and s < 0)
    w2 = bool(round(m.c2r.iloc[i] - m.c1r.iloc[i], 1) <= -3)
    return w1, w2


def is_confirmed(last_day, today):
    """마지막 데이터 날짜의 달이 끝났는가"""
    if (today.year, today.month) > (last_day.year, last_day.month):
        return True
    # 그달 마지막 평일(12월은 31일 휴장)까지 데이터가 들어왔으면 확정
    nxt = (pd.Timestamp(last_day) + pd.offsets.MonthEnd(0)).date()
    d = nxt
    while d.weekday() >= 5 or (d.month == 12 and d.day == 31):
        d -= timedelta(days=1)
    return last_day >= d


# ═══════════════════════════════ 5-2. 예측 기록 ═══════════════════════════════
REC_COLS = ["기록일", "대상월", "조합1", "조합2", "기울기", "판정", "경고①", "경고②", "그때코스피",
            "3개월뒤코스피", "수익률", "맞았나"]


def update_records(data_dir, m, confirmed_last, today):
    p = os.path.join(data_dir, RECORD_CSV)
    if os.path.exists(p):
        rec = pd.read_csv(p, encoding="utf-8-sig", dtype=str).fillna("")
        for c in REC_COLS:
            if c not in rec.columns:
                rec[c] = ""
        rec = rec[REC_COLS]
    else:
        rec = pd.DataFrame(columns=REC_COLS)
    labels = [str(x) for x in m.index]
    conf_n = len(m) if confirmed_last else len(m) - 1          # 확정된 달 개수
    changed = False
    # 새 확정 달 추가 (처음엔 가장 최근 확정 달 하나, 이후엔 빠진 달도 이어서)
    if conf_n > 0:
        done = set(rec["대상월"])
        if len(rec):
            last_rec = max(done)
            todo = [i for i in range(conf_n) if labels[i] > last_rec]
        else:
            todo = [conf_n - 1]
        for i in todo:
            if labels[i] in done:
                continue
            w1, w2 = warnings_at(m, i)
            s = m.slope.iloc[i]
            rec.loc[len(rec)] = [today.strftime("%Y-%m-%d"), labels[i], f"{m.c1r.iloc[i]:.1f}", f"{m.c2r.iloc[i]:.1f}",
                                 "" if pd.isna(s) else f"{s:.1f}", judge(m.c1r.iloc[i], s),
                                 "O" if w1 else "", "O" if w2 else "", f"{m.k.iloc[i]:.2f}", "", "", ""]
            say(f"지수기록 추가: {labels[i]} 조합1 {m.c1r.iloc[i]:.1f} {judge(m.c1r.iloc[i], s)}")
            changed = True
    # 3개월 지난 줄 결과 채우기 (그 칸만)
    pos = {lab: i for i, lab in enumerate(labels)}
    for r in rec.index:
        if rec.at[r, "3개월뒤코스피"]:
            continue
        i = pos.get(rec.at[r, "대상월"])
        if i is None or i + 3 >= conf_n:
            continue
        k0, k3 = float(rec.at[r, "그때코스피"]), float(m.k.iloc[i + 3])
        ret = (k3 / k0 - 1) * 100
        j = rec.at[r, "판정"]
        hit = ""
        if j.startswith("상승 우세") or j.startswith("약한 상승"):
            hit = "O" if ret > 0 else "X"
        elif j in ("약세", "하락 경고"):
            hit = "O" if ret < 0 else "X"
        rec.at[r, "3개월뒤코스피"] = f"{k3:.2f}"
        rec.at[r, "수익률"] = f"{ret:.2f}"
        rec.at[r, "맞았나"] = hit
        say(f"지수기록 결과: {rec.at[r, '대상월']} → 3개월 뒤 {ret:+.2f}% {hit}")
        changed = True
    return rec, changed


def record_summary(rec):
    if len(rec) < 6:
        return None
    order = ["상승 우세", "약한 상승", "약세", "하락 경고", "바닥권"]
    rows = []
    for o in order:
        sub = rec[rec["판정"].str.split(" · ").str[0] == o]
        sc = sub[sub["맞았나"].isin(["O", "X"])]
        rows.append(dict(j=o, n=int(len(sub)), done=int(len(sc)), hit=int((sc["맞았나"] == "O").sum()),
                         avg=None if not len(sub[sub["수익률"] != ""]) else round(float(sub[sub["수익률"] != ""]["수익률"].astype(float).mean()), 1)))
    allsc = rec[rec["맞았나"].isin(["O", "X"])]
    return dict(n=int(len(rec)), first=str(rec["대상월"].min()), done=int(len(allsc)),
                hit=int((allsc["맞았나"] == "O").sum()), rows=rows)


# ═══════════════════════════════ 5-1. 페이지 만들기 ═══════════════════════════════
def r1(x):
    return None if pd.isna(x) else round(float(x), 1)


def build_html(m, last_day, provisional, rec_sum):
    with open(TEMPLATE, encoding="utf-8") as f:
        tpl = f.read()
    for ph in ("__DATA__", "__CUR__", "__REC__"):
        if tpl.count(ph) != 1:
            raise StopRun(f"템플릿에 {ph} 자리가 정확히 한 번 있어야 합니다")
    D = {"lab": [str(x) for x in m.index],
         "k": [round(float(x), 2) for x in m.k],
         "c1": [r1(x) for x in m.c1],
         "c2": [r1(x) for x in m.c2],
         "s": [r1(x) for x in m.slope]}
    L = m.iloc[-1]
    CUR = {"asof": last_day.strftime("%Y-%m-%d"), "provisional": bool(provisional),
           "sp": round(float(L.sp), 4), "sp12": round(float(L.sp12), 4),
           "z_sp": round(float(L.z_sp), 3), "z_ins": round(float(L.z_ins), 3), "z_fr": round(float(L.z_fr), 3)}
    js = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":"))
    return tpl.replace("__DATA__", js(D)).replace("__CUR__", js(CUR)).replace("__REC__", js(rec_sum))


# ═══════════════════════════════ 5-3. GitHub Pages 올리기 ═══════════════════════════════
README = """# 스프레드·기관 지수

국채 스프레드 12개월 변화와 기관 수급으로 만든 100 기준 지수. 매일 19:00 자동 수집 직후 갱신됩니다.

- 주소: https://{user}.github.io/{repo}/
- 아이폰: Safari로 주소 접속 → 공유 버튼 → "홈 화면에 추가"하면 앱처럼 열립니다.
- 갱신 시각: 매일 19:00 수집 직후 + 1~2분 (GitHub 반영 시간)
- 올리기가 실패하면(네트워크 등) PC 로그에만 남고 다음 실행 때 다시 올립니다.

과거 통계 기준이며 확정 예측이 아닙니다.
"""
GITIGNORE = "*\n!index.html\n!robots.txt\n!README.md\n!.gitignore\n!.nojekyll\n"


def git(repo, *args, timeout=60, interactive=False):
    env = dict(os.environ)
    if not interactive:
        env["GIT_TERMINAL_PROMPT"] = "0"
        env["GCM_INTERACTIVE"] = "never"
    r = subprocess.run(["git", "-C", repo] + list(args), capture_output=True, timeout=timeout, env=env)
    out = (r.stdout + r.stderr).decode("utf-8", "replace").strip()
    return r.returncode, out


def publish(repo, html, today, interactive=False):
    """index.html 이 바뀌었을 때만 커밋. 아직 안 올라간 커밋이 있으면 push (실패하면 다음 실행 때 재시도)"""
    if not os.path.isdir(os.path.join(repo, ".git")):
        return "건너뜀(저장소 미설정 — 지수_처음설정.bat 을 한 번 실행하세요)"
    p = os.path.join(repo, "index.html")
    old = open(p, "rb").read() if os.path.exists(p) else b""
    new = html.encode("utf-8")
    msg = []
    if old != new:
        atomic_write(p, new)
        git(repo, "add", "index.html")
        code, out = git(repo, "commit", "-m", f"지수 갱신 {today:%Y-%m-%d}")
        if code != 0:
            return "커밋 실패: " + out[-300:]
        msg.append("커밋함")
    else:
        msg.append("내용 같음(커밋 안 함)")
    code, out = git(repo, "rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}")
    if code != 0:
        args = ["push", "-u", "origin", "HEAD:main"]
    else:
        c2, ahead = git(repo, "rev-list", "--count", "@{u}..HEAD")
        if c2 == 0 and ahead.strip() == "0":
            return ", ".join(msg) + ", 올릴 것 없음"
        args = ["push"]
    try:
        code, out = git(repo, *args, timeout=180, interactive=interactive)
    except subprocess.TimeoutExpired:
        return ", ".join(msg) + ", 올리기 시간 초과(다음 실행 때 재시도)"
    if code != 0:
        last = out.splitlines()[-1][:200] if out else ""
        return ", ".join(msg) + ", 올리기 실패(다음 실행 때 재시도): " + last
    return ", ".join(msg) + ", 올림"


def setup_repo(user, REPO_DIR):
    if not re.fullmatch(r"[A-Za-z0-9-]{1,39}", user or ""):
        raise StopRun("GitHub 아이디가 올바르지 않습니다")
    try:
        subprocess.run(["git", "--version"], capture_output=True, check=True)
    except Exception:
        raise StopRun("Git이 설치되어 있지 않습니다. https://git-scm.com/download/win 에서 설치 후 다시 실행하세요")
    os.makedirs(REPO_DIR, exist_ok=True)
    if "onedrive" in REPO_DIR.lower():
        raise StopRun("저장소 폴더를 OneDrive 밖에 두어야 합니다")
    if not os.path.isdir(os.path.join(REPO_DIR, ".git")):
        git(REPO_DIR, "init")
        git(REPO_DIR, "checkout", "-B", "main")
    git(REPO_DIR, "config", "user.name", user)
    git(REPO_DIR, "config", "user.email", f"{user}@users.noreply.github.com")
    url = f"https://github.com/{user}/{REPO_NAME}.git"
    c, _ = git(REPO_DIR, "remote", "get-url", "origin")
    git(REPO_DIR, "remote", "set-url" if c == 0 else "add", "origin", url)
    files = {"README.md": README.format(user=user, repo=REPO_NAME), "robots.txt": "User-agent: *\nDisallow: /\n",
             ".gitignore": GITIGNORE, ".nojekyll": ""}
    for name, body in files.items():
        atomic_write(os.path.join(REPO_DIR, name), body)
    git(REPO_DIR, "add", *files.keys())
    git(REPO_DIR, "commit", "-m", "처음 설정")
    say(f"저장소 폴더: {REPO_DIR}")
    say(f"연결 주소: {url}")


# ═══════════════════════════════ 화면 디자인 자동 받기 ═══════════════════════════════
def update_template():
    """GitHub 에 올려 둔 최신 디자인을 받아 템플릿을 바꿈. 실패하거나 이상한 파일이면 기존 것을 그대로 씀.
    (HTML 디자인 파일만 받음. 계산 프로그램(.py)은 자동으로 바꾸지 않음)"""
    if not TEMPLATE_URL:
        return
    try:
        import urllib.request
        req = urllib.request.Request(f"{TEMPLATE_URL}?t={int(time.time())}", headers={"Cache-Control": "no-cache"})
        with urllib.request.urlopen(req, timeout=15) as r:
            data = r.read(3_000_001)
        txt = data.decode("utf-8")
        ok = (len(data) < 3_000_000 and "</html>" in txt and "<script" in txt
              and all(txt.count(ph) == 1 for ph in ("__DATA__", "__CUR__", "__REC__")))
        if not ok:
            say("화면 디자인: 받은 파일이 올바르지 않아 기존 디자인 사용")
            return
        old = open(TEMPLATE, "rb").read() if os.path.exists(TEMPLATE) else b""
        if old != data:
            if old:
                shutil.copy2(TEMPLATE, TEMPLATE.replace(".html", "_이전.html"))
            atomic_write(TEMPLATE, data)
            say("화면 디자인: 새 디자인을 받아 적용")
    except Exception as e:
        say(f"화면 디자인: 확인 못 함, 기존 디자인 사용 ({type(e).__name__})")


# ═══════════════════════════════ 실행 ═══════════════════════════════
def run(data_dir, repo_dir, today, check=False, interactive=False):
    base, from_txt = load_base(data_dir)
    coll = load_collect(data_dir, check=check, need_from=base.index.max() - pd.Timedelta(days=14))
    hist = merge(base, coll)
    last_day = hist["코스피"].dropna().index.max().date()
    m = compute(monthly(hist))
    lm = m.index[-1]
    confirmed = is_confirmed(last_day, today) and str(lm) == f"{last_day:%Y-%m}"
    provisional = not confirmed
    i = len(m) - 1
    j = judge(m.c1r.iloc[i], m.slope.iloc[i])
    w1, w2 = warnings_at(m, i)
    summary = (f"{lm} {'잠정' if provisional else '확정'} 조합1 {m.c1r.iloc[i]:.1f} "
               f"(3개월 변화 {m.slope.iloc[i]:+.1f}) {j} / 경고① {'해당' if w1 else '아님'} / 경고② {'해당' if w2 else '아님'}"
               f" / 데이터 {last_day}")
    say(summary)
    if check:
        say("※ 점검 모드라 아무 파일도 쓰지 않았습니다")
        return summary
    update_template()
    rec, rec_changed = update_records(data_dir, m, confirmed, today)
    html = build_html(m, last_day, provisional, record_summary(rec))
    # ── 여기까지 오류 없으면 쓰기 시작 ──
    save_base(data_dir, hist)
    if rec_changed:
        atomic_write(os.path.join(data_dir, RECORD_CSV), rec.to_csv(index=False), encoding="utf-8-sig")
    atomic_write(os.path.join(data_dir, LOCAL_HTML), html)
    pub = publish(repo_dir, html, today, interactive=interactive)
    say("웹: " + pub)
    return summary


def keep_awake():
    """최대 절전에서 깨어나 실행될 때, 끝날 때까지 다시 잠들지 않게 Windows에 요청 (프로그램이 끝나면 자동 해제)"""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001)  # ES_CONTINUOUS | ES_SYSTEM_REQUIRED
        except Exception:
            pass


def main():
    keep_awake()
    args = sys.argv[1:]
    data_dir = os.environ.get("SPREAD_DATA_DIR", DATA_DIR)
    repo_dir = os.environ.get("SPREAD_REPO_DIR", REPO_DIR)
    today = date.fromisoformat(os.environ["SPREAD_TODAY"]) if os.environ.get("SPREAD_TODAY") else date.today()
    if "--setup" in args:
        user = args[args.index("--setup") + 1] if len(args) > args.index("--setup") + 1 else ""
        try:
            setup_repo(user, repo_dir)
            summary = run(data_dir, repo_dir, today, interactive=True)
            write_log(data_dir, "처음 설정", summary)
            return 0
        except Exception as e:
            print("\n오류:", e)
            if not isinstance(e, StopRun):
                traceback.print_exc()
            return 1
    check = "--check" in args
    if not os.path.isdir(data_dir):
        print("data 폴더가 없습니다:", data_dir)
        return 1
    lock = None
    try:
        if not check:
            lock = acquire_lock(data_dir)
        summary = run(data_dir, repo_dir, today, check=check)
        if not check:
            write_log(data_dir, "성공", summary)
        return 0
    except Exception as e:
        msg = str(e) if isinstance(e, StopRun) else f"{type(e).__name__}: {e}"
        if not isinstance(e, StopRun):
            _log_lines.append(traceback.format_exc().strip().splitlines()[-3:][0])
        print("오류:", msg)
        if not check:
            write_log(data_dir, "오류(이전 페이지 유지)", msg)
        return 1
    finally:
        if lock and os.path.exists(lock):
            os.remove(lock)


if __name__ == "__main__":
    sys.exit(main())
