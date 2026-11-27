"""상장 메자닌(CB·BW·EB) 발행결정 자동 수집기.

GitHub Actions에서 매일 실행된다.
  1) OpenDART 공시목록에서 발행결정 주요사항보고서를 찾고
  2) 발행결정 상세 API로 금액·이율·전환가 등 표 값을 받고
  3) 공시 원문에서 인수인·Put·Call을 추출해
  4) data/deals.json 을 갱신하고, 새 건이 있으면 메일로 요약을 보낸다.

환경변수(GitHub Secrets):
  DART_API_KEY   (필수) OpenDART 인증키
  MAIL_USER      (선택) 보내는 Gmail 주소
  MAIL_PASSWORD  (선택) Gmail 앱 비밀번호
  MAIL_TO        (선택) 받는 주소 (없으면 MAIL_USER)
  OUR_KEYWORDS   (선택) 우리 회사 조합을 찾는 키워드, 쉼표 구분 (메일에만 사용)
  SITE_URL       (선택) 페이지 주소 (메일 하단 링크용)
"""
from __future__ import annotations

import io
import json
import os
import re
import smtplib
import sys
import time
import zipfile
from datetime import date, datetime, timedelta, timezone
from email.mime.text import MIMEText
from pathlib import Path
from difflib import SequenceMatcher
from statistics import median

import requests
from bs4 import BeautifulSoup

KST = timezone(timedelta(hours=9))
API = "https://opendart.fss.or.kr/api"
ROOT = Path(__file__).parent
DATA_FILE = ROOT / "data" / "deals.json"
DATA_VERSION = 4  # 파서가 바뀌면 올림 → 다음 실행 때 처음부터 다시 수집
CONFIG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))

KINDS = {
    "CB": ("전환사채권발행결정", "cvbdIsDecsn"),
    "BW": ("신주인수권부사채권발행결정", "bdwtIsDecsn"),
    "EB": ("교환사채권발행결정", "exbdIsDecsn"),
}
MARKETS = {"Y": "유가", "K": "코스닥"}


# ───────────────────────── DART 호출 ─────────────────────────
class Dart:
    def __init__(self, key: str):
        self.key = key
        self.s = requests.Session()
        self.calls = 0

    def get(self, path: str, **params):
        """DART가 느리거나 끊겨도 최대 6번, 점점 길게 기다리며 다시 시도한다."""
        params["crtfc_key"] = self.key
        waits = [5, 15, 30, 60, 120]
        for attempt in range(len(waits) + 1):
            try:
                r = self.s.get(f"{API}/{path}", params=params, timeout=60)
                self.calls += 1
                time.sleep(0.15)
                r.raise_for_status()
                return r
            except requests.RequestException as e:
                if attempt == len(waits):
                    raise
                print(f"  재시도 {attempt + 1}/{len(waits)} {path}: {type(e).__name__} → {waits[attempt]}초 대기")
                time.sleep(waits[attempt])

    def json(self, path: str, **params) -> dict:
        d = self.get(path, **params).json()
        st = d.get("status")
        if st not in ("000", "013"):  # 013 = 조회 데이터 없음
            raise RuntimeError(f"DART {path} 오류 {st}: {d.get('message')}")
        return d

    def list_chunks(self, bgn: date, end: date):
        """주요사항보고서(B001) 목록을 3개월 단위로 끊어서 차례로 돌려준다."""
        cur = bgn
        while cur <= end:
            to = min(end, cur + timedelta(days=89))
            out, page = [], 1
            while True:
                d = self.json("list.json", bgn_de=f"{cur:%Y%m%d}", end_de=f"{to:%Y%m%d}",
                              pblntf_detail_ty="B001", page_no=page, page_count=100)
                out += d.get("list", [])
                if page >= int(d.get("total_page") or 1):
                    break
                page += 1
            yield cur, to, out
            cur = to + timedelta(days=1)

    def decision(self, api: str, corp_code: str, day: str) -> list[dict]:
        return self.json(f"{api}.json", corp_code=corp_code, bgn_de=day, end_de=day).get("list", [])

    def document_text(self, rcept_no: str) -> str:
        r = self.get("document.xml", rcept_no=rcept_no)
        try:
            z = zipfile.ZipFile(io.BytesIO(r.content))
        except zipfile.BadZipFile:
            return ""
        html = ""
        for n in z.namelist():
            raw = z.read(n)
            for enc in ("utf-8", "euc-kr", "cp949"):
                try:
                    html += raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
        return html


# ───────────────────────── 값 정리 ─────────────────────────
def num(v) -> float | None:
    if v is None:
        return None
    s = str(v).replace(",", "").replace("%", "").strip()
    if s in ("", "-", "–"):
        return None
    try:
        return float(s)
    except ValueError:
        m = re.search(r"-?\d+(\.\d+)?", s)
        return float(m.group()) if m else None


def ymd(v) -> str | None:
    """'2026년 11월 27일', '2026-11-27', '20261127' → '2026-11-27'."""
    if not v:
        return None
    m = re.search(r"(\d{4})\D{0,3}(\d{1,2})\D{0,3}(\d{1,2})", str(v))
    if not m:
        return None
    y, mo, d = map(int, m.groups())
    try:
        return date(y, mo, d).isoformat()
    except ValueError:
        return None


def pick(row: dict, *keys):
    for k in keys:
        if row.get(k) not in (None, "", "-"):
            return row[k]
    return None


def months_between(a: str | None, b: str | None) -> float | None:
    if not a or not b:
        return None
    da, db = date.fromisoformat(a), date.fromisoformat(b)
    return round((db - da).days / 365.25 * 12, 1)


# ───────────────────────── 원문 파싱 ─────────────────────────
def doc_tables(html: str) -> tuple[str, list[list[list[str]]]]:
    soup = BeautifulSoup(html, "lxml")
    tables = []
    for t in soup.find_all("table"):
        rows = []
        for tr in t.find_all("tr"):
            cells = [re.sub(r"\s+", " ", c.get_text(" ", strip=True))
                     for c in tr.find_all(["td", "th", "te", "tu"])]
            if cells:
                rows.append(cells)
        tables.append(rows)
    text = re.sub(r"[ \t　]+", " ", soup.get_text("\n"))
    return text, tables


def investors_from_tables(tables) -> list[str]:
    """'특정인에 대한 대상자별 사채발행내역' 표의 발행대상자명."""
    names = []
    for rows in tables:
        for i, r in enumerate(rows):
            joined = "".join(r).replace(" ", "")
            if "대상자명" in joined or "발행대상자" in joined:
                col = next((j for j, c in enumerate(r) if "대상자" in c.replace(" ", "")), 0)
                for r2 in rows[i + 1:]:
                    if col < len(r2):
                        n = r2[col].strip()
                        if n and n not in ("-", "합계", "계", "소계") and not n.startswith("※"):
                            names.append(n)
                break
        if names:
            break
    seen, out = set(), []
    for n in names:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out


def fund_map_from_tables(tables) -> dict[str, tuple[str, str]]:
    """'구분 | 집합투자기구 | 집합투자업자 | 신탁업자' 표 → {'1': (펀드명, 운용사)}"""
    out = {}
    for rows in tables:
        hdr = next((i for i, r in enumerate(rows) if any("집합투자기구" in c.replace(" ", "") for c in r)), None)
        if hdr is None:
            continue
        h = [c.replace(" ", "") for c in rows[hdr]]
        fi = next(j for j, c in enumerate(h) if "집합투자기구" in c)
        mi = next((j for j, c in enumerate(h) if "집합투자업자" in c or "운용" in c), None)
        for r in rows[hdr + 1:]:
            m = re.search(r"펀드\s*(\d+)", " ".join(r[:max(fi, 1)]))
            if m and fi < len(r):
                out[m.group(1)] = (r[fi].strip(), r[mi].strip() if mi is not None and mi < len(r) else "")
    return out


def norm_key(n: str) -> str:
    return re.sub(r"[\s\-·_()（）\[\]㈜]|주식회사", "", n)


def clean_gp(n: str) -> str:
    """'케이비증권㈜' → '케이비증권', '주식회사 엔엑스브이피' → '엔엑스브이피'."""
    n = re.sub(r"\(.*?\)|（.*?）", "", n)
    n = re.sub(r"㈜|\(주\)|주식회사|유한책임회사|유한회사", "", n)
    return re.sub(r"\s+", " ", n).strip(" ,·")


def gp_map_from_tables(tables) -> dict[str, list[str]]:
    """'(1) 법인 또는 단체에 관한 기본정보' 표 → {조합명(정규화): [업무집행조합원, ...]}
    행이 8칸이면 [명칭, 출자자수, 대표, 지분, 업무집행자, 지분, 최대주주, 지분],
    6칸이면 위 조합의 공동 업무집행자 줄(명칭·출자자수 칸이 합쳐져 빠짐)."""
    out: dict[str, list[str]] = {}
    for rows in tables:
        if not any("업무집행" in "".join(r).replace(" ", "") for r in rows[:2]):
            continue
        cur = None
        for r in rows:
            joined = "".join(r).replace(" ", "")
            if "명칭" in joined or joined.startswith("성명"):
                continue
            if len(r) >= 8:
                cur, gp = norm_key(r[0]), r[4]
            elif len(r) == 6 and cur:
                gp = r[2]
            else:
                continue
            gp = clean_gp(gp)
            if cur and gp and gp != "-" and gp not in out.setdefault(cur, []):
                out[cur].append(gp)
    return out


def match_key(name: str, gpmap: dict) -> str | None:
    """표마다 띄어쓰기·오타가 조금씩 달라서 정확히 같지 않아도 매우 비슷하면 같은 조합으로 본다."""
    k = norm_key(name)
    if k in gpmap:
        return k
    best = max(gpmap, key=lambda g: SequenceMatcher(None, k, g).ratio(), default=None)
    return best if best and SequenceMatcher(None, k, best).ratio() >= 0.9 else None


TRUST = re.compile(r"\(?\s*(?:본건\s*)?펀드\s*([\d,\s및]+)\s*의?\s*신탁업자[^)]*\)?")


def resolve_investors(names: list[str], fmap: dict, gpmap: dict | None = None) -> list[dict]:
    """신탁업자 명의 인수인을 실제 펀드·운용사로 바꾸고, 조합·펀드마다 운용사(gps)를 붙인다."""
    gpmap = gpmap or {}
    out = []
    for n in names:
        m = TRUST.search(n)
        if m:
            trustee = n[: m.start()].strip()
            nums = re.findall(r"\d+", m.group(1))
            hits = [fmap[k] for k in nums if k in fmap]
            if hits:
                for fund, mgr in hits:
                    gps = [clean_gp(mgr)] if mgr else [investor_group(fund)]
                    out.append({"name": fund, "type": investor_type(fund + " " + mgr), "group": " · ".join(gps),
                                "gps": gps, "gp_src": "펀드표" if mgr else "이름", "via": trustee})
                continue
            out.append({"name": n, "type": "자산운용", "group": "신탁(운용사 미확인)", "gps": ["신탁(운용사 미확인)"],
                        "gp_src": "없음", "via": trustee})
            continue
        if "신탁업자" in n or "특정금전신탁" in n:
            out.append({"name": n, "type": "자산운용", "group": "신탁(운용사 미확인)", "gps": ["신탁(운용사 미확인)"],
                        "gp_src": "없음", "via": n.split("(")[0].strip()})
            continue
        t = investor_type(n)
        m = re.search(r"\(\s*업무집행(?:사원|조합원)\s*[:：]?\s*([^)]+)\)", n)
        base = re.sub(r"\(\s*업무집행.*?\)", "", n).strip()
        if m:                                   # '○○사모투자 합자회사(업무집행사원 : △△)'
            gps, src = [clean_gp(m.group(1))], "공시"
        elif (k := match_key(base, gpmap)):     # 기본정보 표의 업무집행조합원
            gps, src = gpmap[k], "공시"
        elif t in ("신기술조합", "VC펀드", "PEF", "자산운용"):
            gps, src = [investor_group(base)], "이름"   # 표가 없으면 이름 앞머리로 추정
        else:                                   # 증권사·캐피탈·법인·개인은 본인이 곧 운용 주체
            gps, src = [clean_gp(base)], "본인"
        out.append({"name": base, "type": t, "group": " · ".join(gps), "gps": gps, "gp_src": src})
    # 같은 펀드가 두 번 나오면 하나로
    seen, uniq = set(), []
    for i in out:
        if i["name"] not in seen:
            seen.add(i["name"])
            uniq.append(i)
    return uniq


PUT_PAT = r"조기\s*상환\s*청구권|Put\s*Option"
CALL_PAT = r"매도\s*청구권|Call\s*Option"


def windows(text: str, pat: str, other: str, length: int) -> list[str]:
    """pat이 나오는 모든 위치의 본문 조각. 다음 옵션 제목(other)이 나오면 거기서 자른다."""
    out = []
    for m in re.finditer(pat, text):
        w = text[m.start(): m.start() + length]
        cut = re.search(other, w[20:])
        out.append(w[: cut.start() + 20] if cut else w)
    return out


def none_stated(chunk: str) -> bool:
    return bool(re.search(r"해당\s*사항\s*(이\s*)?없|해당없음|미부여", chunk[:160]))


def period_months(w: str, pay_date: str | None) -> int | None:
    """'발행일로부터 2년 6개월' → 30, '발행일로부터 18개월' → 18, '2년' → 24, 날짜 → 납입일과의 차이."""
    m = re.search(r"발행일\s*(?:로부터|부터|이후)\s*(?:(\d{1,2})\s*년)?\s*(?:(\d{1,3})\s*개월)?", w)
    if m and (m.group(1) or m.group(2)):
        return int(m.group(1) or 0) * 12 + int(m.group(2) or 0)
    m = re.search(r"(?:(\d{1,2})\s*년\s*)?(\d{1,3})\s*개월|(\d{1,2})\s*년\s*(?:이|이\s*되는|경과)", w)
    if m:
        if m.group(3):
            return int(m.group(3)) * 12
        return int(m.group(1) or 0) * 12 + int(m.group(2))
    d = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", w)
    if d and pay_date and ymd(d.group()) and ymd(d.group()) > pay_date:
        return round(months_between(pay_date, ymd(d.group())))
    return None


def find_ytc(chunks: list[str]) -> float | None:
    """콜옵션 매매대금 수익률. '수익률' 가까이 있는 '연 X%'를 우선, 지연배상 이율은 제외."""
    best = None
    for w in chunks:
        for m in re.finditer(r"연\s*(?:복리|단리)?\s*(\d{1,2}(?:\.\d+)?)\s*%", w):
            before, after = w[max(0, m.start() - 40): m.start()], w[m.end(): m.end() + 30]
            if re.search(r"지연|연체|배상", before + after):
                continue
            v = float(m.group(1))
            if "수익률" in before[-25:]:
                return v
            if best is None:
                best = v
    return best


def parse_options(text: str, pay_date: str | None) -> dict:
    """원문 서술에서 Put/Call을 규칙으로 추정. (정확하지 않을 수 있음 → 페이지에 추정값으로 표시)
    9-1항 요약과 '23. 기타 투자판단에 참고할 사항'의 상세가 모두 있으므로 모든 등장 위치를 본다."""
    o = {"put": None, "put_start_m": None, "call": None, "call_pct": None, "call_ytc": None}

    puts = windows(text, PUT_PAT, CALL_PAT, 1500)
    if text and not puts:
        o["put"] = False
    if puts:
        o["put"] = any(not none_stated(w) for w in puts)
        for w in puts:
            if not none_stated(w) and (v := period_months(w, pay_date)) is not None:
                o["put_start_m"] = v
                break

    calls = windows(text, CALL_PAT, PUT_PAT, 4000)
    if text and not calls:
        o["call"] = False
    if calls:
        live = [w for w in calls if not none_stated(w)]
        o["call"] = bool(live)
        for w in live:
            N = r"(?<![\d,.])(\d{1,3}(?:\.\d+)?)\s*%"
            m = re.search(r"(?:권면|전자등록|발행)\s*(?:총액|금액|총금액|가액)\s*의?\s*" + N, w) or \
                re.search(r"(?:Call\s*option|콜\s*옵션|매도청구권)\s*\(?\s*" + N, w, re.I) or \
                re.search(r"취득\s*규모[^%]{0,60}?" + N, w) or \
                re.search(r"(?:권면|전자등록)\s*(?:총액|금액|총금액)?[^%]{0,60}?" + N, w)
            if m and 0 < float(m.group(1)) <= 100:
                o["call_pct"] = float(m.group(1))
                break
        o["call_ytc"] = find_ytc(live)
    return o


# ───────────────────────── 인수인 분류 ─────────────────────────
def investor_type(name: str) -> str:
    n = name.replace(" ", "")
    if "신기술" in n:
        return "신기술조합"
    if re.search(r"투자신탁|집합투자|헤지", n):
        return "자산운용"
    if re.search(r"벤처|창업|투자조합|VC", n):
        return "VC펀드"
    if re.search(r"사모투자|PEF|합자회사", n):
        return "PEF"
    if "증권" in n:
        return "증권사"
    if "운용" in n or "신탁" in n:
        return "자산운용"
    if "캐피탈" in n:
        return "캐피탈"
    if "은행" in n or "저축" in n:
        return "은행"
    if re.search(r"주식회사|\(주\)|㈜|조합|펀드|회사", n):
        return "기타법인"
    return "개인/기타"


GROUP_CUT = re.compile(
    r"(신기술사업투자조합|신기술투자조합|신기술조합|신기술|사모투자|벤처투자조합|벤처펀드|창업|"
    r"투자조합|자산운용|증권|캐피탈|저축은행|펀드|제?\s*\d+\s*호|\(주\)|주식회사|㈜)")


def investor_group(name: str) -> str:
    """'대신 신기술투자조합 제21호' → '대신'. config.json의 aliases로 보정."""
    for alias, kws in CONFIG.get("group_aliases", {}).items():
        if any(k in name for k in kws):
            return alias
    base = re.sub(r"^\s*(주식회사|\(주\)|㈜)\s*", "", name)
    head = GROUP_CUT.split(base, maxsplit=1)[0]
    head = re.sub(r"[\s\-·_()]+$", "", head).strip(" -·")
    head = re.sub(r"^\(?주\)?", "", head).strip()
    return head or name[:8]


# ───────────────────────── 한 건 처리 ─────────────────────────
def build_deal(dart: Dart, rep: dict, kind: str) -> dict | None:
    api = KINDS[kind][1]
    rows = dart.decision(api, rep["corp_code"], rep["rcept_dt"])
    row = next((r for r in rows if r.get("rcept_no") == rep["rcept_no"]), rows[-1] if rows else None)
    if not row:
        print(f"  상세 없음: {rep['corp_name']} {rep['rcept_no']}")
        return None

    amount = num(pick(row, "bd_fta"))
    price = num(pick(row, "cv_prc", "ex_prc", "exrqpd_prc"))
    floor = num(pick(row, "act_mktprcfl_cvprc_lwtrsprc", "act_mktprcfl_exprc_lwtrsprc"))
    pay = ymd(pick(row, "pymd"))
    mat = ymd(pick(row, "bd_mtd"))

    html = dart.document_text(rep["rcept_no"])
    text, tables = doc_tables(html) if html else ("", [])
    invs = resolve_investors(investors_from_tables(tables), fund_map_from_tables(tables), gp_map_from_tables(tables))
    opts = parse_options(text, pay)

    refix_floor = round(floor / price * 100, 1) if floor and price and floor < price else None
    funds = {k: num(row.get(f)) for k, f in
             [("시설", "fdpp_fclt"), ("영업양수", "fdpp_bsninh"), ("운영", "fdpp_op"),
              ("채무상환", "fdpp_dtrp"), ("타법인증권", "fdpp_ocsa"), ("기타", "fdpp_etc")]}
    use = [k for k, v in funds.items() if v]

    return {
        "key": f"{rep['corp_code']}-{kind}-{re.sub(r'[^0-9A-Za-z]', '', str(row.get('bd_tm') or rep['rcept_no']))}",
        "rcept_no": rep["rcept_no"],
        "first_dt": date.fromisoformat(f"{rep['rcept_dt'][:4]}-{rep['rcept_dt'][4:6]}-{rep['rcept_dt'][6:]}").isoformat(),
        "corp": rep["corp_name"],
        "corp_code": rep["corp_code"],
        "market": MARKETS.get(rep.get("corp_cls"), rep.get("corp_cls")),
        "kind": kind,
        "round": row.get("bd_tm"),
        "amount_eok": round(amount / 1e8, 1) if amount else None,
        "coupon": num(row.get("bd_intr_ex")),
        "ytm": num(row.get("bd_intr_sf")),
        "maturity": mat,
        "tenor_y": round(months_between(pay, mat) / 12, 1) if pay and mat else None,
        "price": price,
        "refix_floor": refix_floor,
        "pay_date": pay,
        "use": use,
        "method": row.get("bdis_mthn"),
        "investors": invs,
        **opts,
        "est": ["put", "put_start_m", "call", "call_pct", "call_ytc"],
        "corrections": 0,
        "withdrawn": False,
    }


def kind_of(report_nm: str) -> str | None:
    n = report_nm.replace(" ", "")
    for k, (title, _) in KINDS.items():
        if title in n:
            return k
    return None


# ───────────────────────── 메일 ─────────────────────────
def flags(d: dict, rules: dict) -> list[str]:
    f = []
    if d.get("call_pct") is not None and d["call_pct"] >= rules["call_pct_min"]:
        f.append(f"Call {d['call_pct']:g}%")
    if d["kind"] != "EB" and d.get("refix_floor") is None:
        f.append("리픽싱 없음")
    if d.get("ytm") is not None and d["ytm"] < rules["ytm_min"]:
        f.append(f"YTM {d['ytm']:g}%")
    if d.get("put_start_m") and d["put_start_m"] > rules["put_start_max_m"]:
        f.append(f"Put {d['put_start_m']}M~")
    if d.get("tenor_y") and d["tenor_y"] >= 30:
        f.append("영구채형")
    return f


def send_mail(new: list[dict], all_deals: list[dict]):
    user, pw = os.getenv("MAIL_USER"), os.getenv("MAIL_PASSWORD")
    if not (user and pw) or not new:
        return
    to = os.getenv("MAIL_TO") or user
    rules = CONFIG["flag_rules"]
    ours = [k.strip() for k in os.getenv("OUR_KEYWORDS", "").split(",") if k.strip()]
    site = os.getenv("SITE_URL", "")

    def line(d):
        inv = ", ".join(i["name"] for i in d["investors"][:3]) + (f" 외 {len(d['investors'])-3}" if len(d["investors"]) > 3 else "")
        tag = "  ⚠ " + " · ".join(flags(d, rules)) if flags(d, rules) else ""
        me = "  ★우리 회사" if ours and any(k in i["name"] for i in d["investors"] for k in ours) else ""
        price = f"{d['price']:,.0f}" if d.get("price") else "-"
        refix = f"{d['refix_floor']:g}%" if d.get("refix_floor") else "없음"
        return (f"• {d['corp']} ({d['market']}) {d['kind']} {d['round']}회 {d['amount_eok'] or '-'}억 | "
                f"{d['coupon']}/{d['ytm']}% | 전환가 {price}원 | 리픽싱 {refix}\n"
                f"  인수: {inv or '미확인'}{tag}{me}\n"
                f"  https://dart.fss.or.kr/dsaf001/main.do?rcpNo={d['rcept_no']}")

    body = f"새 메자닌 발행결정 {len(new)}건\n\n" + "\n\n".join(line(d) for d in new)
    body += (f"\n\n전체 현황: {site}" if site else "") + "\n\n(Put·Call 값은 원문에서 자동 추정한 값이라 틀릴 수 있습니다)"
    msg = MIMEText(body, _charset="utf-8")
    msg["Subject"] = f"[메자닌] {datetime.now(KST):%m/%d} 신규 {len(new)}건"
    msg["From"], msg["To"] = user, to
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(user, pw)
        s.sendmail(user, [a.strip() for a in to.split(",")], msg.as_string())
    print(f"메일 발송: {to}")


# ───────────────────────── 메인 ─────────────────────────
def save(store: dict, deals: dict, seen: set, backfill_done: bool):
    store.update({"version": DATA_VERSION, "updated": datetime.now(KST).strftime("%Y-%m-%d %H:%M"),
                  "backfill_done": backfill_done, "deals": deals, "seen": sorted(seen)})
    DATA_FILE.parent.mkdir(exist_ok=True)
    DATA_FILE.write_text(json.dumps(store, ensure_ascii=False, indent=0), encoding="utf-8")


def main():
    key = os.getenv("DART_API_KEY")
    if not key:
        sys.exit("DART_API_KEY 가 없습니다 (GitHub Secrets 확인)")
    dart = Dart(key)

    store = json.loads(DATA_FILE.read_text(encoding="utf-8")) if DATA_FILE.exists() else {}
    if store.get("version") != DATA_VERSION:
        print("데이터 형식이 바뀌어 처음부터 다시 수집합니다")
        store = {}
    deals: dict = store.get("deals", {})
    seen = set(store.get("seen", []))
    backfill_done = store.get("backfill_done", False)

    start = date.fromisoformat(CONFIG["start_date"])
    today = datetime.now(KST).date()
    # 처음 수집(또는 중간에 끊긴 수집)은 시작일부터, 그 뒤로는 최근 며칠만
    bgn = max(start, today - timedelta(days=CONFIG.get("lookback_days", 10))) if backfill_done else start
    print(f"목록 조회 {bgn} ~ {today} ({'일일 갱신' if backfill_done else '전체 수집'})")

    new_keys, failed = [], 0
    try:
        for cur, to, chunk in dart.list_chunks(bgn, today):
            reports = sorted((r for r in chunk if r.get("corp_cls") in MARKETS and kind_of(r["report_nm"])),
                             key=lambda r: r["rcept_no"])
            todo = [r for r in reports if r["rcept_no"] not in seen]
            print(f"[{cur} ~ {to}] 발행결정 {len(reports)}건, 새로 처리 {len(todo)}건")
            for rep in todo:
                kind = kind_of(rep["report_nm"])
                try:
                    d = build_deal(dart, rep, kind)
                except Exception as e:  # 한 건 실패가 전체를 멈추지 않도록 (다음 실행 때 다시 시도)
                    failed += 1
                    print(f"  실패 {rep['corp_name']} {rep['rcept_no']}: {type(e).__name__} {e}")
                    continue
                seen.add(rep["rcept_no"])
                if not d:
                    continue
                old = deals.get(d["key"])
                if old:  # 정정공시 → 최신 조건으로 덮되 최초 공시일·정정횟수 유지
                    d["first_dt"] = old["first_dt"]
                    d["corrections"] = old.get("corrections", 0) + 1
                elif backfill_done:
                    new_keys.append(d["key"])
                deals[d["key"]] = d
                print(f"  {'정정' if old else '신규'} {d['corp']} {kind} {d['round']}회 {d['amount_eok']}억 인수인 {len(d['investors'])}")
            save(store, deals, seen, backfill_done)  # 구간마다 저장 → 중간에 끊겨도 여기까지는 남는다
        backfill_done = True
    finally:
        save(store, deals, seen, backfill_done)
        print(f"저장 {len(deals)}건, 실패 {failed}건, DART 호출 {dart.calls}회, 전체 수집 {'완료' if backfill_done else '미완료(다음 실행 때 이어서)'}")

    send_mail([deals[k] for k in new_keys], list(deals.values()))


if __name__ == "__main__":
    main()
