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
from statistics import median

import requests
from bs4 import BeautifulSoup

KST = timezone(timedelta(hours=9))
API = "https://opendart.fss.or.kr/api"
ROOT = Path(__file__).parent
DATA_FILE = ROOT / "data" / "deals.json"
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
        params["crtfc_key"] = self.key
        for attempt in range(3):
            try:
                r = self.s.get(f"{API}/{path}", params=params, timeout=30)
                self.calls += 1
                time.sleep(0.15)
                r.raise_for_status()
                return r
            except requests.RequestException as e:
                if attempt == 2:
                    raise
                print(f"  재시도 {path}: {e}")
                time.sleep(2 + attempt * 3)

    def json(self, path: str, **params) -> dict:
        d = self.get(path, **params).json()
        st = d.get("status")
        if st not in ("000", "013"):  # 013 = 조회 데이터 없음
            raise RuntimeError(f"DART {path} 오류 {st}: {d.get('message')}")
        return d

    def list_reports(self, bgn: date, end: date) -> list[dict]:
        """주요사항보고서 목록 (3개월 단위로 끊어 조회)."""
        out, cur = [], bgn
        while cur <= end:
            to = min(end, cur + timedelta(days=89))
            page = 1
            while True:
                d = self.json("list.json", bgn_de=f"{cur:%Y%m%d}", end_de=f"{to:%Y%m%d}",
                              pblntf_ty="B", page_no=page, page_count=100)
                out += d.get("list", [])
                if page >= int(d.get("total_page") or 1):
                    break
                page += 1
            cur = to + timedelta(days=1)
        return out

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
    for t in soup.find_all(["table", "TABLE"]):
        rows = []
        for tr in t.find_all(["tr", "TR"]):
            cells = [re.sub(r"\s+", " ", c.get_text(" ", strip=True))
                     for c in tr.find_all(["td", "th", "TD", "TH", "TE", "TU"])]
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


def parse_options(text: str, pay_date: str | None) -> dict:
    """원문 서술에서 Put/Call을 규칙으로 추정. (정확하지 않을 수 있음 → 페이지에 추정값으로 표시)
    본문 9-1항은 '23. 기타 투자판단에 참고할 사항 참조'처럼 안내만 있는 경우가 많아
    모든 등장 위치를 보고 숫자가 들어 있는 조각을 우선한다."""
    o = {"put": None, "put_start_m": None, "call": None, "call_pct": None, "call_ytc": None}

    puts = windows(text, PUT_PAT, CALL_PAT, 1500)
    if puts:
        o["put"] = any(not none_stated(w) for w in puts)
        for w in puts:
            if none_stated(w):
                continue
            m = re.search(r"(\d{1,3})\s*개월", w)
            if m:
                o["put_start_m"] = int(m.group(1))
                break
            d = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", w)
            if d and pay_date and ymd(d.group()) and ymd(d.group()) > pay_date:
                o["put_start_m"] = round(months_between(pay_date, ymd(d.group())))
                break

    calls = windows(text, CALL_PAT, PUT_PAT, 2000)
    if calls:
        o["call"] = any(not none_stated(w) for w in calls)
        for w in calls:
            if none_stated(w):
                continue
            m = re.search(r"(?:권면|전자등록|발행)\s*(?:총액|금액|총금액)?[^%]{0,60}?(\d{1,3}(?:\.\d+)?)\s*%", w)
            if m and 0 < float(m.group(1)) <= 100:
                o["call_pct"] = float(m.group(1))
                y = re.search(r"연\s*(?:복리|단리)?\s*(\d{1,2}(?:\.\d+)?)\s*%", w)
                if y:
                    o["call_ytc"] = float(y.group(1))
                break
    return o


# ───────────────────────── 인수인 분류 ─────────────────────────
def investor_type(name: str) -> str:
    n = name.replace(" ", "")
    if "신기술" in n:
        return "신기술조합"
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
    invs = investors_from_tables(tables)
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
        "investors": [{"name": n, "type": investor_type(n), "group": investor_group(n)} for n in invs],
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
    site = CONFIG.get("site_url", "")

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
    body += f"\n\n전체 현황: {site}\n(Put·Call 값은 원문에서 자동 추정한 값이라 틀릴 수 있습니다)"
    msg = MIMEText(body, _charset="utf-8")
    msg["Subject"] = f"[메자닌] {datetime.now(KST):%m/%d} 신규 {len(new)}건"
    msg["From"], msg["To"] = user, to
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(user, pw)
        s.sendmail(user, [a.strip() for a in to.split(",")], msg.as_string())
    print(f"메일 발송: {to}")


# ───────────────────────── 메인 ─────────────────────────
def main():
    key = os.getenv("DART_API_KEY")
    if not key:
        sys.exit("DART_API_KEY 가 없습니다 (GitHub Secrets 확인)")
    dart = Dart(key)

    store = json.loads(DATA_FILE.read_text(encoding="utf-8")) if DATA_FILE.exists() else {"deals": {}, "seen": []}
    deals: dict = store["deals"]
    seen = set(store["seen"])

    start = date.fromisoformat(CONFIG["start_date"])
    today = datetime.now(KST).date()
    bgn = start if not seen else max(start, today - timedelta(days=CONFIG.get("lookback_days", 10)))
    print(f"목록 조회 {bgn} ~ {today}")

    reports = [r for r in dart.list_reports(bgn, today)
               if r.get("corp_cls") in MARKETS and kind_of(r["report_nm"])]
    reports.sort(key=lambda r: r["rcept_no"])
    print(f"발행결정 공시 {len(reports)}건, 신규 {sum(r['rcept_no'] not in seen for r in reports)}건")

    new_keys = []
    for rep in reports:
        if rep["rcept_no"] in seen:
            continue
        kind = kind_of(rep["report_nm"])
        try:
            d = build_deal(dart, rep, kind)
        except Exception as e:  # 한 건 실패가 전체를 멈추지 않도록
            print(f"  실패 {rep['corp_name']} {rep['rcept_no']}: {e}")
            continue
        seen.add(rep["rcept_no"])
        if not d:
            continue
        old = deals.get(d["key"])
        if old and "정정" not in rep["report_nm"]:  # 정정이 아닌데 키가 겹치면 별개 건으로
            d["key"] += "-" + rep["rcept_no"]
            old = None
        if old:  # 정정공시 → 최신 조건으로 덮되 최초 공시일·정정횟수 유지
            d["first_dt"] = old["first_dt"]
            d["corrections"] = old.get("corrections", 0) + 1
        elif bgn != start:
            new_keys.append(d["key"])
        deals[d["key"]] = d
        print(f"  {'정정' if old else '신규'} {d['corp']} {kind} {d['round']}회 {d['amount_eok']}억 인수인 {len(d['investors'])}")

    store = {"updated": datetime.now(KST).strftime("%Y-%m-%d %H:%M"), "deals": deals, "seen": sorted(seen)}
    DATA_FILE.parent.mkdir(exist_ok=True)
    DATA_FILE.write_text(json.dumps(store, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"저장 {len(deals)}건, DART 호출 {dart.calls}회")

    send_mail([deals[k] for k in new_keys], list(deals.values()))


if __name__ == "__main__":
    main()
