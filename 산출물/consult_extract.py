"""관련부서 협의 회신 문서(odt/pdf/hwp) -> 협의의견/조치계획 표 데이터(JSON)

사용: python -I consult_extract.py <협의폴더> <출력.json> [<부서명_보정.json>]
- 별도의견 없음만 있는 부서는 제외, 가나다순 정렬.
- 조치계획은 규칙 기반 '초안'이다(draft=True). 반드시 사람이 확인한다.
- warnings: 의견 본문의 지명(리)이 다른 회신과 다르면 경고(원문 오기 가능성).
"""
import sys, os, re, io, json, html, zipfile, zlib, struct
from collections import Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ---------- 문서 -> 줄 목록 ----------
def odt_lines(p):
    x = zipfile.ZipFile(p).read("content.xml").decode("utf8")
    x = re.sub(r"<text:line-break/>", "\n", x)
    x = re.sub(r"</text:p>|</text:h>", "\n", x)
    x = re.sub(r"</table:table-cell>", "\n⟦CELL⟧\n", x)
    x = re.sub(r"<[^>]+>", "", x)
    return [l.strip() for l in html.unescape(x).split("\n")]

def pdf_lines(p):
    import pymupdf
    return [l.strip() for pg in pymupdf.open(p) for l in pg.get_text().split("\n")]

def hwp_lines(path):
    import olefile
    ole = olefile.OleFileIO(path)
    comp = bool(ole.openstream("FileHeader").read()[36] & 1)
    out = []
    for s in sorted([s for s in ole.listdir() if s[0] == "BodyText"], key=lambda s: int(s[1][7:])):
        d = ole.openstream(s).read()
        if comp: d = zlib.decompress(d, -15)
        i = 0
        while i + 4 <= len(d):
            h = struct.unpack_from("<I", d, i)[0]
            tag, size = h & 0x3FF, (h >> 20) & 0xFFF
            i += 4
            if size == 0xFFF: size = struct.unpack_from("<I", d, i)[0]; i += 4
            if tag == 67:
                out.append("".join(c for c in d[i:i + size].decode("utf-16-le", "ignore") if c >= " "))
            i += size
    return out

def read_lines(p):
    e = os.path.splitext(p)[1].lower()
    return {".odt": odt_lines, ".pdf": pdf_lines, ".hwp": hwp_lines}.get(e, lambda _: None)(p)

# ---------- 분류/추출 ----------
SKIP_TITLE = re.compile(r"중간\s*회신|보완\s*요청|부의안|서식|조건|작성요령")
REQUESTERS = ("민원담당관", "개발행위팀", "시장")
STOP = re.compile(r"^(주무관|협조자|시행|전화번호|우$|기안|검토자)")
SIGN = re.compile(r"[가-힣\s]{2,16}(과|팀|센터|소|서)\s*장\s*$")   # 말미 결재자 직함

def norm(s): return re.sub(r"\s+", "", s)
def clean(l): return l.replace("⟦CELL⟧", "").strip()

def title_of(lines):
    for i, l in enumerate(lines):
        if norm(l) == "제목":
            for n in lines[i + 1:i + 6]:
                if n and n != "⟦CELL⟧": return n
    return ""

def dept_of(lines, fname, overrides):
    stem = os.path.splitext(fname)[0]
    for k, v in overrides.items():
        if k in stem: return v
    joined = "\n".join(norm(l) for l in lines)
    for m in re.finditer(r"([가-힣]{2,10}(?:과|팀|센터|소|서))장", joined):
        d = m.group(1)
        if not any(r in d for r in REQUESTERS) and d != "개발행위허가": return d
    return stem

SPACING = [   # 원문 붙여쓰기 오류(자주 나오는 것만). 기존 PPT 작성 시 사람이 손으로 고치던 항목.
    (r"(제\d+조(?:의\d+)?)(및)", r"\1 \2 "), (r"(제\d+조)(심사|협의|규정)", r"\1 \2"), (r"에따라", "에 따라"),
    (r"에저촉", "에 저촉"), (r"을허가", "을 허가"), (r"를허가", "를 허가"), (r"히이행", "히 이행"), (r"시행령제(\d)", r"시행령 제\1"),
    (r"에의하여", "에 의하여"), (r"에의해", "에 의해"), (r"에대하여", "에 대하여"), (r"에대한", "에 대한"), (r"와인접", "와 인접"),
    (r"에의하여시행", "에 의하여 시행"), (r"를도급", "를 도급"), (r"을도급", "을 도급"), (r"로부터최초", "로부터 최초"),
]
FIXES = {}     # 사용자 지정 치환(예: 원문 지명 오기) — 보정 JSON의 "_fixes"

def respace(p):
    for a, b in SPACING: p = re.sub(a, b, p)
    for a, b in FIXES.items(): p = p.replace(a, b)
    return re.sub(r"\s{2,}", " ", p)

def tidy(p):
    p = respace(p)
    p = re.sub(r"\s+", " ", p).strip()
    p = re.sub(r"\s붙임\s+[^.]*?\d+\s*부\.?(?:\s*끝\.?)?$", "", p)   # 말미 붙임 목록
    if re.match(r"^붙임", p): return ""
    p = re.sub(r"\s*(끝\.?\s*)+$", "", p)                 # 끝.
    p = SIGN.sub("", p)                                   # 결재자 직함
    p = re.sub(r"\s*(끝\.?\s*)+$", "", p).strip()
    return p

def body_items(lines):
    L = [clean(l) for l in lines]
    start = next((i for i, l in enumerate(L) if norm(l) == "제목"), 0) + 1
    while start < len(L) and not L[start]: start += 1
    start += 1
    txt = []
    for l in L[start:]:
        if not l: continue
        if STOP.match(l): break
        txt.append(l)
    j = "\n".join(txt)

    # (a) 표 형식: [관련법 | 검토의견 | 검토자] -> 검토의견 셀(가장 긴 문단). STOP(검토자) 이전에서 잘리므로 전체 줄 사용
    full = "\n".join(l for l in L[start:] if l)
    if "검토의견" in full and "검토자" in full:
        m = re.search(r"검토자\s*\n(.*?)(?:\n[☐□]|\Z)", full, re.S)
        if m:
            cands = [c for c in m.group(1).split("\n") if len(c) > 60]
            if cands: return [tidy(max(cands, key=len))], "table"
    # (b) 표 형식: 검토결과 + "-" 항목
    m = re.search(r"검토결과\s*\n(.*?)(?:\n끝\.|\Z)", j, re.S)
    if m:
        items = [re.sub(r"^[-○]\s*", "", s).strip() for s in re.split(r"\n(?=[-○])", m.group(1)) if s.strip()]
        return [tidy(s) for s in items if tidy(s)], "table"

    # (c) 번호 공문
    paras, cur = [], None
    for l in txt:
        if re.match(r"^\d+\.\s", l):
            if cur is not None: paras.append(cur)
            cur = [l]
        elif cur is not None: cur.append(l)
        else: cur = [l]
    if cur: paras.append(cur)
    out = []
    for p in paras:
        q = re.sub(r"^\d+\.\s*", "", " ".join(p)).strip()
        q = re.split(r"[❍☐□]\s", q)[0]                                  # 내역표(❍/☐) 이후 절단
        q = tidy(q)
        if not q: continue
        if re.search(r"\d+\s*부\.?$", q) and len(q) < 60: continue          # "…결정서 등 각 1부." 첨부 목록
        if re.search(r"호\s*(와\s*)?관련(입니다|됩니다)\.?\s*$", q) and len(q) < 60: continue
        if re.fullmatch(r".{0,45}호(와|에)?\s*관련(하여)?,?\s*(아래와|다음과)\s*같이\s*회신합니다\.?", q): continue
        q = re.sub(r"^[^,.]{0,45}?호(와|에)?\s*관련(하여)?,\s*", "", q)
        q = re.sub(r"^[^,.]{0,45}?호\s*관련\s*(다음과|아래와)\s*같이\s*회신합니다\.?\s*", "", q)
        if re.fullmatch(r"[가-힣()\s]{2,30}\s*회신", q): continue        # 제목성 문단 ("산지전용허가(협의) 회신")
        q = re.sub(r"^[가-힣()\s]{2,30}\s*회신\s+(?=[가-하]\.)", "", q)      # "…회신 가. 본문" -> "가. 본문"
        q = tidy(q)
        if re.match(r"^가\.\s", q) and not re.search(r"\s나\.\s", q): q = re.sub(r"^가\.\s*", "", q)
        if q: out.append(q)
    return out, "gongmun"

def is_no_opinion(items):
    t = " ".join(items)
    if not re.search(r"별도\s*의견\s*없음|이의\s*없", t): return False
    rest = re.sub(r"별도\s*의견\s*없음[^.]*\.", "", t)
    return not re.search(r"다만|단,|하여야|바랍니다|안내|제출|준수|득하|허가를\s*받", rest)

# ---------- 조치계획 초안 ----------
DIRECTIVE = re.compile(r"바랍니다|하여야|해야|받아야|신고|제출|준수|이행|안내|않도록|하시기|알려드리니")

def split_sents(item):
    # "가. ... 나. ..." / 문장 단위 분리
    parts = re.split(r"(?:(?<=[.。])\s+|\s+(?=[가-하]\.\s))", item)
    return [re.sub(r"^[가-하]\.\s*", "", s).strip() for s in parts if s.strip()]

def convert(s, dept):
    s = re.sub(r"^(단|다만),?\s*", "", s).strip()
    s = re.sub(r"^.*저촉\s*사항이\s*없어,\s*(붙임\s*[가-힣\s]*?(을|를)\s*)?", "", s)
    m = re.search(r"(하오니|하므로|바라며),\s*(.+)$", s)            # '…허가(협의)하오니, 이행사항을 준수하여 주시기 바랍니다' -> 뒤 절만
    if m and DIRECTIVE.search(m.group(2)) and len(m.group(2)) > 8:
        tail = m.group(2)
        s = ("붙임 조건 " + tail) if (re.search(r"붙임\s*조건", s[:m.start()]) and re.match(r"(이행|유의)", tail)) else tail
    s = s.replace("우리과", dept).replace("우리 과", dept).replace("우리부서", dept)
    # '…함을 (민원인에게) 안내하여 주시기 바랍니다' : 안내 대상은 신청인 -> 문장 유지 + 안내 하겠음
    s = re.sub(r"(함|음)을\s*알려\s*드리니,?\s*민원인에게\s*안내하여\s*주시기\s*바랍니다\.?", r"\1을 민원인에게 안내 하겠음", s)
    s = re.sub(r"(함|음)을\s*안내하여\s*주시기\s*바랍니다\.?", r"\1을 안내하겠음", s)
    s = re.sub(r"(토록)\s*안내하여\s*주시기\s*바랍니다\.?", r"\1 안내하겠음", s)
    s = re.sub(r"안내하여\s*주시기\s*바랍니다\.?", "안내 하겠음", s)
    if not re.search(r"안내 ?하겠음|안내하겠음", s):
        s = re.sub(r"(준수|이행|제출|신고|신청)하여\s*주시기\s*바랍니다\.?", r"\1하겠음", s)
        s = re.sub(r"(제출|신고|신청|이행|준수)하여야\s*(합니다|함)\.?", r"\1하겠음", s)
        s = re.sub(r"(받아야|하여야)\s*(합니다|함)\.?", lambda m: ("받겠음" if m.group(1) == "받아야" else "하겠음"), s)
        s = re.sub(r"하시기\s*바랍니다\.?", "하겠음", s)
    s = re.sub(r"않도록\s*하겠음", "않도록 하겠음", s)
    return s.rstrip(". ")

def action_draft(item, dept):
    outs = []
    for s in split_sents(item):
        if re.search(r"저촉\s*사항\s*없음\.?$", s): continue
        if re.search(r"별도\s*의견\s*없음", s) and not DIRECTIVE.search(re.sub(r"별도\s*의견\s*없음[^.]*", "", s)): continue
        if not DIRECTIVE.search(s): continue                     # 단순 안내·사실 서술은 조치 대상 아님
        outs.append(convert(s, dept))
    return outs

def main(folder, out_json, ov_path=None):
    overrides = {"농지": "농지전용", "산지": "산지전용", "문화": "문화예술과", "축산": "축산과", "토지": "토지정보과",
                 "환경": "환경정책과", "건설": "건설과", "기후": "기후에너지과", "동면": "동면행정복지센터"}
    if ov_path:
        ov = json.load(open(ov_path, encoding="utf-8-sig"))
        FIXES.update(ov.pop("_fixes", {}))
        overrides.update(ov)
    rows, skipped = [], []
    for f in sorted(os.listdir(folder)):
        p = os.path.join(folder, f)
        if os.path.splitext(f)[1].lower() not in (".odt", ".pdf", ".hwp"): continue
        try: lines = read_lines(p)
        except Exception as ex: skipped.append((f, f"읽기 실패: {ex}")); continue
        if not lines: continue
        title = title_of(lines)
        if not title or "회신" not in title or SKIP_TITLE.search(title):
            skipped.append((f, f"회신 문서 아님: {title[:30]}")); continue
        dept = dept_of(lines, f, overrides)
        items, kind = body_items(lines)
        if not items: skipped.append((f, "본문 추출 실패")); continue
        if is_no_opinion(items): skipped.append((f, "별도의견 없음 -> 제외")); continue
        act = []
        for it in items: act += action_draft(it, dept)
        rows.append({"dept": dept, "source": f, "kind": kind, "opinion": items, "action": act, "draft": True})
    rows.sort(key=lambda r: r["dept"])
    # 지명 일관성 경고
    PLACE = r"(?<=면\s)[가-힣]{2,4}리(?![가-힣])"      # '동면 상걸리' 형태의 지명만 검사
    ri = Counter(w for r in rows for w in re.findall(PLACE, " ".join(r["opinion"])))
    warnings = []
    if ri:
        common = ri.most_common(1)[0][0]
        for r in rows:
            odd = [w for w in set(re.findall(PLACE, " ".join(r["opinion"]))) if w != common]
            if odd and common not in " ".join(r["opinion"]): warnings.append(f"{r['dept']}: 지명 '{'/'.join(odd)}' (다른 회신은 '{common}') — 원문 오기 여부 확인")
    json.dump({"rows": rows, "skipped": skipped, "warnings": warnings}, open(out_json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"표 대상 {len(rows)}개 부서, 제외/스킵 {len(skipped)}건")
    for r in rows: print(" +", r["dept"], f"({r['source']})", f"의견 {len(r['opinion'])}항목 / 조치 {len(r['action'])}항목")
    for s in skipped: print(" -", s[0], "::", s[1])
    for w in warnings: print(" !", w)

if __name__ == "__main__":
    main(*sys.argv[1:])

