"""
ALIO 공기업 채용공고 수집기 - 핵심 로직
opendata.alio.go.kr 공개 API로 채용공고를 수집하고
NCS 및 자격증 정보를 추출해 Excel 파일로 저장합니다.

tkinter를 import하지 않으므로 디스플레이가 없는 환경에서도 동작합니다.
"""
import re
import json
import time
import zipfile
import tempfile
import threading
import queue
from datetime import datetime
from pathlib import Path
from collections import Counter, defaultdict
from typing import Optional

# ── 선택적 패키지 (설치 여부 확인) ──────────────────────────────
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    import pdfplumber
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

try:
    import olefile
    HAS_HWP = True
except ImportError:
    HAS_HWP = False

try:
    import pandas as pd
    import openpyxl  # noqa: F401
    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False

# ── 상수 ─────────────────────────────────────────────────────────
API_URL    = "https://opendata.alio.go.kr/new/odaApiMng/recrutInquiryAjaxList.do"
DETAIL_URL = "https://www.alio.go.kr/mobile/information/informationRecruitDtl.do"
FILE_URL   = "https://www.alio.go.kr/download/download.json"

CONFIG_FILE = Path(__file__).parent / "config.json"

ALIO_JOB_TYPES = [
    "일반직", "행정직", "경영직", "사무직", "관리직", "연구직",
    "기술직", "시설직", "전산직", "교수직", "의사직", "간호직",
    "약무직", "보건직", "위생직", "기능직", "운영직", "공무직",
    "안전직", "교원직", "별정직", "계약직", "전문직", "상담직", "기타",
]

ALIO_DEFAULT_ON = {"전산직"}

NCS_CATEGORIES = [
    "사업관리", "경영·회계·사무", "금융·보험", "교육·자연·사회과학",
    "법률·경찰·소방·교도·국방", "보건·의료", "사회복지·종교",
    "문화·예술·디자인·방송", "운전·운송", "영업판매", "경비·청소",
    "이용·숙박·여행·오락·스포츠", "음식서비스", "건설", "기계",
    "재료", "화학·바이오", "섬유·의복", "전기·전자", "정보통신",
    "식품가공", "인쇄·목재·가구·공예", "환경·에너지·안전", "농림어업",
]


# ── 설정 저장/로드 ────────────────────────────────────────────────
def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def save_config(cfg: dict):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


# ── 파일명 안전화 ─────────────────────────────────────────────────
def sanitize(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


# ── 제목 필터 파싱 ────────────────────────────────────────────────
# 문법: A | B  → A OR B   (둘 중 하나라도 포함)
#       A & B  → A AND B  (둘 다 포함)
#       A & B | C  → (A AND B) OR C
def _title_matches(title: str, expr: str) -> bool:
    """포함 조건 expr이 title에 매칭되면 True"""
    expr = expr.strip()
    if not expr:
        return True
    for or_group in expr.split("|"):
        and_terms = [t.strip() for t in or_group.split("&") if t.strip()]
        if and_terms and all(t in title for t in and_terms):
            return True
    return False

def _title_excluded(title: str, expr: str) -> bool:
    """제외 조건 expr에 매칭되면 True (= 이 공고는 제외)"""
    expr = expr.strip()
    if not expr:
        return False
    for or_group in expr.split("|"):
        and_terms = [t.strip() for t in or_group.split("&") if t.strip()]
        if and_terms and all(t in title for t in and_terms):
            return True
    return False


# ════════════════════════════════════════════════════════════════
# 파일 텍스트 추출
# ════════════════════════════════════════════════════════════════

def _hwp_text(path: Path) -> str:
    if not HAS_HWP:
        return ""
    try:
        ole = olefile.OleFileIO(str(path))
        if ole.exists("PrvText"):
            raw = ole.openstream("PrvText").read()
            text = raw.decode("utf-16-le", errors="ignore")
            ole.close()
            return text
        ole.close()
    except Exception:
        pass
    return ""

def _hwpx_text(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as z:
            if "Preview/PrvText.txt" in z.namelist():
                return z.open("Preview/PrvText.txt").read().decode("utf-8", errors="ignore")
    except Exception:
        pass
    return ""

def _pdf_text(path: Path) -> str:
    if not HAS_PDF:
        return ""
    try:
        with pdfplumber.open(str(path)) as doc:
            return "\n".join(p.extract_text() or "" for p in doc.pages[:30])
    except Exception:
        pass
    return ""

def _zip_text(path: Path) -> str:
    parts = []
    try:
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                ext = Path(name).suffix.lower()
                if ext not in (".hwp", ".hwpx", ".pdf"):
                    continue
                data = z.read(name)
                with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                    tmp.write(data)
                    tmp_path = Path(tmp.name)
                if ext == ".hwp":
                    parts.append(_hwp_text(tmp_path))
                elif ext == ".hwpx":
                    parts.append(_hwpx_text(tmp_path))
                elif ext == ".pdf":
                    parts.append(_pdf_text(tmp_path))
                tmp_path.unlink(missing_ok=True)
    except Exception:
        pass
    return "\n".join(parts)

def extract_file_text(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".hwp":  return _hwp_text(path)
    if ext == ".hwpx": return _hwpx_text(path)
    if ext == ".pdf":  return _pdf_text(path)
    if ext == ".zip":  return _zip_text(path)
    return ""


# ════════════════════════════════════════════════════════════════
# NCS 추출
# ════════════════════════════════════════════════════════════════

NCS_HEADER = "<NCS 분류체계><대분류><중분류><소분류><세분류>"

NCS_MAJOR_SET = {
    "사업관리", "경영·회계·사무", "금융·보험", "교육·자연·사회과학",
    "법률·경찰·소방·교도·국방", "보건·의료", "사회복지·종교",
    "문화·예술·디자인·방송", "운전·운송", "영업판매", "경비·청소",
    "이용·숙박·여행·오락·스포츠", "음식서비스", "건설", "기계",
    "재료", "화학·바이오", "섬유·의복", "전기·전자", "정보통신",
    "식품가공", "인쇄·목재·가구·공예", "환경·에너지·안전", "농림어업",
}
NCS_ALIAS = {
    "경영": "경영·회계·사무", "사무": "경영·회계·사무", "회계": "경영·회계·사무",
    "정보통신": "정보통신", "전산": "정보통신", "IT": "정보통신",
    "전기": "전기·전자", "전자": "전기·전자",
    "보건": "보건·의료", "의료": "보건·의료",
    "환경": "환경·에너지·안전", "에너지": "환경·에너지·안전", "안전": "환경·에너지·안전",
    "건설": "건설", "기계": "기계", "화학": "화학·바이오",
}

def _norm_major(raw: str) -> Optional[str]:
    raw = re.sub(r"^\d+\.\s*", "", raw).strip()
    if raw in NCS_MAJOR_SET:
        return raw
    for k, v in NCS_ALIAS.items():
        if k in raw:
            return v
    return None

def parse_ncs_from_api_str(ncs_str: str) -> list:
    if not ncs_str:
        return []
    return [item.strip().replace(".", "·") for item in ncs_str.split(",") if len(item.strip()) >= 2]

def _parse_ncs_items(text: str) -> tuple:
    """(대분류 목록, 세분류 목록) 반환"""
    idx = text.find(NCS_HEADER)
    if idx == -1:
        return [], []
    snippet = text[idx + len(NCS_HEADER): idx + len(NCS_HEADER) + 1200]
    skip_kw = {
        "대분류", "중분류", "소분류", "세분류", "NCS 분류체계",
        "참고", "참조", "공고문", "직무수행", "능력 단위", "주요사업",
        "전형방법", "교육요건", "기관 주요사업", "핵심직무",
    }
    items = [it for it in re.findall(r"<([^<>\n]{2,80})>", snippet) if it not in skip_kw]
    majors, leaves = [], []
    i = 0
    while i <= len(items) - 4:
        m = _norm_major(items[i])
        if m:
            if m not in majors:
                majors.append(m)
            leaf = re.sub(r"^\d+\.\s*", "", items[i + 3]).strip()
            bad = (
                "→" in leaf or "○" in leaf
                or re.search(r"\d+\.", leaf)
                or len(leaf) > 20
                or any(w in leaf for w in ("참고", "참조", "공고문"))
            )
            if not bad and len(leaf) >= 2 and leaf not in leaves:
                leaves.append(leaf)
            i += 4
        else:
            i += 1
    return majors, leaves


# ════════════════════════════════════════════════════════════════
# 자격증 추출
# ════════════════════════════════════════════════════════════════

EXPLICIT_CERTS = [
    ("정보처리기사",           "정보처리기사"),
    ("정보처리산업기사",       "정보처리산업기사"),
    ("정보보안기사",           "정보보안기사"),
    ("정보보안산업기사",       "정보보안산업기사"),
    ("네트워크관리사",         "네트워크관리사"),
    ("리눅스마스터",           "리눅스마스터"),
    ("SQLD",                  "SQLD"),
    ("SQLP",                  "SQLP"),
    ("ADsP",                  "ADsP"),
    ("ADP",                   "ADP"),
    ("빅데이터분석기사",       "빅데이터분석기사"),
    ("데이터분석준전문가",     "데이터분석준전문가(ADsP)"),
    ("컴퓨터활용능력",         "컴퓨터활용능력"),
    ("워드프로세서",           "워드프로세서"),
    ("MOS",                   "MOS"),
    ("ITQ",                   "ITQ"),
    ("전기기사",               "전기기사"),
    ("전기산업기사",           "전기산업기사"),
    ("전기기능사",             "전기기능사"),
    ("소방설비기사",           "소방설비기사"),
    ("산업안전기사",           "산업안전기사"),
    ("사무자동화산업기사",     "사무자동화산업기사"),
    ("정보시스템감리사",       "정보시스템감리사"),
    ("전자계산기조직응용기사", "전자계산기조직응용기사"),
    ("무선설비기사",           "무선설비기사"),
    ("방송통신기사",           "방송통신기사"),
    ("통신기기기능사",         "통신기기기능사"),
    ("간호사",                 "간호사면허"),
    ("사회복지사",             "사회복지사"),
    ("운전면허",               "운전면허"),
]

_CERT_SUFFIX_RE = re.compile(r"(?:기사|산업기사|기능사|면허)")
# 후보 이름을 뒤에서 앞으로 훑을 때 허용할 문자.
# 줄바꿈/탭/괄호는 제외한다 — 다른 문장·항목(예: "자격증(면허) 소지자" 같은 안내 문구)까지
# 이어붙이면서 앞부분이 잘리거나 엉뚱한 텍스트가 섞이는 문제의 원인이었다.
_CERT_CHARCLASS_RE = re.compile(r"[가-힣A-Z· ]")
_CERT_MAX_LEN = 18
_QUAL_SECTION_RE = re.compile(
    # 헤더 단어 뒤의 구분자(":", "-", 공백, 괄호, 태그 꺾쇠 등)만 건너뛴다.
    # 예전엔 개행 전까지 아무 문자나 최대 10자를 무조건 건너뛰었는데,
    # "<자격요건><응시 자격><간호사 면허증...>" 처럼 헤더 바로 뒤에 짧은
    # 하위 태그가 연달아 붙는 유사-HTML 형식 문서에서는 이 고정폭 스킵이
    # 실제 내용 한가운데서 끝나버려 자격증명 앞글자가 잘려나가는 원인이었다.
    # 구분자가 아닌 실제 글자(한글/영문/숫자)는 절대 건너뛰지 않아야 한다.
    r"(?:응시자격|자격기준|필요자격|자격요건|우대사항|필수자격|보유자격|취득자격)[\s:：\-–·.,\(\)\[\]<>]{0,10}(.{0,600})",
    re.DOTALL,
)
_CERT_NOISE = [
    r"^에 따른", r"^이상의", r"^관련", r"^지원한", r"^직무와", r"^직종에",
    r"^점\)", r"^\d", r"^[A-Z]\s",
    r"자격증명", r"면허\s*\)", r"\(면허", r"및\s", r"서류", r"비고",
    r"합계", r"총점", r"교원자격증", r"자격 및", r"해당 자격증",
    r"필요한 자격증", r"관련 자격증", r"종료시\)", r"정량평가",
    r"\(국가", r"른 해당", r"직무 관련", r"자격목록", r"최종합격",
    r"자격증 사본", r"^에 관련된", r"^위의", r"^이전",
    r"관련된 자격증", r"제시된 자격증", r"응시자격", r"자격기준",
    r"^자격[·\s]", r"^자격증",
    # 실제 자격증명에는 절대 나타나지 않는 일반 서술 어휘 — 안내문 조각을 걸러낸다.
    r"기준", r"다르거나", r"명시된", r"반드시", r"^기타", r"^최초",
    r"^있어", r"^전문\s", r"^분야\s",
    # "OO분야 기사/산업기사/기능사"는 특정 자격증명이 아니라 막연한 서술이다.
    r"분야\s*(?:기사|산업기사|기능사)",
]
# 진짜 자격증명 앞(혹은 두 자격증명 사이)에 끼어들어 오추출을 일으키는
# 조사/접속사·안내 문구. 후보 문자열 전체에서 이 경계 표현이 나타나는
# "마지막" 위치를 찾아 그 뒤쪽만 남긴다 — 예를 들어
# "임상병리사면허 또는 방사선사 면허" 처럼 두 자격증이 접속사로 이어진
# 채로 통째로 매치되는 경우, "또는" 뒤의 "방사선사 면허"만 남긴다.
_BOUNDARY_RE = re.compile(
    r"또는|및|혹은|그리고|위의|이전|관련된|제시된|해당하는|해당|필수|"
    r"응시자격|자격기준|자격요건|필요자격|우대사항|보유자격|취득자격|자격증명|"
    r"[은는이가을를와과의에로](?=\s)|"
    # 숫자는 애초에 후보 문자에서 제외되므로("1명", "2급" 등) 숫자 뒤에 오는
    # 단위명사만 홀로 남는다 — 이 경우 단위명사까지 마저 잘라낸다.
    r"[명급종개회차위](?=\s)"
)

# 원문마다 "·" 대신 비슷하게 생긴 다른 유니코드 가운뎃점(‧ ∙ • ・ ⋅)을 쓰는 경우가 있어
# 같은 자격증명이 서로 다른 문자열로 취급되어 통계가 쪼개지는 문제가 있었다. 하나로 통일한다.
_DOT_VARIANTS_RE = re.compile("[‧∙•・⋅]")

def _normalize_dots(s: str) -> str:
    s = _DOT_VARIANTS_RE.sub("·", s)
    return re.sub(r"\s*·\s*", "·", s)

def _strip_leading_noise(cert: str) -> str:
    cert = _normalize_dots(cert).strip(" ·")
    last_end = 0
    for bm in _BOUNDARY_RE.finditer(cert):
        last_end = bm.end()
    return cert[last_end:].strip(" ·")

def _is_valid_cert(cert: str) -> bool:
    cert = re.sub(r"\s+", " ", cert).strip()
    if len(cert) < 4:
        return False
    for pat in _CERT_NOISE:
        if re.search(pat, cert):
            return False
    return True

def extract_certs(text: str) -> list:
    found: set = set()
    text_lower = text.lower()
    for kw, canonical in EXPLICIT_CERTS:
        if kw.lower() in text_lower:
            found.add(canonical)

    section_found: set = set()
    for m in _QUAL_SECTION_RE.finditer(text):
        section = m.group(1)
        # 접미사(기사/산업기사/기능사/면허)가 나오는 지점마다 독립적으로 뒤에서
        # 앞으로 스캔한다. 예전에는 finditer가 이전 매치 끝에서부터 이어서
        # 스캔했기 때문에, 앞쪽에 잡음(조사·안내문구)이 함께 매치되면 바로 뒤에
        # 오는 진짜 자격증명의 앞글자가 잘려나가는 문제가 있었다.
        for sm in _CERT_SUFFIX_RE.finditer(section):
            end = sm.end()
            start = end
            while start > 0 and (end - start) < _CERT_MAX_LEN and _CERT_CHARCLASS_RE.match(section[start - 1]):
                start -= 1
            candidate = _strip_leading_noise(section[start:end])
            if _is_valid_cert(candidate):
                section_found.add(candidate)

    tag_found: set = set()
    for cert in re.findall(r"<([^<>]{4,30}(?:기사|산업기사|기능사|면허))>", text):
        cert = _strip_leading_noise(cert)
        if _is_valid_cert(cert):
            tag_found.add(cert)

    # 태그 규칙(<...>)은 항목 경계가 명확해 일반 스캔보다 더 완전한 이름을 잡아낸다.
    # 예: "<영선·조경기사>"는 온전히 잡히지만, 같은 텍스트를 일반 스캔이 훑으면
    # 가운뎃점에서 멈춰 "조경기사"만 잡는다. 이런 경우에만 짧은 쪽을 버린다.
    # (일반 스캔끼리는 서로 다른 자격증이 우연히 부분 문자열 관계일 수 있으므로
    # — 예: "건축기사"와 "실내건축기사" — 서로 건드리지 않는다.)
    section_found = {
        c for c in section_found
        if not any(c != t and c in t for t in tag_found)
    }
    regex_found = section_found | tag_found

    for c in regex_found:
        # 이미 명시적으로 잡힌 자격증명을 잡음과 함께 다시 잡은 경우(부분 문자열로
        # 포함)는 중복/오염된 항목이므로 건너뛴다.
        if any(kw != c and kw in c for kw, _ in EXPLICIT_CERTS):
            continue
        found.add(c)
    return sorted(found)


# ── NCS 직업기초능력 ─────────────────────────────────────────────
NCS_ABILITIES = [
    "의사소통능력", "수리능력", "문제해결능력", "자기개발능력",
    "자원관리능력", "대인관계능력", "정보능력", "기술능력",
    "조직이해능력", "직업윤리",
]

def extract_ncs_abilities(text: str) -> list:
    return [ab for ab in NCS_ABILITIES if ab in text]


# ════════════════════════════════════════════════════════════════
# API 수집 및 파일 다운로드
# ════════════════════════════════════════════════════════════════

def _fetch_page(session, keyword: str, page: int, page_size: int) -> tuple:
    r = session.get(
        API_URL,
        params={"pageNo": str(page), "pageUnit": str(page_size), "recrutPbancTtl": keyword},
        timeout=15,
    )
    data = r.json()["data"]
    return data["result"], int(data["totalCount"])

def _get_file_list(session, seq: str) -> list:
    r = session.get(DETAIL_URL, params={"seq": seq}, timeout=15)
    return re.findall(r'/download/download\.json\?fileNo=(\d+)[^>]*>([^<]+)</a>', r.text)

def _download_file(session, file_no: str, filename: str, seq: str, save_dir: Path) -> Optional[str]:
    save_path = save_dir / sanitize(filename or f"{seq}_{file_no}")
    if save_path.exists():
        return save_path.name
    r = session.get(FILE_URL, params={"fileNo": file_no}, timeout=30, stream=True)
    if r.status_code != 200:
        return None
    if not filename:
        cd = r.headers.get("Content-Disposition", "")
        m = re.search(r'filename=[""]*(.+?)[""]*\s*$', cd)
        if m:
            raw = m.group(1).strip('"')
            try:
                filename = raw.encode("latin-1").decode("utf-8")
            except Exception:
                filename = raw
        save_path = save_dir / sanitize(filename or f"{seq}_{file_no}")
    with open(save_path, "wb") as f:
        for chunk in r.iter_content(chunk_size=8192):
            f.write(chunk)
    return save_path.name


# ════════════════════════════════════════════════════════════════
# 공고 분석 및 Excel 생성
# ════════════════════════════════════════════════════════════════

def _analyze_record(rec: dict, download_dir: Path) -> dict:
    inst_nm  = rec.get("inst_nm", "")
    title    = rec.get("title", "")

    ncs_cats: list = parse_ncs_from_api_str(rec.get("ncs_str", ""))
    ncs_leaves: list = []
    certs: set = set()
    abilities: set = set()

    qual_text = rec.get("qual_text", "")
    if qual_text:
        for c in extract_certs(qual_text):
            certs.add(c)
        for ab in extract_ncs_abilities(qual_text):
            abilities.add(ab)

    for fname in rec.get("files", []):
        fpath = download_dir / fname
        if not fpath.exists() or fpath.suffix.lower() not in (".hwp", ".hwpx", ".pdf", ".zip"):
            continue
        text = extract_file_text(fpath)
        if not text:
            continue
        majors, leaves = _parse_ncs_items(text)
        if majors and not ncs_cats:
            ncs_cats = majors
        for lf in leaves:
            if lf not in ncs_leaves:
                ncs_leaves.append(lf)
        for c in extract_certs(text):
            certs.add(c)
        for ab in extract_ncs_abilities(text):
            abilities.add(ab)

    ncs_display = ncs_leaves if ncs_leaves else ncs_cats

    return {
        "기업명":     inst_nm,
        "공고제목":   title,
        "NCS 과목":   " / ".join(ncs_display) if ncs_display else "-",
        "요구 자격증": " / ".join(sorted(certs)) if certs else "-",
        "_certs":    sorted(certs),
        "_ncs":      ncs_display,
        "_abilities": sorted(
            abilities,
            key=lambda x: NCS_ABILITIES.index(x) if x in NCS_ABILITIES else 99,
        ),
    }


def build_excel(rows: list, out_path: Path):
    if not HAS_EXCEL:
        raise RuntimeError("pandas/openpyxl이 설치되지 않았습니다.\nsetup.bat을 실행하세요.")

    df_detail = pd.DataFrame([
        {
            "기업명":     r["기업명"],
            "공고제목":   r["공고제목"],
            "요구 자격증": r["요구 자격증"],
            "NCS 과목":   r["NCS 과목"],
        }
        for r in rows
    ])
    df_detail.index += 1

    # 자격증 통계
    cert_counter:   Counter = Counter()
    cert_companies: defaultdict = defaultdict(set)
    for r in rows:
        for c in r["_certs"]:
            cert_counter[c] += 1
            cert_companies[c].add(r["기업명"])

    df_cert = pd.DataFrame([
        {
            "자격증명": k,
            "공고 수":  v,
            "기업 수":  len(cert_companies[k]),
            "해당 기업": ", ".join(sorted(cert_companies[k])[:6])
                         + ("..." if len(cert_companies[k]) > 6 else ""),
        }
        for k, v in cert_counter.most_common()
    ])
    df_cert.index += 1

    # NCS 통계
    ncs_counter:   Counter = Counter()
    ncs_companies: defaultdict = defaultdict(set)
    for r in rows:
        for n in r["_ncs"]:
            ncs_counter[n] += 1
            ncs_companies[n].add(r["기업명"])

    df_ncs = pd.DataFrame([
        {
            "NCS 과목": k,
            "공고 수":  v,
            "기업 수":  len(ncs_companies[k]),
            "해당 기업": ", ".join(sorted(ncs_companies[k])[:6])
                         + ("..." if len(ncs_companies[k]) > 6 else ""),
        }
        for k, v in ncs_counter.most_common()
    ])
    df_ncs.index += 1

    # NCS 직업기초능력 통계
    ability_counter:   Counter = Counter()
    ability_companies: defaultdict = defaultdict(set)
    for r in rows:
        for ab in r["_abilities"]:
            ability_counter[ab] += 1
            ability_companies[ab].add(r["기업명"])

    df_ability = pd.DataFrame([
        {
            "NCS 직업기초능력": k,
            "공고 수":         v,
            "기업 수":         len(ability_companies[k]),
            "해당 기업":        ", ".join(sorted(ability_companies[k])[:6])
                               + ("..." if len(ability_companies[k]) > 6 else ""),
        }
        for k, v in sorted(
            ability_counter.items(),
            key=lambda x: NCS_ABILITIES.index(x[0]) if x[0] in NCS_ABILITIES else 99,
        )
    ])
    df_ability.index += 1

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_detail.to_excel(writer, sheet_name="공고별",       index=True, index_label="번호")
        df_cert.to_excel(  writer, sheet_name="자격증_통계",  index=True, index_label="순위")
        df_ncs.to_excel(   writer, sheet_name="NCS_통계",     index=True, index_label="순위")
        df_ability.to_excel(writer, sheet_name="NCS_직기능",  index=True, index_label="순위")

        for sheet in writer.sheets.values():
            for col in sheet.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                sheet.column_dimensions[col[0].column_letter].width = min(max_len + 4, 70)


# ════════════════════════════════════════════════════════════════
# 수집 워커 (백그라운드 스레드 / headless 양쪽에서 호출)
# ════════════════════════════════════════════════════════════════

def collect_worker(params: dict, msg_q: queue.Queue, stop_ev: threading.Event):
    """
    msg_q 메시지 형식:
      ("log",      text)
      ("progress", current, total, label)
      ("done",     excel_path, n_rows)
      ("error",    text)
    """
    def log(msg: str):
        msg_q.put(("log", msg))

    def prog_dl(cur: int, total: int, label: str = ""):
        msg_q.put(("prog_dl", cur, total, label))

    def prog_an(cur: int, total: int, label: str = ""):
        msg_q.put(("prog_an", cur, total, label))

    try:
        keywords      = params["keywords"]
        include_expr = params.get("include_expr", "").strip()
        exclude_expr = params.get("exclude_expr", "").strip()
        max_per_kw    = params.get("max_per_kw", 500)
        output_dir    = Path(params["output_dir"])
        do_download   = params.get("download_files", True)

        download_dir = output_dir / "Downloads"
        download_dir.mkdir(parents=True, exist_ok=True)

        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Referer": "https://www.alio.go.kr/",
        })

        seen_seqs: set = set()
        collected: list = []  # (keyword, item dict)

        log("=" * 50)
        log(f"수집 시작  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        log(f"키워드: {', '.join(keywords)}")
        if include_expr:
            log(f"포함 필터: {include_expr}")
        if exclude_expr:
            log(f"제외 필터: {exclude_expr}")
        log("=" * 50)

        # ── 1단계: 공고 목록 수집 ───────────────────────────────
        for keyword in keywords:
            if stop_ev.is_set():
                break

            log(f"\n['{keyword}'] 검색 중...")
            kw_new = 0
            kw_filtered = 0
            kw_dup = 0
            page_size = 100
            total_received = 0

            for page in range(1, 99999):
                if stop_ev.is_set():
                    break
                try:
                    items, total_count = _fetch_page(session, keyword, page, page_size)
                except Exception as e:
                    log(f"  API 오류 (page {page}): {e}")
                    break

                if not items:
                    break
                if page == 1:
                    log(f"  서버 총 {total_count:,}개 공고 (최대 {max_per_kw}개 수집)")

                total_received += len(items)
                page_filtered = 0
                page_dup = 0

                for item in items:
                    seq = str(item.get("recrutPblntSn", ""))
                    if not seq:
                        continue
                    if seq in seen_seqs:
                        page_dup += 1
                        continue
                    inst_nm = item.get("instNm", "")
                    title   = item.get("recrutPbancTtl", "")
                    filter_text = f"{inst_nm} {title}"  # 기관명 + 제목 합쳐서 필터 적용
                    if include_expr and not _title_matches(filter_text, include_expr):
                        page_filtered += 1
                        continue
                    if exclude_expr and _title_excluded(filter_text, exclude_expr):
                        page_filtered += 1
                        continue
                    seen_seqs.add(seq)
                    collected.append((keyword, item))
                    kw_new += 1
                    if kw_new >= max_per_kw:
                        break

                kw_filtered += page_filtered
                kw_dup += page_dup

                if kw_new >= max_per_kw:
                    break
                if total_received >= total_count:
                    break
                time.sleep(0.1)

            parts = [f"수집: {kw_new}개", f"필터 제외: {kw_filtered}개"]
            if kw_dup: parts.append(f"키워드 중복: {kw_dup}개")
            parts.append(f"전체 누적: {len(seen_seqs)}개")
            log(f"  '{keyword}' " + ",  ".join(parts))

        if stop_ev.is_set() and not collected:
            msg_q.put(("error", "수집이 중지되었습니다."))
            return

        total = len(collected)
        log(f"\n총 {total}개 공고 수집 완료.")

        # ── 2단계: 파일 다운로드 ────────────────────────────────
        records = []
        if do_download:
            log("첨부파일 다운로드 중...")
        else:
            log("첨부파일 다운로드 건너뜀 (API 데이터만 사용)")

        for idx, (keyword, item) in enumerate(collected):
            if stop_ev.is_set():
                log("중지됨.")
                break

            seq       = str(item.get("recrutPblntSn", ""))
            inst_nm   = item.get("instNm", "")
            title     = item.get("recrutPbancTtl", "")
            ncs_str   = item.get("ncsCdNmLst", "")
            qual_text = item.get("aplyQlfcCn", "")

            prog_dl(idx + 1, total, f"[{inst_nm}] {title[:38]}")

            dl_files = []
            if do_download:
                try:
                    file_list = _get_file_list(session, seq)
                    for file_no, fname in file_list:
                        result = _download_file(session, file_no, fname.strip(), seq, download_dir)
                        if result:
                            dl_files.append(result)
                    time.sleep(0.15)
                except Exception as e:
                    log(f"  파일 오류 [{inst_nm}]: {e}")

            records.append({
                "seq":       seq,
                "inst_nm":   inst_nm,
                "title":     title,
                "ncs_str":   ncs_str,
                "qual_text": qual_text,
                "keyword":   keyword,
                "files":     dl_files,
            })

        # ── 3단계: NCS·자격증 분석 ──────────────────────────────
        log(f"\n분석 중... ({len(records)}개 공고)")
        rows = []
        for idx, rec in enumerate(records):
            if stop_ev.is_set():
                break
            prog_an(idx + 1, len(records), f"{rec.get('inst_nm','')} · {rec.get('title','')[:35]}")
            rows.append(_analyze_record(rec, download_dir))

        # ── 4단계: Excel 저장 ───────────────────────────────────
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = output_dir / f"분석결과_{ts}.xlsx"
        log(f"\nExcel 저장 중: {excel_path.name}")
        build_excel(rows, excel_path)

        msg_q.put(("done", str(excel_path), len(rows)))

    except Exception as e:
        import traceback
        msg_q.put(("error", f"{e}\n\n{traceback.format_exc()}"))
