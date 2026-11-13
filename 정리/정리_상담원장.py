# 상담원장 정리 스크립트
#
# 따르는 결정 기록: my-wiki/decisions/d-008-상담원장-정리의-다섯-결정.md
#   d-008-1  «종결» → «완료», «대기» → «보류» (CS기획팀 답)
#   d-008-2  처리시간 음수는 값을 그대로 두고 _표시 칸에 이유를 적는다
#   d-008-3  만족도 6~10점은 값을 그대로 두고 _표시 칸에 이유를 적는다. 평균 만족도에서는 뺀다
#   d-008-4  칸 이름은 원본 그대로 둔다 (_표시 칸만 맨 끝에 더한다)
#   d-008-5  고객번호 앞뒤 공백을 없앤다
#
# 그 밖의 처리는 원본의 «작성기준» 시트와, 처음 정리할 때의 처리 그대로다.
# 원본은 읽기만 한다. 결과는 정리본/ 의 세 파일에 쓴다.
#
# 실행:  python 정리/정리_상담원장.py      (cleaner 폴더에서)

import csv
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent.parent          # cleaner/
SRC = HERE.parent / "messy" / "상담원장_2025.xlsx"
OUT = HERE / "정리본"
OUT_CLEAN = OUT / "상담원장_정리.csv"
OUT_DROP = OUT / "지운줄.csv"
OUT_SUMMARY = OUT / "변경요약.md"

FLAG = "_표시"
PHONE = {"전화", " 전화", "유선", "전화상담", "TEL"}
CHANNEL = {"전화": "전화", " 전화": "전화", "유선": "전화", "전화상담": "전화", "TEL": "전화",
           "챗봇": "챗봇", "chatbot": "챗봇", "Chatbot": "챗봇", "챗 봇": "챗봇",
           "앱": "앱", "모바일앱": "앱", "app": "앱", "APP": "앱",
           "이메일": "이메일", "e-mail": "이메일", "E-mail": "이메일", "메일": "이메일"}
REGION = {"경기도": "경기", "서울시": "서울", "서울특별시": "서울"}
STATUS = {"완료 ": "완료", "처리완료": "완료",
          "종결": "완료",          # d-008-1
          "대기": "보류"}          # d-008-1
REASK = {"Y": "Y", "y": "Y", "예": "Y", "TRUE": "Y", "N": "N", "n": "N", "아니오": "N", "FALSE": "N"}
WHY_NEG = "처리시간 음수 — 틀린 값이지만 진짜 값을 몰라 그대로 둠 (d-008-2)"
WHY_SAT = "만족도 1~5 밖 — NPS를 잘못 넣은 것으로 보이나 미확인, 평균에서 뺌 (d-008-3)"


def iso(s):
    m = re.match(r"^(\d{4})[-/.](\d{2})[-/.](\d{2})$", str(s).strip())
    return f"{m[1]}-{m[2]}-{m[3]}"


def main():
    rows = list(load_workbook(SRC, read_only=True, data_only=True)["원장"].iter_rows(values_only=True))
    head = list(rows[0])
    raw = [dict(zip(head, r)) for r in rows[1:]]
    n_src = len(raw)
    change = Counter()          # 칸 값을 바꾼 줄 수
    drop = {}                   # 원본 순번 → 지운 이유

    # 1 테스트 접수 (작성기준: AG999 는 시스템 점검용)
    for i, d in enumerate(raw):
        if d["상담원ID"] == "AG999":
            drop[i] = "테스트 접수 (상담원ID AG999)"
    alive = [i for i in range(n_src) if i not in drop]

    # 2 표기 맞추기
    rec = {}
    for i in alive:
        d = dict(raw[i])
        for col, new in (
            ("접수일", iso(d["접수일"])),
            ("고객번호", d["고객번호"].strip()),                      # d-008-5
            ("접수채널", CHANNEL[d["접수채널"]]),
            ("상담원팀", d["상담원팀"].replace(" ", "")),
            ("지역", REGION.get(d["지역"], d["지역"])),
            ("처리상태", STATUS.get(d["처리상태"], d["처리상태"])),    # d-008-1 포함
        ):
            if new != d[col]:
                key = col
                if col == "처리상태":
                    key = f"처리상태 «{d[col]}» → «{new}»"
                change[key] += 1
                d[col] = new
        if isinstance(d["월정액"], str):
            d["월정액"] = int(d["월정액"].replace(",", "").replace("원", ""))
            change["월정액"] += 1
        if d["재문의"] is not None and REASK[d["재문의"]] != d["재문의"]:
            d["재문의"] = REASK[d["재문의"]]
            change["재문의"] += 1
        rec[i] = d

    # 3 완전 중복 (작성기준: 12월 초 접수분 이중 추출) — 표기를 맞춘 뒤 모든 칸이 같으면 먼저 나온 줄을 남긴다
    seen = {}
    for i in alive:
        k = tuple(rec[i][c] for c in head)
        if k in seen:
            drop[i] = f"완전 중복 — 원본 {seen[k] + 2}행과 모든 칸이 같음 (12월 이중 추출)"
        else:
            seen[k] = i
    alive = [i for i in alive if i not in drop]

    # 4 옛 상태 줄 (작성기준: 최종수정일이 늦은 줄이 최신)
    group = defaultdict(list)
    for i in alive:
        group[rec[i]["접수번호"]].append(i)
    for ids in group.values():
        ids.sort(key=lambda i: (rec[i]["최종수정일"], i))
        last = ids[-1]
        for i in ids[:-1]:
            drop[i] = f"옛 상태 줄 — 같은 접수번호의 최신 줄은 원본 {last + 2}행 (최종수정일 {rec[last]['최종수정일']})"
    keep = [i for i in alive if i not in drop]

    # 5 남긴 줄의 값 처리
    out = []
    for i in keep:
        d = rec[i]
        flags = []
        # 상담시간: 값이 180을 넘으면 초 단위로 보고 분으로 바꾼다
        if d["상담시간"] > 180:
            d["상담시간"] = round(d["상담시간"] / 60, 2)
            change["상담시간 초→분 (180 초과)"] += 1
        # 만족도
        v = d["만족도"]
        if isinstance(v, str) and v.endswith("점"):
            d["만족도"] = int(v[:-1])
            change["만족도 «N점» → 숫자"] += 1
        elif v in (0, "-", "미응답"):                    # 작성기준: 옛 시스템 미응답
            d["만족도"] = None
            change["만족도 미응답(0 · - · 미응답) → 빈칸"] += 1
        elif isinstance(v, int) and v > 5:               # d-008-3
            flags.append(WHY_SAT)
        # 처리시간 음수 — d-008-2
        t = d["처리시간"]
        if isinstance(t, (int, float)) and t < 0:
            flags.append(WHY_NEG)
        d[FLAG] = " · ".join(flags)
        out.append(d)

    # 6 쓰기
    OUT.mkdir(exist_ok=True)
    cols = head + [FLAG]                                  # d-008-4: 원본 칸 이름 그대로 + 맨 끝 _표시
    with open(OUT_CLEAN, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(cols)
        for d in out:
            w.writerow(["" if d[c] is None else d[c] for c in cols])
    with open(OUT_DROP, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["원본행"] + head + ["지운 이유"])
        for i in sorted(drop):
            w.writerow([i + 2] + ["" if raw[i][c] is None else raw[i][c] for c in head] + [drop[i]])

    reasons = Counter(r.split(" — ")[0].split(" (")[0] for r in drop.values())
    flagc = Counter(x for d in out for x in d[FLAG].split(" · ") if x)
    lines = [
        "# 상담원장 정리 — 변경 요약",
        "",
        "- 원본: `messy/상담원장_2025.xlsx` «원장» 시트 (읽기만 함)",
        "- 스크립트: `cleaner/정리/정리_상담원장.py`",
        "- 따른 결정: `my-wiki/decisions/d-008-상담원장-정리의-다섯-결정.md` · 원본 «작성기준» 시트",
        "",
        "## 줄 수",
        "",
        "| 단계 | 줄 |",
        "|---|---|",
        f"| 원본 (머리글 제외) | {n_src:,} |",
    ]
    for k in ("테스트 접수", "완전 중복", "옛 상태 줄"):
        lines.append(f"| − {k} | −{reasons[k]:,} |")
    lines += [f"| = 정리본 | {len(out):,} |", "",
              f"빼기 확인: {n_src:,} − {len(drop):,} = {n_src - len(drop):,} → 정리본 {len(out):,} "
              + ("(같음)" if n_src - len(drop) == len(out) else "(다름)"), "",
              "지운 줄은 모두 `정리본/지운줄.csv` 에 원본 값 그대로, 맨 끝 칸에 이유와 함께 있다.", "",
              "## 값을 바꾼 줄", "",
              "접수일 ~ 처리상태 줄은 중복을 지우기 전에 센 수라 지운 줄도 들어 있다. "
              "상담시간 · 만족도 줄은 정리본에 남은 줄만 센 수다.", "",
              "| 칸 | 무엇을 | 줄 |", "|---|---|---|"]
    desc = {
        "접수일": "`YYYY/MM/DD` · `YYYY.MM.DD` → `YYYY-MM-DD`",
        "고객번호": "앞뒤 공백 없앰 (d-008-5)",
        "접수채널": "17가지 표기 → 전화 · 챗봇 · 앱 · 이메일",
        "상담원팀": "띄어쓰기 없앰 («요금 팀» → «요금팀»)",
        "지역": "경기도 → 경기 · 서울시 / 서울특별시 → 서울",
        "월정액": "«33,000원» 같은 글자 → 숫자",
        "재문의": "y · 예 · TRUE → Y, n · 아니오 · FALSE → N",
    }
    for k, v in change.items():
        if k in desc:
            lines.append(f"| {k} | {desc[k]} | {v:,} |")
    for k, v in sorted(change.items()):
        if k.startswith("처리상태"):
            lines.append(f"| 처리상태 | {k[5:]}{' (d-008-1)' if ('종결' in k or '대기' in k) else ''} | {v:,} |")
    lines += [
        f"| 상담시간 | **값이 180을 넘는 줄은 초 단위로 보고 60으로 나눠 분으로 바꿈** | {change['상담시간 초→분 (180 초과)']:,} |",
        f"| 만족도 | «5점» 같은 글자 → 숫자 | {change['만족도 «N점» → 숫자']:,} |",
        f"| 만족도 | 옛 시스템 미응답(0 · - · 미응답) → 빈칸 (작성기준) | {change['만족도 미응답(0 · - · 미응답) → 빈칸']:,} |",
        "",
        "## 값은 그대로 두고 `_표시` 칸에 적은 줄",
        "",
        "| 이유 | 줄 |", "|---|---|",
    ]
    for k, v in flagc.items():
        lines.append(f"| {k} | {v:,} |")
    lines += ["", "## 정리본의 처리상태", "", "| 값 | 줄 |", "|---|---|"]
    for k, v in Counter(d["처리상태"] for d in out).most_common():
        lines.append(f"| {k} | {v:,} |")
    lines += ["", "## 칸", "", f"원본 칸 {len(head)}개의 이름과 순서는 그대로다 (d-008-4). 맨 끝에 `{FLAG}` 칸 하나를 더했다."]
    OUT_SUMMARY.write_text("\n".join(lines) + "\n", encoding="utf-8")

    # 다섯 숫자와 고객 수
    n = len(out)
    t = sum(d["상담시간"] for d in out) / n
    rq = [d["재문의"] for d in out if d["재문의"] is not None]
    sat = [d["만족도"] for d in out if d["만족도"] is not None and WHY_SAT not in d[FLAG]]
    done = sum(d["처리상태"] == "완료" for d in out)
    sys.stdout.reconfigure(encoding="utf-8")
    print(f"상담 건수      {n:,}")
    print(f"평균 상담시간  {t:.2f}분")
    print(f"재문의율      {100 * rq.count('Y') / len(rq):.2f}%  (Y {rq.count('Y'):,} ÷ {len(rq):,})")
    print(f"평균 만족도    {sum(sat) / len(sat):.3f}점  (응답 {len(sat):,}, 표시한 줄 뺌)")
    print(f"완료율        {100 * done / n:.2f}%  (완료 {done:,} ÷ {n:,})")
    print(f"고객 수        {len({d['고객번호'] for d in out}):,}")
    print(f"만족도 값이 있는 줄 중 처리시간 음수 표시 줄: {sum(1 for d in out if WHY_NEG in d[FLAG] and d['만족도'] is not None)}")


if __name__ == "__main__":
    main()
