"""
ALIO 공기업 채용공고 수집기 - headless CLI
GUI(main.py)와 동일한 collector_core.collect_worker를 사용하지만
tkinter/디스플레이 없이 커맨드라인에서 실행합니다. (GitHub Actions 등에서 사용)

사용 예:
    python headless_collect.py --keywords "전산직,정보통신" --max-per-kw 100 --output-dir ./output
"""
import argparse
import os
import sys
import threading
import queue
from pathlib import Path

from collector_core import collect_worker, HAS_REQUESTS, HAS_EXCEL

PROGRESS_PRINT_EVERY = 10  # 진행 로그를 몇 건마다 찍을지 (Actions 로그량 조절)


def parse_args():
    p = argparse.ArgumentParser(description="ALIO 공기업 채용공고 수집기 (headless)")
    p.add_argument("--keywords", required=True,
                    help="쉼표로 구분된 검색 키워드 목록 (예: 전산직,정보통신)")
    p.add_argument("--include-expr", default="", help="포함 필터 표현식 (A|B, A&B 문법)")
    p.add_argument("--exclude-expr", default="", help="제외 필터 표현식")
    p.add_argument("--max-per-kw", type=int, default=100, help="키워드당 최대 수집 개수")
    p.add_argument("--output-dir", default="./output", help="결과 저장 폴더")
    p.add_argument("--no-download", action="store_true",
                    help="첨부파일 다운로드 건너뛰기 (속도는 빠르지만 NCS/자격증 추출 정확도 하락)")
    return p.parse_args()


def main():
    args = parse_args()

    if not HAS_REQUESTS or not HAS_EXCEL:
        print("필수 패키지가 설치되지 않았습니다. requirements.txt를 설치하세요.", file=sys.stderr)
        sys.exit(1)

    keywords = [k.strip() for k in args.keywords.split(",") if k.strip()]
    if not keywords:
        print("키워드를 최소 1개 이상 입력하세요.", file=sys.stderr)
        sys.exit(1)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    params = {
        "keywords": keywords,
        "include_expr": args.include_expr,
        "exclude_expr": args.exclude_expr,
        "max_per_kw": args.max_per_kw,
        "output_dir": str(output_dir),
        "download_files": not args.no_download,
    }

    msg_q: queue.Queue = queue.Queue()
    stop_ev = threading.Event()

    worker = threading.Thread(target=collect_worker, args=(params, msg_q, stop_ev), daemon=True)
    worker.start()

    result_path = None
    error_msg = None

    while True:
        try:
            msg = msg_q.get(timeout=1.0)
        except queue.Empty:
            if not worker.is_alive():
                break
            continue

        kind = msg[0]
        if kind == "log":
            print(msg[1], flush=True)
        elif kind == "prog_dl":
            _, cur, total, label = msg
            if cur == total or cur % PROGRESS_PRINT_EVERY == 0:
                print(f"[다운로드 {cur}/{total}] {label}", flush=True)
        elif kind == "prog_an":
            _, cur, total, label = msg
            if cur == total or cur % PROGRESS_PRINT_EVERY == 0:
                print(f"[분석 {cur}/{total}] {label}", flush=True)
        elif kind == "done":
            _, excel_path, n_rows = msg
            result_path = excel_path
            print(f"\n완료: {n_rows}건 -> {excel_path}", flush=True)
            break
        elif kind == "error":
            error_msg = msg[1]
            print(f"\n오류: {error_msg}", file=sys.stderr, flush=True)
            break

    worker.join(timeout=5)

    if error_msg:
        sys.exit(1)
    if not result_path:
        print("결과 파일이 생성되지 않았습니다.", file=sys.stderr)
        sys.exit(1)

    # GitHub Actions에서 이후 스텝(release 업로드 등)이 결과 경로를 참조할 수 있도록 출력
    github_output = os.environ.get("GITHUB_OUTPUT")
    if github_output:
        with open(github_output, "a", encoding="utf-8") as f:
            f.write(f"excel_path={result_path}\n")
            f.write(f"excel_name={Path(result_path).name}\n")


if __name__ == "__main__":
    main()
