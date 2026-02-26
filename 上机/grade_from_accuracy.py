#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import csv
import math


def _is_number(x: str) -> bool:
    try:
        float(x)
        return True
    except Exception:
        return False


def read_input_csv(path: str):
    """Read CSV with 2 columns: (student_id, accuracy). Header optional."""
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = [r for r in reader if any(cell.strip() for cell in r)]

    if not rows:
        raise ValueError("输入CSV为空")

    # detect header: if 2nd cell of first row is not numeric
    start = 0
    if len(rows[0]) >= 2 and not _is_number(rows[0][1].strip()):
        start = 1

    data = []
    for line_no, r in enumerate(rows[start:], start=start + 1):
        if len(r) < 2:
            raise ValueError(f"第{line_no}行列数不足2: {r}")

        sid = r[0].strip()
        if not sid:
            raise ValueError(f"第{line_no}行学号为空")

        try:
            acc = float(r[1])
        except Exception:
            raise ValueError(f"第{line_no}行准确率无法解析为数字: {r[1]!r}")

        if math.isnan(acc) or math.isinf(acc):
            raise ValueError(f"第{line_no}行准确率为 NaN/Inf")

        data.append((sid, acc))

    if not data:
        raise ValueError("未读到任何有效数据行")

    return data


def minmax_scores(data, score_min=80.0, score_max=100.0):
    accs = [acc for _, acc in data]
    a_min = min(accs)
    a_max = max(accs)

    if a_max == a_min:
        # all same accuracy -> give full score
        return [(sid, score_max) for sid, _ in data], a_min, a_max

    scale = (score_max - score_min) / (a_max - a_min)
    scored = []
    for sid, acc in data:
        s = score_min + (acc - a_min) * scale
        # numerical safety clamp
        if s < score_min:
            s = score_min
        elif s > score_max:
            s = score_max
        scored.append((sid, s))
    return scored, a_min, a_max


def write_output_csv(path: str, scored_rows, int_score=False, ndigits=2):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["学号", "分数"])
        for sid, score in scored_rows:
            if int_score:
                out = int(round(score))
            else:
                out = round(score, ndigits)
            w.writerow([sid, out])


def main():
    ap = argparse.ArgumentParser(description="按准确率数值做min-max线性缩放到[80,100]，输出CSV。")
    ap.add_argument("input_csv", help="输入CSV：两列（学号，准确率），可带表头")
    ap.add_argument("output_csv", help="输出CSV：两列（学号，分数）")
    ap.add_argument("--min", dest="score_min", type=float, default=70.0, help="最低分(默认80)")
    ap.add_argument("--max", dest="score_max", type=float, default=100.0, help="最高分(默认100)")
    ap.add_argument("--int", dest="int_score", action="store_true", help="输出整数分（四舍五入）")
    ap.add_argument("--ndigits", type=int, default=2, help="保留小数位数（默认2）")
    args = ap.parse_args()

    if args.score_max < args.score_min:
        ap.error("--max 必须 >= --min")

    data = read_input_csv(args.input_csv)
    scored, a_min, a_max = minmax_scores(data, score_min=args.score_min, score_max=args.score_max)
    write_output_csv(args.output_csv, scored, int_score=args.int_score, ndigits=args.ndigits)

    print(f"OK. accuracy_min={a_min}, accuracy_max={a_max}, score_range=[{args.score_min},{args.score_max}] -> {args.output_csv}")


if __name__ == "__main__":
    main()