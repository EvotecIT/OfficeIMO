#!/usr/bin/env python3
"""Run one OCR engine on one fixture and emit a common JSON result."""

from __future__ import annotations

import argparse
import csv
import io
import json
import os
from pathlib import Path
import subprocess
import sys


def tesseract_result(args: argparse.Namespace) -> dict:
    root = args.tesseract_root.resolve()
    executable = root / "usr/bin/tesseract"
    tessdata = root / "usr/share/tesseract-ocr/5/tessdata"
    environment = os.environ.copy()
    environment["LD_LIBRARY_PATH"] = str(root / "usr/lib/x86_64-linux-gnu")
    environment["TESSDATA_PREFIX"] = str(tessdata)
    completed = subprocess.run(
        [str(executable), str(args.image), "stdout", "-l", args.language, "--psm", "6", "tsv"],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
        timeout=120,
    )
    words = []
    for row in csv.DictReader(io.StringIO(completed.stdout), delimiter="\t"):
        text = (row.get("text") or "").strip()
        if not text or row.get("level") != "5":
            continue
        confidence = float(row["conf"])
        words.append({
            "text": text,
            "confidence": max(0.0, confidence / 100.0),
            "x": int(row["left"]),
            "y": int(row["top"]),
            "width": int(row["width"]),
            "height": int(row["height"]),
        })
    return {"engine": "Tesseract", "text": " ".join(word["text"] for word in words), "words": words}


def rapid_result(args: argparse.Namespace) -> dict:
    sys.path.insert(0, str(args.rapid_packages.resolve()))
    from rapidocr import RapidOCR

    recognition_language = "pl" if "pol" in args.language.split("+") else "en"
    engine = RapidOCR(params={
        "Global.model_root_dir": str(args.rapid_models.resolve()),
        "Rec.lang_type": recognition_language,
    })
    result = engine(str(args.image))
    boxes = result.boxes.tolist() if result.boxes is not None else []
    texts = result.txts or []
    scores = result.scores.tolist() if hasattr(result.scores, "tolist") else (result.scores or [])
    words = []
    for box, text, score in zip(boxes, texts, scores):
        xs = [point[0] for point in box]
        ys = [point[1] for point in box]
        words.append({
            "text": text,
            "confidence": float(score),
            "x": float(min(xs)),
            "y": float(min(ys)),
            "width": float(max(xs) - min(xs)),
            "height": float(max(ys) - min(ys)),
        })
    return {"engine": "RapidOCR", "text": " ".join(texts), "words": words}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--engine", choices=("tesseract", "rapidocr"), required=True)
    parser.add_argument("--image", type=Path, required=True)
    parser.add_argument("--language", required=True)
    parser.add_argument("--tesseract-root", type=Path, required=True)
    parser.add_argument("--rapid-packages", type=Path, required=True)
    parser.add_argument("--rapid-models", type=Path, required=True)
    args = parser.parse_args()
    if not args.image.is_file():
        raise FileNotFoundError(args.image)
    args.rapid_models.mkdir(parents=True, exist_ok=True)
    result = tesseract_result(args) if args.engine == "tesseract" else rapid_result(args)
    print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))


if __name__ == "__main__":
    main()
