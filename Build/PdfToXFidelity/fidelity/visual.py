"""Page-image similarity between the source PDF and the converted document re-rendered to PDF by LibreOffice.

Pages are rasterised by Poppler's pdftoppm (in WSL) at a low resolution in greyscale and compared with a
mean SSIM over 7x7 windows. SSIM rewards a blank page against a mostly white page and punishes reflowed text
heavily, so the primary score is `visual_layout`: the Dice overlap of ink density on a coarse 22x16 grid per page
(0 for a blank or missing page, 1 for the same layout). `blank_ssim` records what a white page would score.
"""
from __future__ import annotations

import glob
import hashlib
import os
import shutil
import subprocess
import threading

import numpy as np
from PIL import Image

DPI = 36
MAX_PAGES = 8


def _wsl(path: str) -> str:
    path = os.path.abspath(path)
    drive, rest = os.path.splitdrive(path)
    return "/mnt/" + drive[0].lower() + rest.replace("\\", "/")


_locks: dict[str, threading.Lock] = {}
_locks_guard = threading.Lock()


def render(pdf: str, png_dir: str) -> list[np.ndarray]:
    """Rasterises once per PDF; parallel jobs for the same source wait for the first render to finish."""
    key = hashlib.sha1(os.path.abspath(pdf).encode()).hexdigest()[:16]
    folder = os.path.join(png_dir, key)
    with _locks_guard:
        lock = _locks.setdefault(key, threading.Lock())
    with lock:
        if not os.path.isdir(folder):
            staging = folder + ".partial"
            shutil.rmtree(staging, ignore_errors=True)
            os.makedirs(staging)
            subprocess.run(["wsl", "-e", "pdftoppm", "-r", str(DPI), "-gray", "-png", "-l", str(MAX_PAGES), _wsl(pdf),
                            _wsl(os.path.join(staging, "p"))], capture_output=True, timeout=300, check=False)
            os.replace(staging, folder)
    pages = sorted(glob.glob(os.path.join(folder, "p-*.png")))
    return [np.asarray(Image.open(p).convert("L"), dtype=np.float64) for p in pages]


def _box(img: np.ndarray, k: int = 7) -> np.ndarray:
    c = np.cumsum(np.cumsum(np.pad(img, ((1, 0), (1, 0))), 0), 1)
    return (c[k:, k:] - c[:-k, k:] - c[k:, :-k] + c[:-k, :-k]) / (k * k)


def ssim(a: np.ndarray, b: np.ndarray) -> float:
    c1, c2 = (0.01 * 255) ** 2, (0.03 * 255) ** 2
    ma, mb = _box(a), _box(b)
    va = _box(a * a) - ma * ma
    vb = _box(b * b) - mb * mb
    cov = _box(a * b) - ma * mb
    s = ((2 * ma * mb + c1) * (2 * cov + c2)) / ((ma * ma + mb * mb + c1) * (va + vb + c2))
    return float(s.mean())


def ink_overlap(a: np.ndarray, b: np.ndarray, grid: tuple[int, int] = (22, 16)) -> float:
    """Dice overlap of ink density on a coarse grid: 0 for a blank page, 1 for the same layout."""
    def cells(img):
        ink = 1.0 - img / 255.0
        h, w = ink.shape
        gh, gw = grid
        trimmed = ink[: h - h % gh, : w - w % gw]
        return trimmed.reshape(gh, trimmed.shape[0] // gh, gw, trimmed.shape[1] // gw).mean(axis=(1, 3))
    ca, cb = cells(a), cells(b)
    total = ca.sum() + cb.sum()
    return float(2 * np.minimum(ca, cb).sum() / total) if total else 1.0


def _fit(img: np.ndarray, shape) -> np.ndarray:
    if img.shape == shape:
        return img
    return np.asarray(Image.fromarray(img.astype(np.uint8)).resize((shape[1], shape[0]), Image.BILINEAR), dtype=np.float64)


def compare_pdfs(source_pdf: str, output_pdf: str, png_dir: str) -> dict:
    src = render(source_pdf, png_dir)
    out = render(output_pdf, png_dir) if os.path.exists(output_pdf) else []
    if not src:
        return {"visual_layout": float("nan"), "visual_ssim": float("nan"), "blank_ssim": float("nan"), "src_pages": 0, "out_pages": len(out)}
    n = max(len(src), len(out))
    scores, blanks, layout = [], [], []
    for i in range(n):
        if i < len(src) and i < len(out):
            s = src[i]
            o = _fit(out[i], s.shape)
            blanks.append(ssim(s, np.full_like(s, 255.0)))
            scores.append(ssim(s, o))
            layout.append(ink_overlap(s, o))
        else:
            # A missing or extra page scores 0; its blank baseline is the page itself against white.
            page = src[i] if i < len(src) else out[i]
            blanks.append(ssim(page, np.full_like(page, 255.0)))
            scores.append(0.0)
            layout.append(0.0)
    return {"visual_layout": sum(layout) / n, "visual_ssim": sum(scores) / n, "blank_ssim": sum(blanks) / n,
            "src_pages": len(src), "out_pages": len(out)}
