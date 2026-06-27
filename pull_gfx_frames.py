#!/usr/bin/env python3
"""
pull_gfx_frames.py — pull distinct graphic frames out of a video, ready for the GFX QC app.

Pipeline:
  1. ffmpeg samples frames from the video (default 1 per second).
  2. Perceptual-hash de-dup drops near-identical frames (a lower-third that sits
     on screen for 6s becomes ONE candidate, not 6).
  3. The OpenAI vision model gates each unique candidate: "is there a broadcast
     graphic here — yes/no". Only yes-frames are kept. (This is the accurate part
     that the old edge/blob script got wrong.)
  4. A second de-dup pass keeps one still per distinct graphic.
  5. Keepers are written to an output folder, named with their timecode, ready to
     upload to the QC web app.

Requirements:
  - ffmpeg installed (brew install ffmpeg)
  - pip install openai pillow
  - An OpenAI API key, found via (in order):
      $OPENAI_API_KEY,  or the QC app's saved config.json.

Usage:
  python3 pull_gfx_frames.py "MyEpisode.mp4"
  python3 pull_gfx_frames.py "MyEpisode.mp4" --fps 2 --outdir gfx_out
  python3 pull_gfx_frames.py "MyEpisode.mp4" --yes   (skip the cost confirmation)
"""

import os
import sys
import json
import base64
import shutil
import argparse
import subprocess
import tempfile
from pathlib import Path

from PIL import Image

try:
    from openai import OpenAI
except Exception:
    print("ERROR: the 'openai' package isn't installed. Run:  pip install openai")
    sys.exit(1)


# -----------------------------
# API key resolution (reuse the QC app's saved key if present)
# -----------------------------
def find_api_key() -> str:
    env = (os.environ.get("OPENAI_API_KEY") or "").strip()
    if env:
        return env
    cfg = Path.home() / "Library" / "Application Support" / "OffAxisGFXQC" / "config.json"
    try:
        if cfg.exists():
            data = json.loads(cfg.read_text())
            return (data.get("openai_api_key") or "").strip()
    except Exception:
        pass
    return ""


# -----------------------------
# ffmpeg helpers
# -----------------------------
def check_ffmpeg():
    if shutil.which("ffmpeg") is None:
        print("ERROR: ffmpeg not found. Install it with:  brew install ffmpeg")
        sys.exit(1)


def extract_frames(video: str, fps: float, workdir: str) -> list:
    """Sample frames at the given fps into workdir. Returns sorted list of paths."""
    out_pattern = os.path.join(workdir, "frame_%06d.jpg")
    cmd = [
        "ffmpeg", "-hide_banner", "-loglevel", "error",
        "-i", video,
        "-vf", f"fps={fps}",
        "-q:v", "3",
        out_pattern,
    ]
    print(f"Extracting frames at {fps} fps (this can take a minute on long videos)...")
    subprocess.run(cmd, check=True)
    frames = sorted(Path(workdir).glob("frame_*.jpg"))
    return [str(p) for p in frames]


# -----------------------------
# Perceptual hash (dHash) — no extra dependency
# -----------------------------
def dhash(path: str, hash_size: int = 8) -> int:
    img = Image.open(path).convert("L").resize((hash_size + 1, hash_size), Image.LANCZOS)
    pixels = list(img.getdata())
    bits = 0
    idx = 0
    for row in range(hash_size):
        for col in range(hash_size):
            left = pixels[row * (hash_size + 1) + col]
            right = pixels[row * (hash_size + 1) + col + 1]
            bits = (bits << 1) | (1 if left > right else 0)
            idx += 1
    return bits


def hamming(a: int, b: int) -> int:
    return bin(a ^ b).count("1")


def dedup_by_hash(frames: list, threshold: int = 6) -> list:
    """Drop frames whose hash is within `threshold` bits of the previously kept frame."""
    kept = []
    last_hash = None
    for f in frames:
        try:
            h = dhash(f)
        except Exception:
            continue
        if last_hash is None or hamming(h, last_hash) > threshold:
            kept.append((f, h))
            last_hash = h
    return kept


# -----------------------------
# AI gate: is there a broadcast graphic?
# -----------------------------
GATE_PROMPT = (
    "You are checking a single TV frame. Answer ONLY valid JSON: {\"graphic\": true|false}.\n"
    "Set graphic=true if the frame contains an intentional broadcast GRAPHIC such as a "
    "lower-third / name or title bar, a price or product-info graphic, an on-screen "
    "stat/score bug that is part of the design, or a full-screen graphic/title card.\n"
    "Set graphic=false for plain camera footage with no added graphic, slates, black, "
    "or burned-in timecode only. When unsure, answer false."
)


def encode_image(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def ai_is_graphic(client: OpenAI, path: str, model: str) -> bool:
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": GATE_PROMPT},
                {"role": "user", "content": [
                    {"type": "text", "text": "Is there a broadcast graphic? JSON only."},
                    {"type": "image_url",
                     "image_url": {"url": f"data:image/jpeg;base64,{encode_image(path)}"}},
                ]},
            ],
            response_format={"type": "json_object"},
        )
        data = json.loads(resp.choices[0].message.content)
        return bool(data.get("graphic"))
    except Exception as e:
        print(f"  (skip {os.path.basename(path)}: {e})")
        return False


# -----------------------------
# Timecode from frame index
# -----------------------------
def timecode_for(frame_path: str, fps: float) -> str:
    # frame_000123.jpg -> index 123 (1-based from ffmpeg)
    stem = Path(frame_path).stem
    try:
        idx = int(stem.split("_")[1])
    except Exception:
        idx = 0
    seconds = (idx - 1) / fps if fps else 0
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02d}{m:02d}{s:02d}"  # HHMMSS, filename-safe


# -----------------------------
# Main
# -----------------------------
def main():
    ap = argparse.ArgumentParser(description="Pull distinct graphic frames from a video.")
    ap.add_argument("video", help="Path to the video file (.mp4/.mov/etc.)")
    ap.add_argument("--fps", type=float, default=1.0, help="Frames sampled per second (default 1).")
    ap.add_argument("--outdir", default=None, help="Output folder (default: <video>_gfx).")
    ap.add_argument("--model", default="gpt-4.1-mini", help="OpenAI vision model.")
    ap.add_argument("--dedup", type=int, default=6, help="Hash distance for de-dup (higher = more aggressive).")
    ap.add_argument("--max-ai", type=int, default=0, help="Cap how many frames get sent to the AI (0 = no cap).")
    ap.add_argument("--yes", action="store_true", help="Skip the cost confirmation prompt.")
    args = ap.parse_args()

    if not os.path.exists(args.video):
        print(f"ERROR: video not found: {args.video}")
        sys.exit(1)

    check_ffmpeg()

    api_key = find_api_key()
    if not api_key:
        print("ERROR: no OpenAI API key. Set OPENAI_API_KEY or save one in the QC app first.")
        sys.exit(1)
    client = OpenAI(api_key=api_key)

    outdir = args.outdir or (str(Path(args.video).with_suffix("")) + "_gfx")
    os.makedirs(outdir, exist_ok=True)

    with tempfile.TemporaryDirectory() as workdir:
        frames = extract_frames(args.video, args.fps, workdir)
        if not frames:
            print("No frames extracted. Is the video valid?")
            sys.exit(1)
        print(f"Sampled {len(frames)} frames.")

        candidates = dedup_by_hash(frames, threshold=args.dedup)
        print(f"After de-dup: {len(candidates)} unique candidates.")

        if args.max_ai and len(candidates) > args.max_ai:
            print(f"Capping AI checks at {args.max_ai} (of {len(candidates)}).")
            candidates = candidates[:args.max_ai]

        if not args.yes:
            print(f"\nAbout to send {len(candidates)} frames to {args.model} for yes/no checks.")
            ans = input("Proceed? [y/N] ").strip().lower()
            if ans not in ("y", "yes"):
                print("Cancelled.")
                sys.exit(0)

        print("Checking frames with the vision model...")
        keepers = []
        last_kept_hash = None
        for i, (path, h) in enumerate(candidates, 1):
            if ai_is_graphic(client, path, args.model):
                # second de-dup: don't keep two near-identical graphics back to back
                if last_kept_hash is None or hamming(h, last_kept_hash) > args.dedup:
                    keepers.append((path, h))
                    last_kept_hash = h
            if i % 25 == 0:
                print(f"  ...{i}/{len(candidates)} checked, {len(keepers)} kept")

        print(f"\nKept {len(keepers)} distinct graphic frames.")
        for n, (path, h) in enumerate(keepers, 1):
            tc = timecode_for(path, args.fps)
            dst = os.path.join(outdir, f"gfx_{n:03d}_tc{tc}.jpg")
            shutil.copyfile(path, dst)

    print(f"Done. Frames saved to: {outdir}")
    print("Upload these to the QC web app.")


if __name__ == "__main__":
    main()
