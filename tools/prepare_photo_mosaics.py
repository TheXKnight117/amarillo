#!/usr/bin/env python3
"""Convert photos into compact color runs for Roblox SurfaceGui mosaics."""

from __future__ import annotations

import argparse
import json
import subprocess
from pathlib import Path


WIDTH = 64
HEIGHT = 48
QUANTUM = 16


def render_rgb(path: Path) -> bytes:
    command = [
        "ffmpeg",
        "-v",
        "error",
        "-i",
        str(path),
        "-vf",
        (
            f"scale={WIDTH}:{HEIGHT}:force_original_aspect_ratio=decrease,"
            f"pad={WIDTH}:{HEIGHT}:(ow-iw)/2:(oh-ih)/2:color=0xF6D6E5"
        ),
        "-frames:v",
        "1",
        "-f",
        "rawvideo",
        "-pix_fmt",
        "rgb24",
        "-",
    ]
    return subprocess.check_output(command)


def quantize(value: int) -> int:
    return min(255, round(value / QUANTUM) * QUANTUM)


def convert(path: Path) -> dict[str, object]:
    raw = render_rgb(path)
    expected = WIDTH * HEIGHT * 3
    if len(raw) != expected:
        raise RuntimeError(f"{path}: expected {expected} RGB bytes, received {len(raw)}")

    pixels = [
        tuple(quantize(channel) for channel in raw[index : index + 3])
        for index in range(0, len(raw), 3)
    ]
    runs: list[list[int]] = []
    for y in range(HEIGHT):
        x = 0
        while x < WIDTH:
            color = pixels[y * WIDTH + x]
            end = x + 1
            while end < WIDTH and pixels[y * WIDTH + end] == color:
                end += 1
            runs.append([x, y, end - x, *color])
            x = end

    return {
        "width": WIDTH,
        "height": HEIGHT,
        "runs": runs,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("images", nargs="+", type=Path)
    args = parser.parse_args()

    payload = {"photos": [convert(path) for path in args.images]}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(payload, separators=(",", ":")),
        encoding="utf-8",
    )

    run_count = sum(len(photo["runs"]) for photo in payload["photos"])
    print(f"Wrote {len(payload['photos'])} photos as {run_count} color runs")


if __name__ == "__main__":
    main()
