#!/usr/bin/env python3
"""Run Roboflow inference on a local image and save selected cropped outputs."""

from __future__ import annotations

import argparse
import math
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

try:
    from PIL import Image
except ImportError:
    Image = None

try:
    from roboflow import Roboflow
except ImportError:
    Roboflow = None

# Set this if you prefer hard-coding an image path instead of using --image.
DEFAULT_IMAGE_PATH = "Sample11.jpg"
DEFAULT_OUTPUT_DIR = "outputs"
DEFAULT_CONFIDENCE = 50
DEFAULT_OVERLAP = 50


@dataclass
class RoboflowSettings:
    api_key: str
    workspace: str
    project: str
    version: int
    confidence: int
    overlap: int


def ensure_dependencies() -> None:
    missing: List[str] = []
    if load_dotenv is None:
        missing.append("python-dotenv")
    if Image is None:
        missing.append("Pillow")
    if Roboflow is None:
        missing.append("roboflow")
    if missing:
        joined = ", ".join(missing)
        raise RuntimeError(
            f"Missing required package(s): {joined}. "
            "Install them with: pip install -r requirements.txt"
        )


def required_env(key: str) -> str:
    value = os.getenv(key, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {key}")
    return value


def validate_percent(name: str, value: int) -> None:
    if not 0 <= value <= 100:
        raise ValueError(f"{name} must be between 0 and 100. Got: {value}")


def sanitize_stem(stem: str) -> str:
    compact = re.sub(r"\s+", "", stem)
    compact = re.sub(r"[^A-Za-z0-9_-]", "", compact)
    return compact or "image"


def clear_previous_outputs(output_dir: Path, image_prefix: str) -> None:
    crop_pattern = re.compile(rf"^{re.escape(image_prefix)}-\d+\.[A-Za-z0-9]+$")
    annotated_pattern = re.compile(rf"^{re.escape(image_prefix)}-annotated\.[A-Za-z0-9]+$")
    predictions_name = f"{image_prefix}-predictions.json"

    for file_path in output_dir.glob(f"{image_prefix}-*"):
        if not file_path.is_file():
            continue
        name = file_path.name
        if crop_pattern.match(name) or annotated_pattern.match(name) or name == predictions_name:
            file_path.unlink()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Run Roboflow detection and save crops for only middle detections "
            "(skip first and last)."
        )
    )
    parser.add_argument(
        "--image",
        default=None,
        help=(
            "Path to input image. If omitted, DEFAULT_IMAGE_PATH constant inside "
            "the script is used."
        ),
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help=f"Directory for outputs (default: {DEFAULT_OUTPUT_DIR}).",
    )
    parser.add_argument(
        "--confidence",
        type=int,
        default=None,
        help=f"Confidence threshold 0-100 (default: {DEFAULT_CONFIDENCE}).",
    )
    parser.add_argument(
        "--overlap",
        type=int,
        default=None,
        help=f"Overlap threshold 0-100 (default: {DEFAULT_OVERLAP}).",
    )
    return parser.parse_args()


def load_settings(args: argparse.Namespace) -> RoboflowSettings:
    load_dotenv()

    version_raw = required_env("ROBOFLOW_VERSION")
    try:
        version = int(version_raw)
    except ValueError as exc:
        raise ValueError(f"ROBOFLOW_VERSION must be an integer. Got: {version_raw}") from exc

    confidence = args.confidence if args.confidence is not None else DEFAULT_CONFIDENCE
    overlap = args.overlap if args.overlap is not None else DEFAULT_OVERLAP
    validate_percent("Confidence threshold", confidence)
    validate_percent("Overlap threshold", overlap)

    return RoboflowSettings(
        api_key=required_env("ROBOFLOW_API_KEY"),
        workspace=required_env("ROBOFLOW_WORKSPACE"),
        project=required_env("ROBOFLOW_PROJECT"),
        version=version,
        confidence=confidence,
        overlap=overlap,
    )


def to_box_coords(prediction: Dict[str, Any], image_width: int, image_height: int) -> Dict[str, int]:
    x = float(prediction["x"])
    y = float(prediction["y"])
    w = float(prediction["width"])
    h = float(prediction["height"])

    left = max(0, math.floor(x - w / 2.0))
    top = max(0, math.floor(y - h / 2.0))
    right = min(image_width, math.ceil(x + w / 2.0))
    bottom = min(image_height, math.ceil(y + h / 2.0))

    if right <= left:
        right = min(image_width, left + 1)
    if bottom <= top:
        bottom = min(image_height, top + 1)

    return {"left": left, "top": top, "right": right, "bottom": bottom}


def draw_and_save_outputs(
    image_path: Path,
    predictions: List[Dict[str, Any]],
    output_dir: Path,
    image_prefix: str,
) -> Dict[str, Any]:
    base_image = Image.open(image_path).convert("RGB")
    image_width, image_height = base_image.size

    image_ext = image_path.suffix.lower() if image_path.suffix else ".jpg"
    if image_ext not in {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}:
        image_ext = ".jpg"

    detections: List[Dict[str, Any]] = []
    sorted_predictions = sorted(
        predictions,
        key=lambda p: (float(p.get("y", 0.0)), float(p.get("x", 0.0))),
    )

    for prediction in sorted_predictions:
        try:
            box = to_box_coords(prediction, image_width=image_width, image_height=image_height)
            center_x = float(prediction.get("x", 0.0))
            center_y = float(prediction.get("y", 0.0))
            box_width = float(prediction.get("width", 0.0))
            box_height = float(prediction.get("height", 0.0))
            confidence = float(prediction.get("confidence", 0.0))
        except (KeyError, TypeError, ValueError):
            continue

        index = len(detections)
        left = box["left"]
        top = box["top"]
        right = box["right"]
        bottom = box["bottom"]

        class_name = str(prediction.get("class", "unknown"))

        detections.append(
            {
                "index": index,
                "class": class_name,
                "confidence": confidence,
                "x": center_x,
                "y": center_y,
                "width": box_width,
                "height": box_height,
                "left": left,
                "top": top,
                "right": right,
                "bottom": bottom,
                "detection_id": prediction.get("detection_id"),
            }
        )

    saved_detections: List[Dict[str, Any]] = []
    for detection in detections[1:-1]:
        crop_path = output_dir / f"{image_prefix}-{detection['index']}{image_ext}"
        crop = base_image.crop(
            (
                detection["left"],
                detection["top"],
                detection["right"],
                detection["bottom"],
            )
        )
        crop.save(crop_path)
        detection["crop_path"] = str(crop_path.resolve())
        saved_detections.append(detection)

    return {"detections": saved_detections}


def run_inference(image_path: Path, output_dir: Path, settings: RoboflowSettings) -> Dict[str, Any]:
    rf = Roboflow(api_key=settings.api_key)
    workspace = rf.workspace(settings.workspace)
    project = workspace.project(settings.project)
    model = project.version(settings.version).model

    prediction_result = model.predict(
        str(image_path),
        confidence=settings.confidence,
        overlap=settings.overlap,
    ).json()

    if not isinstance(prediction_result, dict):
        raise ValueError("Unexpected Roboflow response format (expected JSON object).")

    predictions = prediction_result.get("predictions", [])
    if not isinstance(predictions, list):
        raise ValueError("Unexpected Roboflow response format ('predictions' is not a list).")

    image_prefix = sanitize_stem(image_path.stem)
    result_dir = output_dir / image_prefix
    result_dir.mkdir(parents=True, exist_ok=True)
    clear_previous_outputs(result_dir, image_prefix)

    drawing_result = draw_and_save_outputs(
        image_path=image_path,
        predictions=predictions,
        output_dir=result_dir,
        image_prefix=image_prefix,
    )

    return {
        "result_dir": result_dir,
        "detections": drawing_result["detections"],
    }


def main() -> int:
    args = parse_args()
    ensure_dependencies()
    settings = load_settings(args)

    image_path = Path(args.image or DEFAULT_IMAGE_PATH).expanduser()
    if not image_path.exists() or not image_path.is_file():
        raise FileNotFoundError(
            f"Input image not found: {image_path}. "
            "Set DEFAULT_IMAGE_PATH in script or pass --image."
        )

    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    result = run_inference(image_path=image_path.resolve(), output_dir=output_dir, settings=settings)

    print(f"Input image: {image_path.resolve()}")
    print(f"Output folder: {result['result_dir']}")
    print(f"Saved crops: {len(result['detections'])} (skipping first and last detection)")

    if result["detections"]:
        print("\nBounding boxes and crops:")
        for detection in result["detections"]:
            print(
                f"[{detection['index']}] class={detection['class']} "
                f"confidence={detection['confidence']:.3f} "
                f"center=({detection['x']:.1f}, {detection['y']:.1f}) "
                f"size=({detection['width']:.1f}, {detection['height']:.1f}) "
                f"box=({detection['left']}, {detection['top']}, "
                f"{detection['right']}, {detection['bottom']}) "
                f"crop={detection['crop_path']}"
            )

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
