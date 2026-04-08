# Offline Document Text Upscaler (Real-ESRGAN)

This project enhances blurry document/text images (receipts, invoices, cropped text) using:
- preprocessing (denoise + luminance contrast + optional sharpening/deblur)
- Real-ESRGAN super-resolution (x2/x4)
- text-focused postprocessing (edge-aware sharpening + local contrast)

It runs fully local/offline after dependencies and model weights are present.

## Project Structure

- `main.py` -> CLI app for single image or batch folder processing
- `enhancer.py` -> enhancement pipeline and function API
- `upscale.py` -> hardcoded runner so you can run `python upscale.py`
- `requirements.txt` -> Python dependencies

## Installation (exact steps)

1. Create and activate a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate
```

2. Install dependencies:

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

If `torch` install fails on your machine, install PyTorch first from the official selector and then re-run `pip install -r requirements.txt`.

## Run Commands

### Quick run (hardcoded defaults)

```bash
python upscale.py
```

Default behavior in `upscale.py`:
- input: `Sample20.jpg`
- output folder: `outputs`
- scale: `4`
- denoise + sharpen: enabled

You can edit `DEFAULT_ARGS` inside `upscale.py` to hardcode your own paths.

### Single image via CLI

```bash
python main.py --input Sample20.jpg --output outputs --scale 4 --sharpen --denoise
```

### Folder/batch via CLI

```bash
python main.py --input images --output outputs --batch --scale 4 --sharpen --denoise
```

### Same CLI through `upscale.py`

```bash
python upscale.py --input images --output outputs --batch --scale 2 --no-sharpen
```

## CLI Flags

- `--input` image path or folder path
- `--output` output folder (auto-created if missing)
- `--scale` upscale factor (`2` or `4`)
- `--batch` process entire input folder
- `--sharpen` / `--no-sharpen`
- `--denoise` / `--no-denoise`

Supported formats: PNG, JPG, JPEG.

## Function API

```python
from enhancer import enhance_image

enhance_image("input.jpg", "outputs/input_enhanced_x4.jpg", scale=4)
```

## Model Weights

The script auto-downloads the required Real-ESRGAN weight file into `weights/` on first run:
- x4: `RealESRGAN_x4plus.pth`
- x2: `RealESRGAN_x2plus.pth`

If internet is blocked, download manually and place files in:

```text
./weights/
```

Then rerun the command offline.

## Why this is suitable for blurry text images

- Document-aware preprocessing improves low-contrast strokes before SR.
- Real-ESRGAN restores missing high-frequency details while upscaling.
- Edge-aware postprocessing sharpens character boundaries without strong halos.
- Pipeline is tuned for readability of text regions, not face/photo beautification.

