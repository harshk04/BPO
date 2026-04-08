# Loan Document Processing Pipeline

An end-to-end intelligent document processing system that automatically extracts and processes loan information from document images using computer vision and AI.

## Overview

This project combines multiple technologies to:
1. **Detect loan document fields** using Roboflow object detection
2. **Extract data** from detected regions using Google Gemini AI vision API
3. **Calculate loan metrics** (interest, insurance, principal reduction)
4. **Generate structured Excel output** with processed loan records

## Features

- Batch processing of multiple loan documents
- Automatic field detection using Roboflow
- AI-powered data extraction with error handling
- Loan calculations (interest rates, insurance, principal reduction)
- Structured Excel output generation
- Comprehensive logging of processing steps

## Project Structure

- `final.py` - Main orchestration pipeline (Roboflow detection → Gemini extraction → Excel output)
- `maingem.py` - Google Gemini AI data extraction module
- `roboflow_inference.py` - Roboflow object detection integration
- `excel_loan_calculator.py` - Loan calculation and financial metrics
- `run_logger.py` - Logging utility
- `images/` - Input loan document images
- `outputs/` - Detection results and cropped regions from Roboflow
- `logs/` - Processing logs

## Installation

1. Clone or download the project
2. Create and activate a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

3. Install dependencies:

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

## Configuration

### Environment Variables

Create a `.env` file in the project root with:

```
ROBOFLOW_API_KEY=your_roboflow_api_key
ROBOFLOW_PROJECT=your_project_name
ROBOFLOW_VERSION=your_model_version
GOOGLE_API_KEY=your_google_api_key
```

### Roboflow Setup

- Configure Roboflow detection thresholds in `roboflow_inference.py`
- Adjust confidence and overlap parameters as needed

## Usage

### Run Main Pipeline

```bash
python final.py
```

This will:
- Process images from the `images/` folder
- Run Roboflow detection
- Extract data using Gemini AI
- Calculate loan metrics
- Generate `final_loan_outputs.xlsx`

### Process Individual Images

```bash
python maingem.py
```

Processes images from the configured input folder and extracts structured loan data.

## Output

The pipeline generates:
- `final_loan_outputs.xlsx` - Structured loan records with calculated metrics
- `extracted_loan_records.xlsx` - Raw extracted data (from Gemini)
- `outputs/` - Detection results and cropped loan document fields
- `logs/` - Detailed processing logs

## Data Extracted

- Customer reference number
- Customer name
- Address (city, state)
- Loan details (amount, period, interest rate)
- Guarantor information
- Purchase value and down payment
- Calculated metrics (total interest, insurance rates, etc.)

