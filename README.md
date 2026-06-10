# ClauseIQ

A local document intelligence API with a built-in frontend.

## Setup

1. Create and activate the backend virtual environment:

```powershell
cd c:\Users\Qq\Desktop\Clause\backend
python -m venv venv
.\venv\Scripts\Activate.ps1
```

2. Install dependencies:

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Run the app

```powershell
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Then open:

- `http://127.0.0.1:8000`
- `http://127.0.0.1:8000/docs`

## Tests

Run the backend tests:

```powershell
cd c:\Users\Qq\Desktop\Clause\backend
python -m unittest test_main.py
```

## Notes

- PDF uploads require `PyMuPDF`.
- DOCX uploads require `python-docx`.
- Image uploads require `Pillow` and `pytesseract`.
- If a dependency is missing, the API now returns a helpful installation message.
