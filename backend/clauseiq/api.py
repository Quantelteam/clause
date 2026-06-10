"""FastAPI app, job store, and endpoints."""

from __future__ import annotations

import logging
import time
import uuid
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Request, Response, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

from .config import MAX_FILE_SIZE_BYTES, ALLOWED_EXTENSIONS, INDEX_HTML_PATH
from .models import (
    JobStatus, ProcessedDocument, JobOut, ChunkOut,
    UploadResponse, JobStatusResponse, HealthOut,
)
from .pipeline import run_pipeline, _detect_doc_type, DocumentType, serialize_result

logger = logging.getLogger("clauseiq.api")

_job_store: dict[str, dict] = {}


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("ClauseIQ starting up")
    yield
    logger.info("ClauseIQ shutting down | jobs=%d", len(_job_store))


app = FastAPI(
    title="ClauseIQ",
    description="Legal contract intelligence pipeline API",
    version="0.1.0",
    lifespan=lifespan,
    docs_url="/docs",
    redoc_url="/redoc",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:3000",
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "http://localhost:8080",
        "http://localhost:8000",
        "http://127.0.0.1:8000",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.exception_handler(Exception)
async def _global_exc(request: Request, exc: Exception):
    logger.exception("Unhandled exception %s %s", request.method, request.url.path)
    return JSONResponse(status_code=500, content={"detail": "Internal error."})


async def _process_uploaded_file(file: UploadFile):
    if not file.filename:
        raise HTTPException(400, "No filename.")
    suffix = Path(file.filename).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise HTTPException(400, f"'{suffix}' not allowed. Supported: {sorted(ALLOWED_EXTENSIONS)}")
    content = await file.read()
    if not content:
        raise HTTPException(400, "Empty file.")
    if len(content) > MAX_FILE_SIZE_BYTES:
        raise HTTPException(400, f"File exceeds {MAX_FILE_SIZE_BYTES // 1048576} MB.")
    if _detect_doc_type(file.filename, content) == DocumentType.UNKNOWN:
        raise HTTPException(400, "Unrecognized document type.")
    job_id = str(uuid.uuid4())
    _job_store[job_id] = {"status": JobStatus.PROCESSING, "filename": file.filename}
    try:
        result = run_pipeline(content, file.filename, job_id)
    except (ValueError, RuntimeError, ImportError) as exc:
        _job_store[job_id] = {"status": JobStatus.FAILED, "message": str(exc)}
        raise HTTPException(400, str(exc)) from exc
    except Exception:
        _job_store[job_id] = {"status": JobStatus.FAILED, "message": "Pipeline error"}
        raise HTTPException(500, "Processing failed.")
    _job_store[job_id] = {"status": JobStatus.COMPLETED, "result": result}
    return job_id, result


@app.get("/", tags=["system"])
async def root():
    if not INDEX_HTML_PATH.exists():
        raise HTTPException(404, "Frontend not found.")
    return FileResponse(INDEX_HTML_PATH, media_type="text/html")


@app.get("/health", response_model=HealthOut, tags=["system"])
@app.get("/api/v1/health", response_model=HealthOut, tags=["system"])
async def health():
    return HealthOut(status="ok", version=app.version, timestamp=time.time())


@app.post("/process", response_model=JobOut, status_code=202, tags=["pipeline"])
async def process_document(file: UploadFile = File(...)):
    job_id, result = await _process_uploaded_file(file)
    return JobOut(
        job_id=result.job_id,
        filename=result.filename,
        status=JobStatus.COMPLETED,
        page_count=result.page_count,
        chunk_count=len(result.chunks),
        processing_time_ms=result.processing_time_ms,
        markdown=result.markdown,
        sha256=result.sha256,
        chunks=[
            ChunkOut(
                chunk_id=c.chunk_id,
                text=c.text,
                token_estimate=c.token_estimate,
                metadata=c.metadata,
            )
            for c in result.chunks
        ],
    )


@app.post("/api/v1/documents/upload", response_model=UploadResponse, status_code=202, tags=["pipeline"])
async def upload_document(file: UploadFile = File(...)):
    job_id, _ = await _process_uploaded_file(file)
    return UploadResponse(job_id=job_id, status=JobStatus.COMPLETED)


@app.get("/api/v1/jobs/{job_id}", response_model=JobStatusResponse, tags=["pipeline"])
async def get_job_status(job_id: str):
    entry = _job_store.get(job_id)
    if not entry:
        raise HTTPException(404, f"Job {job_id} not found")
    st = entry["status"]
    if st == JobStatus.COMPLETED and "result" in entry:
        return JobStatusResponse(job_id=job_id, status=st, results=serialize_result(entry["result"]))
    if st == JobStatus.FAILED:
        return JobStatusResponse(job_id=job_id, status=st, error=entry.get("message"))
    return JobStatusResponse(job_id=job_id, status=st)


@app.get("/jobs/{job_id}", response_model=JobOut, tags=["pipeline"])
async def get_job(job_id: str):
    entry = _job_store.get(job_id)
    if not entry:
        raise HTTPException(404, f"Job {job_id} not found")
    result = entry.get("result")
    if result:
        return JobOut(
            job_id=result.job_id,
            filename=result.filename,
            status=entry["status"],
            page_count=result.page_count,
            chunk_count=len(result.chunks),
            processing_time_ms=result.processing_time_ms,
            markdown=result.markdown,
            sha256=result.sha256,
            chunks=[
                ChunkOut(
                    chunk_id=c.chunk_id,
                    text=c.text,
                    token_estimate=c.token_estimate,
                    metadata=c.metadata,
                )
                for c in result.chunks
            ],
        )
    return JobOut(
        job_id=job_id,
        filename=entry.get("filename", ""),
        status=entry["status"],
        message=entry.get("message", ""),
    )


@app.delete("/jobs/{job_id}", status_code=204, tags=["pipeline"])
async def delete_job(job_id: str) -> Response:
    if job_id not in _job_store:
        raise HTTPException(404)
    del _job_store[job_id]
    return Response(status_code=204)
