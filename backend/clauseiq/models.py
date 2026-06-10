"""Enums, dataclasses, Pydantic schemas."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class JobStatus(str, Enum):
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"


class DocumentType(str, Enum):
    PDF = "pdf"
    DOCX = "docx"
    TXT = "txt"
    HTML = "html"
    IMAGE = "image"
    UNKNOWN = "unknown"


@dataclass
class TextBlock:
    text: str
    page: int
    x0: float = 0.0
    y0: float = 0.0
    x1: float = 0.0
    y1: float = 0.0
    font_size: float = 12.0
    is_bold: bool = False
    is_italic: bool = False
    block_type: str = "body"
    heading_level: int = 0


@dataclass
class DocumentChunk:
    chunk_id: str
    text: str
    token_estimate: int
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class ProcessedDocument:
    job_id: str
    filename: str
    doc_type: DocumentType
    page_count: int
    chunks: list[DocumentChunk]
    markdown: str
    processing_time_ms: int
    sha256: str


class ChunkOut(BaseModel):
    chunk_id: str
    text: str
    token_estimate: int
    metadata: dict[str, Any]


class JobOut(BaseModel):
    job_id: str
    filename: str
    status: JobStatus
    message: str = ""
    page_count: int = 0
    chunk_count: int = 0
    processing_time_ms: int = 0
    chunks: list[ChunkOut] = Field(default_factory=list)
    markdown: str = ""
    sha256: str = ""


class UploadResponse(BaseModel):
    job_id: str
    status: JobStatus
    message: str = ""


class JobStatusResponse(BaseModel):
    job_id: str
    status: JobStatus
    results: dict[str, Any] | None = None
    error: str | None = None


class HealthOut(BaseModel):
    status: str
    version: str
    timestamp: float
