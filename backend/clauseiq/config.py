"""Global constants."""

from pathlib import Path

MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024  # 50 MB
ALLOWED_EXTENSIONS = {
    ".pdf", ".docx", ".doc",
    ".txt", ".html", ".htm",
    ".png", ".jpg", ".jpeg", ".tiff", ".tif",
}
CHUNK_SIZE_TOKENS = 512
CHUNK_OVERLAP_TOKENS = 64

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
INDEX_HTML_PATH = REPO_ROOT / "index.html"
