"""Run the server: python -m clauseiq."""

import argparse
import os
import socket
import sys
from pathlib import Path

# Ensure the backend root is visible when running python __main__.py from the clauseiq folder.
HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import uvicorn


def main():
    parser = argparse.ArgumentParser(description="ClauseIQ API")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=int(os.getenv("CLAUSE_PORT", "8000")))
    parser.add_argument("--reload", action="store_true")
    parser.add_argument(
        "--log-level",
        default="info",
        choices=["critical", "error", "warning", "info", "debug", "trace"],
    )
    args = parser.parse_args()

    def port_free(host, port):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                s.bind((host, port))
                return True
            except OSError:
                return False

    port = next((p for p in range(args.port, args.port + 6) if port_free(args.host, p)), None)
    if port is None:
        raise RuntimeError("No free ports")
    if port != args.port:
        print(f"Port {args.port} in use → using {port}")

    uvicorn.run("clauseiq.api:app", host=args.host, port=port, log_level=args.log_level, reload=args.reload)


if __name__ == "__main__":
    main()
