from __future__ import annotations

import os

from waitress import serve

from app import app, ensure_storage


def main() -> None:
    ensure_storage()
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8000"))
    threads = int(os.environ.get("WAITRESS_THREADS", "4"))
    serve(app, host=host, port=port, threads=threads)


if __name__ == "__main__":
    main()
