from __future__ import annotations

import os
import socket
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from tkinter import Tk, messagebox

from waitress import serve

from app import app, ensure_storage


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def wait_for_server(url: str, timeout_seconds: float = 30.0) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=2) as response:
                return response.status == 200
        except (urllib.error.URLError, TimeoutError, OSError):
            time.sleep(0.4)
    return False


def show_error_dialog(message: str) -> None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showerror("Reliability App", message)
    root.destroy()


def run_server(host: str, port: int) -> None:
    ensure_storage()
    serve(app, host=host, port=port, threads=4)


def main() -> int:
    host = "127.0.0.1"
    requested_port = int(os.environ.get("PORT", "0") or "0")
    port = requested_port if requested_port > 0 else find_free_port()
    os.environ.setdefault("HOST", host)
    os.environ["PORT"] = str(port)
    os.environ.setdefault("TRUST_PROXY", "0")
    os.environ.setdefault("PREFERRED_URL_SCHEME", "http")

    server_thread = threading.Thread(target=run_server, args=(host, port), daemon=False)
    server_thread.start()

    url = f"http://{host}:{port}/"
    health_url = f"http://{host}:{port}/healthz"
    if not wait_for_server(health_url):
        show_error_dialog(
            "The Reliability app could not start correctly. Please close the app and try again."
        )
        return 1

    webbrowser.open(url)

    try:
        server_thread.join()
    except KeyboardInterrupt:
        return 0
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
