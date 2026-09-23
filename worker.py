"""Render background worker for persistent invoice-analysis jobs."""

import signal
import threading

from app import run_invoice_analysis_worker


def main():
    stop_event = threading.Event()

    def stop_worker(*_args):
        stop_event.set()

    signal.signal(signal.SIGINT, stop_worker)
    signal.signal(signal.SIGTERM, stop_worker)
    run_invoice_analysis_worker(stop_event)


if __name__ == "__main__":
    main()
