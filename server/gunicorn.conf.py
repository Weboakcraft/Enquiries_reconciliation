"""Gunicorn settings for the hosted Flask build.

Reconciling a large week and building the workbook are CPU-bound and can run
for a few seconds, so the timeout is generous and the worker count is small.
"""
import multiprocessing
import os

bind = f"0.0.0.0:{os.environ.get('PORT', '5000')}"
workers = int(os.environ.get("WEB_CONCURRENCY", min(4, multiprocessing.cpu_count() * 2 + 1)))
threads = 2
timeout = 180
graceful_timeout = 30
keepalive = 5
max_requests = 200
max_requests_jitter = 40
accesslog = "-"
errorlog = "-"
