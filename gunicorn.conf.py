# gunicorn.conf.py
# Optimised for Render.com free tier (512MB RAM, single worker)
workers     = 1       # single worker — avoids memory doubling
worker_class = "sync" # sync worker — most memory efficient
timeout     = 300     # 5 min timeout — pipeline needs time to run
max_requests = 1      # restart worker after each request — prevents memory leaks
graceful_timeout = 30
