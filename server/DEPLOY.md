# Hosting the Flask build

The Flask app is the original: a small server that parses the uploads, keeps the
run history as JSON files on disk, and serves the same interface. It is here for
anyone who wants a single shared installation rather than the browser build in
`../web`.

**Read this first.** The app has no login of any kind. Every visitor can upload,
read every saved run and download the full lead register — names, phone numbers,
cities. That is fine on a laptop. It is not fine on a public URL. Before putting
this on the internet, put something in front of it: your host's password
protection or IP allow-list, a reverse proxy with basic auth, or a private
network. The browser build has no server to reach, so it does not carry this risk.

The run history lives in `data/runs/`. It is plain JSON on a disk, so a host
without a persistent volume loses the week-on-week trend on every deploy. Every
option below mounts or reserves one.

## Docker (any host)

```bash
cd server
docker build -t oakcraft-reconciler .
docker run -d --name reconciler -p 8000:8000 \
  -v oak-data:/app/data --restart unless-stopped oakcraft-reconciler
```

Then open http://localhost:8000. `oak-data` is a named volume, so the history
survives `docker rm` and image rebuilds. Back it up:

```bash
docker run --rm -v oak-data:/data -v "$PWD":/backup alpine \
  tar czf /backup/reconciler-history.tar.gz -C /data .
```

## Render

`render.yaml` is a ready blueprint. In Render: **New → Blueprint**, pick this
repo, and it reads the file. It provisions a Docker web service with a 1 GB disk
mounted at `/app/data`.

Two things to change after it is up:
- The free plan has no persistent disk. The blueprint asks for `starter` for
  that reason; on free, the history resets on each deploy.
- Add access protection before sharing the URL. See the warning above.

## Fly.io

```bash
cd server
fly launch --no-deploy          # accept the detected Dockerfile
fly volumes create oak_data --size 1
```

Then add to the generated `fly.toml`:

```toml
[[mounts]]
  source = "oak_data"
  destination = "/app/data"

[http_service]
  internal_port = 8000
```

and `fly deploy`.

## Railway / Heroku-style buildpacks

`Procfile` covers these:

```
web: gunicorn -c gunicorn.conf.py app:app
```

Set the root directory to `server`. Attach a persistent volume at `/app/data` —
without one, `data/runs/` is wiped whenever the dyno restarts.

## Plain VPS with systemd

```ini
# /etc/systemd/system/reconciler.service
[Unit]
Description=OakCraft Lead Reconciler
After=network.target

[Service]
User=www-data
WorkingDirectory=/srv/reconciler/server
Environment=PORT=8000
ExecStart=/srv/reconciler/.venv/bin/gunicorn -c gunicorn.conf.py app:app
Restart=always

[Install]
WantedBy=multi-user.target
```

Put nginx or Caddy in front for TLS, and add basic auth there.

## Settings

| Variable | Default | What it does |
|---|---|---|
| `PORT` | 8000 in Docker, 5000 locally | Port the server binds |
| `WEB_CONCURRENCY` | up to 4 | Gunicorn worker processes |

Upload size is capped at 40 MB per request in `app.py` (`MAX_MB`). If you raise
it, raise your proxy's body-size limit to match — nginx's `client_max_body_size`
defaults to 1 MB and will reject uploads long before Flask sees them.

## Checking it works

```bash
cd server
pip install -r requirements.txt
python selftest.py sample-data/sample_meta_leads_previous_week.xlsx
```

That runs the whole pipeline on the command line and prints the control totals.
If it balances, the deployment will produce the same numbers.
