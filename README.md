# Deploying AMPIN_FTG_4 Dashboard to Render

This turns your local script into a real website with its own URL
(something like `https://ampin-ftg-4-dashboard.onrender.com`) that you —
or anyone you share the link with — can open from any browser, without
your own PC needing to be on.

## Why Render (and not Vercel)

This app keeps a real, persistent connection out to your FTP servers on
port 21 every time someone loads the dashboard. Vercel runs your backend
code as short-lived, stateless "serverless functions" with rotating IP
addresses — not a great fit for an app built around outbound FTP. Render
runs your app as a normal, always-on web service, same as running it on
a small cloud computer, which matches what this script actually needs.

---

## What you need first

1. A **GitHub account** (free) — Render deploys from a GitHub repo.
2. A **Render account** (free to sign up: https://render.com) — you'll
   add a paid plan only for the actual running service (~$7/month for
   "Starter", which stays on 24/7 — the free tier sleeps after 15 minutes
   of no traffic, which would silently break your 15-minute auto-refresh).

---

## Step 1 — Put the project on GitHub

1. Create a new **private** GitHub repository (keep it private since
   your `.env.template` shows the shape of your credentials, even though
   it doesn't contain the real password once you remove it — see Step 2).
2. Upload these files to the repo:
   - `dashboard.py`
   - `requirements.txt`
   - `render.yaml` (optional but recommended — lets Render auto-configure
     itself)
3. **Do NOT upload your real `.env` file** (the one with actual
   passwords) to GitHub, even in a private repo. Credentials go into
   Render's dashboard directly instead (Step 3) — never into source
   control.

---

## Step 2 — Connect Render to your repo

1. Log into https://dashboard.render.com
2. Click **New +** → **Web Service**
3. Connect your GitHub account if you haven't already, then select your
   repository.
4. Render should auto-detect `render.yaml` and pre-fill most settings.
   If it doesn't, fill in manually:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn dashboard:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
   - **Plan**: Starter (or higher) — NOT the free tier, for the reason
     above.

---

## Step 3 — Add your FTP credentials as environment variables

This replaces your local `.env` file. In the Render dashboard, go to your
service → **Environment** tab → **Add Environment Variable**, and add
each of these one at a time (same values as your local `.env`):

```
SECI3_HOST
SECI3_PORT
SECI3_USER
SECI3_PASS
SECI3_REMOTE_DIR

CNI_HOST
CNI_PORT
CNI_USER
CNI_PASS
CNI_REMOTE_DIR

SECI5_HOST
SECI5_PORT
SECI5_USER
SECI5_PASS
SECI5_REMOTE_DIR

POLL_INTERVAL_MINUTES
```

Click **Save Changes** — Render will redeploy automatically.

---

## Step 4 — Deploy

If you used `render.yaml`, Render builds and deploys automatically after
you connect the repo and fill in the environment variables. Otherwise,
click **Create Web Service** to trigger the first deploy.

Watch the **Logs** tab — you should see:
```
Starting dashboard, binding to 0.0.0.0:10000
Auto-refresh every 15.0 minute(s). Reading live from FTP — no local files saved.
```

Once it says **Live** at the top, your dashboard is reachable at the URL
Render shows you (top of the page, looks like
`https://ampin-ftg-4-dashboard.onrender.com`).

---

## Things that work differently online vs. on your PC

- **The page must stay open in someone's browser tab** for auto-refresh
  to happen — same as running it locally. There's no background process
  fetching data when nobody has the page open; each refresh is triggered
  by a browser actively polling `/api/data`. If you want truly unattended
  background fetching independent of anyone viewing the page, that's a
  different architecture (a scheduled background job) — let me know if
  you want that instead.
- **Download Excel button** works the same way — it streams the file
  straight to whoever's browser clicked it.
- **No local files are saved** on Render either — same in-memory-only
  behavior as your local version.

---

## Updating the dashboard later

Any time you want to change something (e.g. column mappings, colors):
1. Edit `dashboard.py` locally (or ask me to make the change and give
   you the updated file).
2. Push the updated file to your GitHub repo.
3. Render auto-redeploys within a minute or two of detecting the change.

---

## Troubleshooting

- **Build fails on `xlrd` or `pandas`** — check `requirements.txt` was
  uploaded correctly; Render needs it to know what to install.
- **"No FTP credentials found" in logs** — means the environment
  variables from Step 3 weren't saved correctly; double-check spelling
  (`SECI3_HOST` not `SECI-3_HOST`, etc.) exactly matches what's in
  `dashboard.py`.
- **Service times out on large date ranges** — the start command already
  sets a 120-second timeout; if you regularly pull very long ranges
  (close to the 31-day cap), let me know and I can adjust this further or
  add a loading state for long requests.
- **Site loads but shows "Could not connect/login"** — check the FTP
  servers themselves are reachable from outside your office network (some
  plant networks restrict FTP access to specific known IPs — you
  confirmed this isn't the case for your servers, but worth re-checking
  if this comes up).