# Excel Copilot — Public Hosting & Analytics (Session Context)

Companion to [PROJECT_CONTEXT.md](PROJECT_CONTEXT.md) and the other session-context files. Records the request, the paths that were ruled out and why, what's actually been done in the repo, and exactly where the live deployment stands so a future session can pick this up without re-deriving it.

---

## 1. Request

Host the app publicly on a free tier, and add analytics covering: user location, device, number of users, and activity across the app (which tools get used, uploads/downloads).

## 2. Hosting paths considered and ruled out

- **Vercel** — rejected. Vercel's Python support is serverless functions: stateless, ephemeral, no guaranteed instance reuse between requests, no reliable background threads. This app's design (§5.1 of `PROJECT_CONTEXT.md`) depends on the opposite: an in-memory `SESSIONS` dict that must survive across separate HTTP calls (upload → pick sheet → apply tool → download), plus background threads for lazy sheet loading and background-built downloads that the client polls. On Vercel these would intermittently just fail, not merely degrade.
- **Render** — technically fine (zero code changes, same architecture works as-is), but Render now asks for a card on signup even for the free web-service tier. The user wanted to avoid that, which is what led to exploring Firebase.
- **Firebase Hosting (Spark/free plan) alone** — rejected: it's a static-file CDN only, cannot run a Python process at all.
- **Firebase Cloud Functions** — rejected: same statelessness problem as Vercel, *and* requires the Blaze (pay-as-you-go) plan (credit card on file) regardless.
- **Firebase App Hosting** (the product at the `firebase.google.com/docs/app-hosting` link the user originally found) — rejected: it has built-in support only for Next.js/Angular. Any other framework, including Python/FastAPI, requires containerizing it yourself and deploying via **Terraform** to get it running underneath App Hosting — more setup than just using Cloud Run directly — and it *still* requires the Blaze plan. It does not remove the card requirement, it just adds Docker + Terraform on top of it.
- **Chosen path: containerize the app → deploy straight to Cloud Run → optionally front it with Firebase Hosting's URL-rewrite integration.** Same underlying Google infrastructure as Firebase App Hosting, fully supports arbitrary containers, no Terraform needed, and Firebase Hosting can proxy all traffic to it via a plain `firebase.json` rewrite (`"run": {"serviceId": ..., "region": ...}`) — this is a separate, older, well-documented Firebase Hosting feature, distinct from the newer "Firebase App Hosting" product.
- The Blaze plan (and therefore a card) turned out to be unavoidable for *any* option capable of running a real persistent Python server for free (Render, Cloud Run, Fly.io all require one, usually for anti-abuse verification rather than billing). The user accepted a card for verification only, since Cloud Run's free tier keeps the actual bill at $0 for this workload.

## 3. Analytics: Google Analytics 4

GA4 was chosen over Umami Cloud/self-hosted (the other options offered) for being fully free with no event-volume cap, and because "Firebase Analytics" for a web app *is* GA4 under the hood — enabling Analytics during Firebase project creation auto-creates a GA4 property, so there's nothing separate to build on the Firebase side.

### What's implemented (already committed and pushed to `origin/main`)
- `frontend/index.html` — GA4 loader script (`gtag.js`), now with the **real Measurement ID `G-8RR1HZ0H2Q`** wired in (the user registered a Firebase web app and plugged this in themselves, commit `d8ae3c7831069c06e505a3d1c393c1ab37e52d9d`, "Add GA4 MID & mobile overflow changes" — also included some unrelated mobile CSS overflow fixes).
- `frontend/app.js` — a `trackEvent(name, params)` helper (no-ops safely if `gtag` isn't loaded) wired into four events:
  - `upload_workbook` (on successful upload, with sheet count + mode)
  - `download_workbook` (on successful download, with mode)
  - `tool_opened` (when any sidebar tool/wizard is opened, with which tool)
  - `operation_applied` (when the Formula Builder/wizard successfully applies an operation, with which operation)
- **Not yet instrumented**: Duplicates (`dup.js`), SQL (`sql.js`), and the sheet-tools actions (`tools.js` — summary rows, copy-to-sheet, add/rename column). Offered twice, not yet requested. The same one-line `trackEvent("name", {...})` pattern applies if/when wanted.

### Where to view it
- `analytics.google.com` → the "Excel Copilot Web" GA4 property → Realtime (fastest way to confirm it's working), or Reports → Tech/Engagement for device/location/event breakdowns over time.
- Inside Firebase console → left sidebar → **Analytics** → **Dashboard** shows the same GA4 data embedded.

## 4. Repository cleanup (done)

Before pushing publicly, cleaned up things flagged in `PROJECT_CONTEXT.md` §11 as outstanding:
- Added a root `.gitignore` (`venv/`, `backend/__pycache__/`, `.claude/`, `*.xlsx`/`*.xlsm`, `sample_data/`).
- Untracked (not deleted from disk) the two personal-looking `.xlsx` files, `sample_data/`, `.claude/settings.local.json`, and compiled `.pyc` files.
- Committed and pushed as `902fec6` ("Add .gitignore, untrack personal data files, wire in GA4 hooks").
- **Caveat, not yet addressed**: the repo (`github.com/ahsanaliabbasi/excel-copilot`, confirmed public) had two earlier commits (`fd66259`, `88bd555`) that already contain those xlsx files — untracking them going forward does not remove them from GitHub's history. If that data is actually sensitive, it needs a history rewrite (`git filter-repo` + force-push), which is destructive and was intentionally not done without explicit confirmation.

## 5. Containerization (done, validated locally)

- `Dockerfile` (repo root) — single-stage, `python:3.11-slim`, installs `backend/requirements.txt`, copies both `backend/` and `frontend/` into the image (required because `main.py`'s `FRONTEND_DIR` is computed as one level above the backend file + `"frontend"` — the container layout must keep `backend/` and `frontend/` as siblings under `/app`). `CMD` runs `uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080}` (shell form, so `$PORT` expands — needed for Cloud Run, which injects `PORT`).
- `.dockerignore` — excludes venvs, caches, `.git`, `.claude`, `Planning/`, xlsx/sample data.
- **Validated in this session**: built the image locally (`docker build -t excel-copilot:local .` — succeeded, ~476 MB), ran it (`docker run -p 8090:8080 -e PORT=8080 ...`), confirmed `/api/health` responded and the frontend served (200), and did a **real file upload** against `sample_data/Data_Exercise_-_Ahsan_Ali.xlsx` through the running container — it processed correctly (147/135-row sheets read back). The Dockerfile is confirmed working end-to-end, not just theoretically correct.

## 6. Firebase/GCP project state (as of 2026-10-03)

- Firebase project created: display name **"excel-copilot"**, **Project ID `excel-copilot-ai`** (this is the exact string `gcloud` needs — visible in the console URL `console.firebase.google.com/project/excel-copilot-ai/...`).
- Google Analytics enabled during project creation, using **"Default Account for Firebase"** (recommended, standard choice) — this auto-created the linked GA4 property.
- A Web app was registered in the Firebase project (needed specifically to get a GA4 Measurement ID, since the ID lives on a per-app data stream, not the bare property) — this produced `G-8RR1HZ0H2Q`, already live in `index.html`.
- **Billing: still on the Spark (free) plan.** The user explicitly skipped the Blaze upgrade at one point. They were told clearly that **Cloud Run cannot deploy at all on Spark** — this is a hard blocker, not a soft limit, and was pending resolution as of the last message in this session.
- `gcloud init` was run locally, authenticated as `ahsandotabbasi@gmail.com`, and was at the "pick cloud project to use" prompt — not yet confirmed pointed at `excel-copilot-ai`.
- **`gcloud run deploy` has not been run yet** in this session (no live Cloud Run URL exists yet). This is the next concrete step once Blaze is confirmed enabled.

## 7. Exact next steps (pick up here)

1. Confirm/complete the **Blaze upgrade**: Firebase console → bottom-left "Upgrade" (or Settings → Usage and billing) → Blaze → attach a billing account.
2. In the `gcloud init` terminal: pick project by ID, enter `excel-copilot-ai`; set default region to `us-central1` if asked.
3. `gcloud services enable run.googleapis.com artifactregistry.googleapis.com cloudbuild.googleapis.com`
4. From the repo root:
   ```bash
   gcloud run deploy excel-copilot \
     --source . \
     --region us-central1 \
     --allow-unauthenticated \
     --memory 1Gi \
     --max-instances 1 \
     --no-cpu-throttling
   ```
   Two flags matter specifically for this app's architecture, not just as generic best practice:
   - `--max-instances 1` — the in-memory `SESSIONS` dict only exists in one process; letting Cloud Run scale out would split requests across instances that don't share that state, silently "losing" a user's session.
   - `--no-cpu-throttling` — Cloud Run's default only allocates CPU while actively handling a request. This app's background sheet-loading thread and background download-builder (polled from separate requests) need CPU between requests to keep making progress; without this flag that work can stall.
5. Test the resulting `https://excel-copilot-xxxxx-uc.a.run.app` URL end to end (upload → tool → download).
6. Optional: front it with Firebase Hosting for a nicer URL — `firebase init hosting` (existing project, any public dir, doesn't matter since everything gets rewritten) → edit `firebase.json`:
   ```json
   { "hosting": { "public": "public", "rewrites": [
       { "source": "**", "run": { "serviceId": "excel-copilot", "region": "us-central1" } }
   ]}}
   ```
   → `firebase deploy --only hosting`.

## 8. Open items

- Confirm Cloud Run deploy succeeds and capture the live URL here for next time.
- Decide whether to purge the old xlsx files from git history (destructive, needs explicit go-ahead).
- Optionally instrument `dup.js`, `sql.js`, `tools.js` with the same `trackEvent` pattern for fuller activity coverage.
- No auth on the app — fine for a demo link, not for sensitive data, if the Cloud Run URL gets shared widely.
- Cost note: Cloud Run free tier (2M requests, 360k GB-seconds, 180k vCPU-seconds/month) should comfortably cover demo-level traffic even with `--no-cpu-throttling` keeping the single instance warm while active.
