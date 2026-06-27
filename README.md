# Off Axis Entertainment GFX QC

Streamlit web app: upload broadcast graphics stills, read timecode + on-screen
text via the OpenAI vision API, run QC checks, and download a color-coded Excel
log with each frame embedded.

## How it works
Each uploaded frame is sent to the OpenAI API (model: gpt-4.1-mini) to extract
the timecode and on-screen text and to do title/action-safe checks. Additional
offline checks (spelling, name consistency, capitalization/accent drift, price
formatting, reserved bug zone) run locally. Output is an .xlsx QC log.

## Run locally
1. `python3 -m venv .venv && source .venv/bin/activate`
2. `pip install -r requirements.txt`
3. Create `.streamlit/secrets.toml` from `secrets.toml.example` and fill in keys.
4. `streamlit run streamlit_app.py`

## Deploy (Streamlit Community Cloud) — free
1. Push this folder to a **private** GitHub repo.
2. Go to https://share.streamlit.io , sign in with GitHub, click **New app**.
3. Pick this repo, branch `main`, main file `streamlit_app.py`. Deploy.
4. In the app's **Settings -> Secrets**, paste:
   ```
   OPENAI_API_KEY = "sk-...your-team-key..."
   APP_PASSWORD   = "your-shared-team-password"
   ```
5. Share the app URL + the password with your team. Done.

## Pushing updates
Edit the code, then:
```
git add -A
git commit -m "your change"
git push
```
Streamlit Cloud redeploys automatically within ~1 minute. Everyone gets the
update on their next page load — no reinstalling anything.

## Security notes
- The OpenAI key lives only in the host's Secrets, never in the repo or the
  browser. The `.gitignore` blocks `secrets.toml` from ever being committed.
- `APP_PASSWORD` gates the public URL with a single shared team password.
- Frames are sent to the OpenAI API for analysis (API data is not used for
  model training by default).
