# Jerry JD Generator — Setup & Deployment

A web app that generates Jerry-standard job descriptions from a role spec.
Claude writes and audits the JD; the Python linter checks it deterministically;
any fails are auto-fixed before you download the final .docx.

---

## What you need

- An Anthropic API key (starts with `sk-ant-`)
- Python 3.9+

---

## Option A — Run locally (one person, quick start)

```bash
# 1. Install dependencies
pip3 install -r requirements.txt --break-system-packages

# 2. Run
streamlit run app.py
```

Opens at http://localhost:8501. Enter your API key in the sidebar.

---

## Option B — Deploy for the whole team (recommended)

### Streamlit Community Cloud (free, easiest)

1. Push this folder to a **private** GitHub repo.
2. Go to share.streamlit.io → New app → select the repo → `app.py`.
3. In **Advanced settings → Secrets**, add:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```
4. Deploy. Everyone on the team visits the URL — no setup needed.

### Replit (also easy, no GitHub needed)

1. Go to replit.com → Create → Import from GitHub (or upload files manually).
2. Add `ANTHROPIC_API_KEY` to Replit Secrets.
3. Click Run. Share the URL with your team.

---

## Files

| File | Purpose |
|---|---|
| `app.py` | Main Streamlit app — all logic lives here |
| `lint_jd.py` | Deterministic linter — checks 14 mechanical rules |
| `requirements.txt` | Python dependencies |

---

## How it works

1. Recruiter pastes the role spec and job title, clicks **Generate JD**
2. Claude generates the JD and performs a full manual audit (all 26 Hard Rules)
3. The Python linter runs deterministically on the generated `.docx`
4. If any hard fails are found, they're automatically sent back to Claude to fix
5. Recruiter downloads the final `.docx` and reviews the audit report
