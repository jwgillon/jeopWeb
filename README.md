# JeoparTy AI Generator

A web app that generates custom Jeopardy `.pptm` game files using the Gemini API.

## How it works

1. User enters their Gemini API key + theme in the browser
2. Browser calls Gemini directly (key never touches the server)
3. JSON game data is sent to the FastAPI backend
4. Backend injects data into the `.pptm` template and returns the file
5. User downloads and plays!

## Project structure

```
jeoparty-app/
├── main.py              # FastAPI backend
├── index.html           # Frontend (served by FastAPI)
├── template.pptm        # Your Jeopardy template (required!)
├── requirements.txt
├── render.yaml          # Render.com deployment config
└── README.md
```

## Local setup

```bash
# 1. Copy your template into this folder
cp /path/to/Jeopardy_-_Generator_-_Final.pptm ./template.pptm

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run locally
uvicorn main:app --reload --port 8000

# 4. Open http://localhost:8000
```

## Deploy to Render.com (free tier)

### Step 1 — Push to GitHub
```bash
git init
git add .
git commit -m "Initial commit"
# Create a repo on github.com, then:
git remote add origin https://github.com/YOUR_USERNAME/jeoparty-generator.git
git push -u origin main
```

> ⚠️ Make sure `template.pptm` is committed to the repo. It needs to be
> present at deploy time. If the file is large, Git LFS is an option.

### Step 2 — Create a Render Web Service

1. Go to [render.com](https://render.com) → **New → Web Service**
2. Connect your GitHub repo
3. Settings:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT`
   - **Instance Type**: Free
4. Click **Create Web Service**
5. Wait ~2 min for deploy → your app is live!

### Step 3 — Use it

Visit your Render URL, enter your Gemini API key, type a theme, click Generate.

## Notes

- The free Render tier spins down after 15 min of inactivity. First request
  after idle takes ~30 seconds. Upgrade to Starter ($7/mo) to avoid this.
- The Gemini API key is used client-side only and never logged or stored.
- Generated `.pptm` files are temporary and deleted after serving.
