# Deploy to Railway

## Files you need (all included)

```
api.py                  ← The API server (1 endpoint)
form_filler_engine.py   ← The form filling logic (already tested)
master_data.md          ← Your company data
requirements.txt        ← Python packages
Dockerfile              ← Container setup
```

## Steps

### 1. Create a GitHub repo
Push all 5 files above to a new GitHub repo.

### 2. Deploy on Railway
- Go to [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub Repo**
- Select your repo
- Railway auto-detects the Dockerfile and deploys

### 3. Set environment variable
In Railway → your service → **Variables** tab → add:
```
ANTHROPIC_API_KEY = sk-ant-your-key-here
```

### 4. Get your URL
Railway gives you a public URL like:
```
https://your-app-name.up.railway.app
```

### 5. Test it
```bash
curl -X POST https://your-app-name.up.railway.app/fill \
  -F "file=@Vendor_Registration_Form.xlsx" \
  --output FILLED_form.xlsx
```

Upload a file, get the filled file back. That's the whole API.

## Updating your company data
Edit `master_data.md` in GitHub → push → Railway auto-redeploys.

## API Reference

### `GET /health`
Returns `{"status": "ok"}` — use for uptime monitoring.

### `POST /fill`
- **Body**: multipart form-data with a `file` field
- **Accepts**: .xlsx, .docx, .pdf
- **Returns**: the filled file in the same format
- **On error**: JSON with error message

## Connecting to n8n Cloud
In n8n, use an **HTTP Request** node:
- Method: POST
- URL: `https://your-app-name.up.railway.app/fill`
- Body: Form-Data → add field `file` → set to binary from previous node
- Response: Binary data (the filled file)
