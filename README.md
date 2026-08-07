# Property Status Report Dashboard

Standalone Next.js version of the Claude artifact, ready to deploy on Vercel.

## What changed from the artifact

Everything is identical except how the AI features (Decision Support, Multi-Year Performance
Report) talk to Claude:

- **In the Claude.ai artifact:** calls `api.anthropic.com` directly. This only works there because
  the platform authenticates the request invisibly — there's no key in the code.
- **In this project:** the browser calls `/api/claude` (a Next.js serverless function you own,
  in `app/api/claude/route.js`). That function holds your real Anthropic API key server-side and
  forwards the request to Anthropic. The key is never sent to the browser or visible in any
  client-side code — this is the only safe way to do this outside of claude.ai.

Everything else — Excel parsing, charts, reconciliation, exception detection, the whole UI — runs
exactly as it did in the artifact. I ran a full production build and smoke test on this project
before handing it off; it compiles clean and serves correctly.

## Deploy to Vercel (about 5 minutes)

1. **Get an Anthropic API key** at https://console.anthropic.com (API usage is billed separately
   from any Claude.ai subscription — check current pricing there).

2. **Push this folder to GitHub:**
   ```bash
   cd kevin-dashboard
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/kevin-dashboard.git
   git push -u origin main
   ```

3. **Import into Vercel:**
   - Go to https://vercel.com/new
   - Select the GitHub repo you just pushed
   - Vercel auto-detects Next.js — you don't need to change any build settings

4. **Add the environment variable** (this is the step people forget):
   - In the Vercel project → **Settings → Environment Variables**
   - Add `ANTHROPIC_API_KEY` = your key from step 1
   - Apply it to Production (and Preview if you want previews to work too)

5. **Deploy.** Vercel builds and gives you a live URL (`your-project.vercel.app`).

If you skip step 4, the site still loads and the Excel/dashboard features work fine — only the two
AI cards will show a clear error ("ANTHROPIC_API_KEY is not set on the server...") instead of
failing silently.

## Local development

```bash
npm install
cp .env.local.example .env.local
# edit .env.local and paste your real key
npm run dev
```

Then open http://localhost:3000.

## Project structure

```
app/
  page.jsx              -> renders the dashboard
  layout.jsx            -> minimal HTML shell
  api/claude/route.js   -> serverless proxy to Anthropic (holds the API key)
components/
  MonthlyStatusReportComparator.jsx  -> the full dashboard (same code as the artifact)
```

## A note on cost

Every time someone opens a property tab or uploads a multi-year set of reports, this makes real,
billed calls to the Anthropic API (one for Decision Support per report, one for the Multi-Year
Performance Report when 2+ years are loaded). There's a concurrency limiter (max 2 simultaneous
calls) and retry-with-backoff built in, but there's no caching or rate-limiting of *how often*
a person can trigger new calls. Worth keeping in mind before sharing the live link widely.
