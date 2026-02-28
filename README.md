---
title: AI-Beta
emoji: 📊
colorFrom: blue
colorTo: green
sdk: docker
app_port: 7860
---

# AI-Beta Company Scorer

Web app that:

1. Accepts pasted company names.
2. Resolves ambiguous names and asks user confirmation.
3. Researches and scores 5 AI-Beta dimensions.
4. Calculates AI-Beta score numerically.
5. Adds one-sentence comments with scoring rationale.
6. Exports results to Excel.

## Setup

```bash
cd /Users/angelaluget/Projects/HarryCodex
cp .env.example .env
npm install
npm run dev
```

Open http://localhost:3000.

## Notes

- Requires `OPENAI_API_KEY` in `.env`.
- Uses OpenAI Responses API with web search for research.
- Resolve/analyze now stream live progress updates in the UI.
- Optional tuning:
  - `RESOLVE_CONCURRENCY` (default `3`)
  - `ANALYZE_CONCURRENCY` (default `2`)
- Formula:

```text
AI Beta = ((Functional Susceptibility + Digital Susceptibility) * Resilience)
          + AI Infrastructure Upside + AI Competitiveness Upside
```

## Hugging Face Space

- Runtime: Docker
- Port: `7860`
- Required Space secret: `OPENAI_API_KEY`

## Deploy

This app can be deployed to Render, Railway, Fly.io, or Hugging Face Spaces as a Node web service.

- Build command: `npm install`
- Start command: `npm start`
- Environment variables: `OPENAI_API_KEY`, optional `OPENAI_MODEL`, `PORT`, `RESOLVE_CONCURRENCY`, `ANALYZE_CONCURRENCY`
