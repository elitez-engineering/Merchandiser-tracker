# PMI SKU OCR worker

Cloudflare Worker: reads a merchandiser's dispenser photo via Claude vision and
returns which SKUs (from `data/pmi-sku-reference.json`) it can see, with price
if legible. Called by the "Run OCR" button in the PMI Intake panel — every
result is a suggestion the admin must confirm before it's saved.

## Deploy (not yet done — needs your go-ahead)

```
cd worker
npm install
npx wrangler login          # if not already logged in
npx wrangler secret put ANTHROPIC_API_KEY   # paste the key when prompted
npx wrangler deploy
```

Wrangler prints the deployed URL (`https://pmi-sku-ocr.<subdomain>.workers.dev`).
Paste that into `PMI_OCR_ENDPOINT` near the top of `../pmi-intake.js` and
commit — OCR is disabled with an explanatory alert until this is set.

CORS is locked to `https://elitez-engineering.github.io` (the app's GitHub
Pages origin) in `src/index.ts` — update `ALLOWED_ORIGIN` there if the app
ever moves.
