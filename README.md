# Omega Slide Maker

Node-powered web app that builds slide deck outlines with OpenAI and applies the optional example PPTX theme.

## Setup

1. Duplicate `.env.example` as `.env` (or export the variables another way).
2. Provide your OpenAI key: `OPENAI_API_KEY=sk-...`.
3. Run `npm install` (loads the tiny dependency used to read `.env`).
4. (Optional) Override `OPENAI_MODEL` or `PORT` if you need different defaults.

## Run locally

```bash
npm run dev
```

Then open [http://localhost:3000](http://localhost:3000) and describe the deck you need.

## Notes

- The backend proxies OpenAI requests so your API key never lives in the browser.
- If the key is missing or the request fails, the app generates a local fallback outline so you still get a draft.
- The server automatically retries supported OpenAI models (`OPENAI_MODEL`, `gpt-4o-mini`, `gpt-4o`, `gpt-4.1-mini`) so you get a live outline even when the first choice is unavailable.
- Generated outlines can be exported as a ready-to-import Google Slides `.pptx` file directly from the app.
- Upload your own `.pptx` template to reuse your branded theme when generating and exporting decks.
- A prebuilt copy of PptxGenJS ships in `vendor/pptxgenjs` so PPTX exports work even without npm registry access.
- The most recent deck is exposed via `window.omegaDeck.getState()` and `window.omegaDeck.buildExportPayload()` for future export tooling.
