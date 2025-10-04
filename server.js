require('dotenv').config();

const http = require('node:http');
const fs = require('node:fs/promises');
const path = require('node:path');
const { URL } = require('node:url');

const PORT = Number(process.env.PORT) || 3000;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || '';

const PUBLIC_DIR = path.join(__dirname, 'public');
const OPENAI_API_URL = 'https://api.openai.com/v1/responses';
const OPENAI_MODEL_ENV = (process.env.OPENAI_MODEL || '').trim();
const FALLBACK_OPENAI_MODELS = ['gpt-4o-mini', 'gpt-4o', 'gpt-4.1-mini'];
const MAX_AI_SLIDES = 12;
const EXAMPLE_THEME_KEY = 'example-pptx';
const CUSTOM_THEME_KEY = 'custom-upload';
const LOCAL_PPTX_MODULE = path.join(__dirname, 'vendor', 'pptxgenjs', 'index.js');

let pptxModulePromise = null;

function getOpenAiModelCandidates() {
  const candidates = [];
  if (OPENAI_MODEL_ENV) {
    candidates.push(OPENAI_MODEL_ENV);
  }
  for (const fallback of FALLBACK_OPENAI_MODELS) {
    if (!candidates.includes(fallback)) {
      candidates.push(fallback);
    }
  }
  return candidates;
}

function shouldRetryWithAlternateModel(error) {
  if (!error) {
    return false;
  }
  const code = String(error.code || '').toLowerCase();
  if (code === 'model_not_found' || code === 'invalid_model' || code === '404') {
    return true;
  }
  const message = String(error.message || '').toLowerCase();
  if (message.includes('model') && (message.includes('not found') || message.includes('does not exist'))) {
    return true;
  }
  return false;
}

async function getPptxGen() {
  if (!pptxModulePromise) {
    pptxModulePromise = (async () => {
      try {
        const module = await import('pptxgenjs');
        const exported = module?.default || module;
        if (typeof exported !== 'function') {
          throw new Error('pptxgenjs module did not export a constructor.');
        }
        return exported;
      } catch (error) {
        console.warn('Falling back to bundled pptxgenjs module:', error?.message || error);
        // eslint-disable-next-line global-require, import/no-dynamic-require
        const localModule = require(LOCAL_PPTX_MODULE);
        const exported = localModule?.default || localModule;
        if (typeof exported !== 'function') {
          throw new Error('Bundled pptxgenjs module invalid.');
        }
        return exported;
      }
    })();
  }
  return pptxModulePromise;
}

function normalizeBullets(bullets) {
  if (!Array.isArray(bullets)) {
    return [];
  }
  return bullets
    .map((bullet) => (typeof bullet === 'string' ? bullet.trim() : ''))
    .filter(Boolean)
    .slice(0, 8);
}

function shapeSlidesForExport(slidesInput) {
  if (!Array.isArray(slidesInput)) {
    return [];
  }
  return slidesInput.map((slide, index) => {
    const title = typeof slide?.title === 'string' && slide.title.trim()
      ? slide.title.trim()
      : `Slide ${index + 1}`;
    const bullets = normalizeBullets(slide?.bullets || slide?.keyPoints || []);
    const notes = typeof slide?.speakerNotes === 'string' ? slide.speakerNotes.trim() : '';

    return { title, bullets, notes };
  });
}

async function generatePptxBuffer(deckPayload) {
  const PptxGenJS = await getPptxGen();
  const pptx = new PptxGenJS();
  pptx.layout = PptxGenJS?.Layouts?.LAYOUT_16x9 || 'LAYOUT_16x9';

  const slides = shapeSlidesForExport(deckPayload.slides);
  if (slides.length === 0) {
    throw Object.assign(new Error('At least one slide is required to export.'), {
      code: 'EMPTY_DECK',
    });
  }

  const summaryItems = Array.isArray(deckPayload.insights)
    ? deckPayload.insights
        .map((note) => (typeof note === 'string' ? note.trim() : ''))
        .filter(Boolean)
        .slice(0, 8)
    : [];

  slides.forEach((slide, index) => {
    const pptSlide = pptx.addSlide();
    pptSlide.addText(slide.title, {
      x: 0.6,
      y: 0.5,
      fontSize: 30,
      bold: true,
      color: '27304E',
    });

    if (slide.bullets.length > 0) {
      pptSlide.addText(
        slide.bullets.map((text) => ({ text })),
        {
          x: 0.9,
          y: 1.5,
          fontSize: 18,
          color: '3D4470',
          lineSpacing: 28,
          bullet: true,
        },
      );
    }

    if (slide.notes) {
      pptSlide.addNotes(slide.notes);
    }
  });

  if (summaryItems.length > 0) {
    const summarySlide = pptx.addSlide();
    summarySlide.addText('Key Takeaways', {
      x: 0.6,
      y: 0.5,
      fontSize: 30,
      bold: true,
      color: '27304E',
    });
    summarySlide.addText(
      summaryItems.map((text) => ({ text })),
      {
        x: 0.9,
        y: 1.6,
        fontSize: 18,
        color: '3D4470',
        lineSpacing: 28,
        bullet: true,
      },
    );
  }

  const result = await pptx.write('nodebuffer');
  if (Buffer.isBuffer(result)) {
    return result;
  }
  if (ArrayBuffer.isView(result)) {
    return Buffer.from(result.buffer);
  }
  if (result instanceof ArrayBuffer) {
    return Buffer.from(result);
  }
  if (typeof result === 'string') {
    return Buffer.from(result, 'base64');
  }
  throw new Error('Unexpected PPTX output format.');
}

function buildExportFilename(brief) {
  if (typeof brief !== 'string' || !brief.trim()) {
    return 'omega-deck';
  }

  const cleaned = brief
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .slice(0, 8)
    .join('-');

  return cleaned || 'omega-deck';
}

const AI_OUTPUT_GUIDE = `Return a JSON object with the following structure:
{
  "slides": [
    {
      "id": "slide-1",
      "title": "Concise slide title",
      "keyPoints": [
        "Bullet point 1",
        "Bullet point 2"
      ],
      "speakerNotes": "Optional concise speaker note for presenters."
    }
  ],
  "insights": ["Optional list of key insights for the presenter"],
  "outline": ["Array listing slide titles in order"],
  "summary": "Optional single paragraph summary of the deck"
}
Rules:
- Provide between 6 and ${MAX_AI_SLIDES} slides.
- keyPoints must contain 2 to 5 short bullet strings.
- speakerNotes is optional but preferred; omit it when not relevant.
- Use plain text only. Do not include markdown, explanations, or additional keys.`;

const CONTENT_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

const slideTitles = [
  'Opening',
  'Problem',
  'Solution',
  'Product',
  'Proof',
  'Demo',
  'Metrics',
  'Roadmap',
  'Pricing',
  'CTA',
];

function sendJson(res, statusCode, data) {
  const payload = JSON.stringify(data);
  res.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(payload),
  });
  res.end(payload);
}

async function readBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString('utf-8');
}

function synthesizeBullets(text) {
  return text
    .split(/[\n\.]/)
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, 4);
}

function quickDraftDeck(brief) {
  const insights = synthesizeBullets(brief);
  const slideCount = Math.max(6, Math.min(10, insights.length + 4));
  const outline = slideTitles.slice(0, slideCount);

  const keywords = brief
    .split(/[\n,]/)
    .map((part) => part.trim())
    .filter(Boolean);

  const slides = outline.map((title, index) => {
    const bullets = [];
    const start = index * 2;
    const segment = keywords.slice(start, start + 2);

    if (segment.length === 0) {
      bullets.push('Highlight the most important takeaways for this section.');
    } else {
      segment.forEach((keyword) => bullets.push(`Focus on ${keyword.toLowerCase()}.`));
    }

    return {
      id: `slide-${index + 1}`,
      title,
      position: index + 1,
      displayTitle: `${index + 1}. ${title}`,
      bullets,
    };
  });

  return { slides, insights, outline };
}

function adaptAiSlides(aiSlides) {
  if (!Array.isArray(aiSlides)) {
    return [];
  }

  return aiSlides.slice(0, MAX_AI_SLIDES).map((slide, index) => {
    const title = (slide?.title || slide?.heading || `Slide ${index + 1}`).trim();
    const bullets = Array.isArray(slide?.keyPoints)
      ? slide.keyPoints
          .map((point) => (typeof point === 'string' ? point.trim() : ''))
          .filter(Boolean)
          .slice(0, 5)
      : [];

    if (bullets.length === 0) {
      bullets.push('Highlight the key takeaway for this section.');
    }

    return {
      id: slide?.id || `slide-${index + 1}`,
      title,
      position: index + 1,
      displayTitle: `${index + 1}. ${title}`,
      bullets,
      speakerNotes: typeof slide?.speakerNotes === 'string' ? slide.speakerNotes.trim() : undefined,
    };
  });
}

function collectInsightsFromAi(aiResponse, brief) {
  const direct = Array.isArray(aiResponse.speakerNotes)
    ? aiResponse.speakerNotes
        .map((note) => (typeof note === 'string' ? note.trim() : ''))
        .filter(Boolean)
    : [];

  if (direct.length > 0) {
    return direct.slice(0, 8);
  }

  const fromSlides = Array.isArray(aiResponse.slides)
    ? aiResponse.slides
        .map((slide) => {
          if (typeof slide?.speakerNotes === 'string') {
            return slide.speakerNotes.trim();
          }
          if (Array.isArray(slide?.keyPoints)) {
            return slide.keyPoints
              .map((point) => (typeof point === 'string' ? point.trim() : ''))
              .filter(Boolean)
              .slice(0, 1)
              .join(' ');
          }
          return '';
        })
        .filter(Boolean)
    : [];

  if (fromSlides.length > 0) {
    return fromSlides.slice(0, 8);
  }

  return synthesizeBullets(brief);
}

async function requestDeckFromOpenAI(brief, themeOptions, model) {
  const themeKey = themeOptions?.key || 'default';
  const themeName = themeOptions?.name || 'custom theme';
  const systemPrompt = 'You are an expert presentation designer. Create concise Google Slides outlines with strong storytelling. Always respond with compact JSON that matches the requested structure. Do not include commentary outside JSON.';
  let themeDirective = 'Use the default Omega theme: modern, minimal, data-forward, and easy to adapt.';
  if (themeKey === EXAMPLE_THEME_KEY) {
    themeDirective = 'Match the mood, palette, and typography of the example pptx theme. Reference slide roles that benefit from that style.';
  } else if (themeKey === CUSTOM_THEME_KEY) {
    themeDirective = `Adapt the outline to align with the uploaded PowerPoint template named "${themeName}". Assume vibrant, on-brand visuals that mirror that deck's structure.`;
  }

  const response = await fetch(OPENAI_API_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${OPENAI_API_KEY}`,
    },
    body: JSON.stringify({
      model,
      input: [
        {
          role: 'system',
          content: [{ type: 'input_text', text: systemPrompt }],
        },
        {
          role: 'user',
          content: [
            {
              type: 'input_text',
              text: `Slide deck brief:\n${brief}\n\nTheme request: ${themeDirective}\n\n${AI_OUTPUT_GUIDE}`,
            },
          ],
        },
      ],
      text: {
        format: { type: 'json_object' },
      },
      temperature: 0.5,
    }),
  });

  const data = await response.json().catch(() => null);

  if (!response.ok || !data) {
    const errorPayload = data?.error || {};
    const message = errorPayload.message || response.statusText || 'Unknown OpenAI API error';
    const error = new Error(`OpenAI API error (${response.status}): ${message}`);
    error.code = errorPayload.code || response.status;
    error.details = errorPayload;
    error.model = model;
    throw error;
  }

  const firstOutput = Array.isArray(data.output)
    ? data.output.find((item) => Array.isArray(item?.content))
    : null;
  const outputNode = firstOutput?.content?.find((part) => ['json', 'json_schema', 'output_text', 'text', 'json_object'].includes(part?.type));
  if (!outputNode) {
    const error = new Error('OpenAI response did not include JSON content.');
    error.code = 'EMPTY_OPENAI_RESPONSE';
    error.details = data;
    error.model = model;
    throw error;
  }

  let parsed;
  if (outputNode.type === 'json' || outputNode.type === 'json_schema' || outputNode.type === 'json_object') {
    parsed = outputNode.json || outputNode.data || outputNode.json_schema || outputNode.parsed || null;
  } else {
    const textPayload = outputNode.text ?? outputNode.output_text ?? outputNode.value;
    if (typeof textPayload !== 'string') {
      const error = new Error('OpenAI response text payload missing.');
      error.code = 'MISSING_OPENAI_TEXT';
      error.details = { data, outputNode };
      error.model = model;
      throw error;
    }
    parsed = JSON.parse(textPayload);
  }

  if (!parsed || typeof parsed !== 'object') {
    const error = new Error('OpenAI response JSON payload malformed.');
    error.code = 'INVALID_OPENAI_JSON';
    error.details = { data, parsed };
    error.model = model;
    throw error;
  }

  return {
    modelUsed: model,
    slides: parsed.slides || [],
    speakerNotes: parsed.speakerNotes || parsed.summary || [],
    raw: data,
  };
}

async function generateDeckWithOpenAI(brief, themeOptions) {
  const modelsToTry = getOpenAiModelCandidates();
  let lastError = null;

  for (const model of modelsToTry) {
    try {
      return await requestDeckFromOpenAI(brief, themeOptions, model);
    } catch (error) {
      lastError = error;
      if (shouldRetryWithAlternateModel(error)) {
        console.warn(`OpenAI model ${model} unavailable: ${error.message || error.code}`);
        continue;
      }
      throw error;
    }
  }

  throw lastError || new Error('No OpenAI models available.');
}

async function handleGenerate(req, res) {
  try {
    const body = await readBody(req);
    const payload = body ? JSON.parse(body) : {};
    const brief = typeof payload.brief === 'string' ? payload.brief.trim() : '';
    const customTheme = typeof payload.customTheme === 'string' ? payload.customTheme : '';
    const customThemeName = typeof payload.customThemeName === 'string' ? payload.customThemeName.trim() : '';
    const hasCustomTheme = Boolean(customTheme && customThemeName);
    const includeExampleTheme = hasCustomTheme ? true : Boolean(payload.includeExampleTheme);
    const themeKey = hasCustomTheme ? CUSTOM_THEME_KEY : includeExampleTheme ? EXAMPLE_THEME_KEY : 'default';
    const themeOptions = { key: themeKey, name: hasCustomTheme ? customThemeName : null };

    if (!brief) {
      return sendJson(res, 400, {
        error: {
          message: 'Brief is required.',
          code: 'BRIEF_REQUIRED',
        },
      });
    }

    if (!OPENAI_API_KEY) {
      const fallback = quickDraftDeck(brief);
      return sendJson(res, 200, {
        provider: 'fallback',
        deck: fallback,
        themeKey,
        customThemeName: hasCustomTheme ? customThemeName : null,
        modelUsed: null,
        rawModel: { error: { message: 'OPENAI_API_KEY missing on server' } },
        error: {
          message: 'OpenAI API key missing. Generated a quick draft locally instead.',
          code: 'MISSING_OPENAI_API_KEY',
        },
      });
    }

    try {
      const aiResult = await generateDeckWithOpenAI(brief, themeOptions);
      const slides = adaptAiSlides(aiResult.slides);

      if (slides.length === 0) {
        throw new Error('OpenAI returned an empty outline.');
      }

      const insights = collectInsightsFromAi(aiResult, brief);
      const outline = slides.map((slide) => slide.title || slide.displayTitle || `Slide ${slide.position}`);

      return sendJson(res, 200, {
        provider: 'openai',
        deck: { slides, insights, outline },
        themeKey,
        customThemeName: hasCustomTheme ? customThemeName : null,
        rawModel: aiResult.raw,
        modelUsed: aiResult.modelUsed,
        error: null,
      });
    } catch (error) {
      console.error('OpenAI deck generation failed:', error);
      const fallback = quickDraftDeck(brief);
      return sendJson(res, 200, {
        provider: 'fallback',
        deck: fallback,
        themeKey,
        customThemeName: hasCustomTheme ? customThemeName : null,
        modelUsed: error?.model || null,
        rawModel: { error: { message: error.message, code: error.code, model: error?.model } },
        error: {
          message: error.message || 'Unknown OpenAI error',
          code: error.code || 'UNKNOWN_OPENAI_ERROR',
          model: error?.model,
        },
      });
    }
  } catch (error) {
    console.error('Failed handling /api/generate:', error);
    return sendJson(res, 500, {
      error: {
        message: 'Internal server error.',
        code: 'INTERNAL_SERVER_ERROR',
      },
    });
  }
}

async function handleExport(req, res) {
  try {
    const body = await readBody(req);
    const payload = body ? JSON.parse(body) : {};
    const slides = Array.isArray(payload.slides) ? payload.slides : [];

    if (slides.length === 0) {
      return sendJson(res, 400, {
        error: {
          message: 'Cannot export a deck without slides.',
          code: 'EMPTY_DECK',
        },
      });
    }

    const filenameBase = buildExportFilename(payload.brief || '');
    const exportBuffer = await generatePptxBuffer(payload);
    const stamped = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `${filenameBase}-${stamped}.pptx`;

    res.writeHead(200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Content-Length': exportBuffer.length,
      'Cache-Control': 'no-store',
      'X-Omega-Filename': filename,
    });
    res.end(exportBuffer);
  } catch (error) {
    console.error('Failed to export PPTX:', error);
    const statusCode = error?.code === 'EMPTY_DECK' ? 400 : 500;
    const message = statusCode === 400 ? error.message : 'Unable to generate PPTX file.';
    return sendJson(res, statusCode, {
      error: {
        message,
        code: error?.code || 'EXPORT_FAILURE',
      },
    });
  }
}

async function serveStaticAsset(res, pathname) {
  const requestedPath = pathname === '/' ? '/index.html' : pathname;
  const resolvedPath = path.join(PUBLIC_DIR, path.normalize(requestedPath).replace(/^\.\.\/?/, ''));

  if (!resolvedPath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403);
    return res.end('Forbidden');
  }

  try {
    const file = await fs.readFile(resolvedPath);
    const ext = path.extname(resolvedPath);
    const contentType = CONTENT_TYPES[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(file);
  } catch (error) {
    if (error.code === 'ENOENT') {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not found');
    } else {
      console.error('Static asset error:', error);
      res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Internal server error');
    }
  }
}

const server = http.createServer(async (req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host}`);

  if (req.method === 'POST' && parsedUrl.pathname === '/api/generate') {
    return handleGenerate(req, res);
  }

  if (req.method === 'POST' && parsedUrl.pathname === '/api/export') {
    return handleExport(req, res);
  }

  if (req.method === 'GET') {
    return serveStaticAsset(res, parsedUrl.pathname);
  }

  res.writeHead(405, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end('Method not allowed');
});

server.listen(PORT, () => {
  console.log(`Omega app running on http://localhost:${PORT}`);
});
