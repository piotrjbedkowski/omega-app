const form = document.getElementById('brief-form');
const textarea = document.getElementById('brief');
const preview = document.getElementById('deck-preview');
const previewMessage = document.getElementById('preview-message');
const previewMessageText = previewMessage?.querySelector('.empty-state__text') || null;
const previewProgress = previewMessage?.querySelector('.empty-state__progress') || null;
const previewProgressBar = previewProgress?.querySelector('.empty-state__progress-bar') || null;
const successBanner = document.getElementById('generation-success');
const successBannerText = successBanner?.querySelector('.success-banner__text') || null;
const deckActions = document.getElementById('deck-actions');
const slidesLink = document.getElementById('google-slides-link');
const downloadButton = document.getElementById('download-deck');
const status = document.getElementById('output-status');
const includeExampleBtn = document.getElementById('include-example') || null;
const themeUploadInput = document.getElementById('theme-upload') || null;
const submitButton = form.querySelector('button[type="submit"]');

const API_ENDPOINT = '/api/generate';
const EXAMPLE_THEME_KEY = 'example-pptx';

const submitDefaultLabel = submitButton.textContent;

const deckState = {
  brief: '',
  slides: [],
  insights: [],
  outline: [],
  includeExampleTheme: false,
  themeKey: 'default',
  generatedAt: null,
  provider: 'local',
  modelUsed: null,
  rawModelResponse: null,
  lastError: null,
  customThemeName: null,
  customThemeData: null,
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

let useExampleTheme = false;
let uploadedTheme = { name: null, data: null };

let downloadButtonDefaultLabel = downloadButton ? downloadButton.textContent : '';
let slidesLinkDefaultLabel = slidesLink ? slidesLink.textContent.trim() : '';
if (downloadButton) {
  downloadButton.disabled = true;
}

const DEFAULT_PROGRESS_DURATION_MS = 15000;
let progressIntervalId = null;
let progressDeadline = 0;
let progressBaseMessage = '';
let progressStartTime = 0;

let progressAnimationId = null;

function stopProgressCountdown() {
  if (progressIntervalId) {
    window.clearInterval(progressIntervalId);
    progressIntervalId = null;
  }
  if (progressAnimationId) {
    window.cancelAnimationFrame(progressAnimationId);
    progressAnimationId = null;
  }
  progressDeadline = 0;
  progressBaseMessage = '';
  progressStartTime = 0;
  if (previewProgressBar) {
    previewProgressBar.style.transform = 'scaleX(0)';
  }
  if (previewMessageText) {
    previewMessageText.textContent = '';
  } else if (previewMessage) {
    previewMessage.textContent = '';
  }
}

function hideSuccessBanner() {
  if (successBanner) {
    successBanner.classList.add('success-banner--hidden');
  }
}

function showSuccessBanner(message) {
  if (!successBanner) {
    return;
  }
  if (successBannerText) {
    successBannerText.textContent = message;
  }
  successBanner.classList.remove('success-banner--hidden');
}

function updateProgressMessage() {
  if (!previewMessageText && !previewMessage) {
    return;
  }
  const now = Date.now();
  const remainingMs = Math.max(0, progressDeadline - now);
  const remainingSeconds = Math.max(0, Math.ceil(remainingMs / 1000));
  const base = progressBaseMessage || 'Working…';
  const message = `${base} (~${remainingSeconds}s remaining)`;
  if (previewMessageText) {
    previewMessageText.textContent = message;
  } else if (previewMessage) {
    previewMessage.textContent = message;
  }

  if (previewProgressBar && progressStartTime) {
    const elapsedMs = Math.min(Date.now() - progressStartTime, progressDeadline - progressStartTime);
    const totalMs = Math.max(progressDeadline - progressStartTime, 1);
    const ratio = Math.min(Math.max(elapsedMs / totalMs, 0), 1);
    previewProgressBar.style.transform = `scaleX(${ratio})`;
  }
}

function startProgressCountdown(baseMessage, durationMs = DEFAULT_PROGRESS_DURATION_MS) {
  stopProgressCountdown();
  progressBaseMessage = baseMessage;
  progressStartTime = Date.now();
  progressDeadline = progressStartTime + durationMs;
  if (previewProgressBar) {
    previewProgressBar.style.transform = 'scaleX(0)';
  }
  updateProgressMessage();
  progressIntervalId = window.setInterval(updateProgressMessage, 1000);
  const animate = () => {
    if (!progressStartTime || !progressDeadline) {
      progressAnimationId = null;
      return;
    }
    const now = Date.now();
    const elapsedMs = Math.min(now - progressStartTime, progressDeadline - progressStartTime);
    const totalMs = Math.max(progressDeadline - progressStartTime, 1);
    const ratio = Math.min(Math.max(elapsedMs / totalMs, 0), 1);
    if (previewProgressBar) {
      previewProgressBar.style.transform = `scaleX(${ratio})`;
    }
    progressAnimationId = window.requestAnimationFrame(animate);
  };
  progressAnimationId = window.requestAnimationFrame(animate);
}

function cloneSlides(slides) {
  return slides.map((slide) => ({
    ...slide,
    bullets: [...slide.bullets],
  }));
}

function showOutputSection() {
  const outputSection = document.querySelector('section.output');
  if (outputSection && outputSection.classList.contains('output--hidden')) {
    outputSection.classList.remove('output--hidden');
  }
}

function persistDeckState({ brief, slides, insights, outline, provider = 'local', modelUsed = null, rawModel = null, error = null, themeKey, customThemeName = null, customThemeData = null }) {
  deckState.brief = brief;
  deckState.slides = slides;
  deckState.insights = insights;
  deckState.outline = outline;
  deckState.includeExampleTheme = themeKey === EXAMPLE_THEME_KEY || themeKey === 'custom-upload';
  deckState.themeKey = themeKey || (useExampleTheme ? EXAMPLE_THEME_KEY : 'default');
  deckState.generatedAt = new Date().toISOString();
  deckState.provider = provider;
  deckState.modelUsed = modelUsed;
  deckState.rawModelResponse = rawModel;
  deckState.lastError = error;
  deckState.customThemeName = customThemeName;
  deckState.customThemeData = customThemeData;
}

function getDeckState() {
  return {
    ...deckState,
    slides: cloneSlides(deckState.slides),
    insights: [...deckState.insights],
    outline: [...deckState.outline],
    rawModelResponse: deckState.rawModelResponse,
  };
}

function buildExportPayload() {
  if (deckState.slides.length === 0) {
    throw new Error('Cannot build export payload before generating a deck outline.');
  }

  return {
    brief: deckState.brief,
    outline: [...deckState.outline],
    slides: cloneSlides(deckState.slides),
    speakerNotes: [...deckState.insights],
    options: {
      includeExampleTheme: deckState.includeExampleTheme,
      themeKey: deckState.themeKey,
      customThemeName: deckState.customThemeName,
    },
    meta: {
      generatedAt: deckState.generatedAt,
      provider: deckState.provider,
      modelUsed: deckState.modelUsed,
      lastError: deckState.lastError,
    },
    rawModelResponse: deckState.rawModelResponse,
    customTheme: deckState.customThemeData,
  };
}

// Surface helpers the export workflow can consume once wired up.
window.omegaDeck = window.omegaDeck || {};
window.omegaDeck.getState = getDeckState;
window.omegaDeck.buildExportPayload = buildExportPayload;

async function requestDeckExport(payload) {
  const response = await fetch('/api/export', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const contentType = response.headers.get('Content-Type') || '';
    let message = `${response.status} ${response.statusText}`;
    if (contentType.includes('application/json')) {
      const data = await response.json().catch(() => null);
      message = data?.error?.message || message;
      const error = new Error(message);
      error.code = data?.error?.code;
      throw error;
    }
    const text = await response.text().catch(() => '');
    if (text) {
      message = text;
    }
    const error = new Error(message);
    error.code = 'EXPORT_FAILED';
    throw error;
  }

  const blob = await response.blob();
  const filename = response.headers.get('X-Omega-Filename')
    || `omega-deck-${new Date().toISOString().replace(/[:.]/g, '-')}.pptx`;

  return { blob, filename };
}

if (downloadButton) {
  downloadButton.addEventListener('click', async () => {
    try {
      const payload = buildExportPayload();
      downloadButton.disabled = true;
      downloadButton.textContent = 'Preparing PPTX…';
      status.textContent = 'Preparing your Google Slides export…';

      const { blob, filename } = await requestDeckExport(payload);
      triggerBlobDownload(blob, filename);

      status.textContent = `Downloaded ${filename}. You can import it into Google Slides.`;
    } catch (error) {
      console.error('Failed to download deck:', error);
      status.textContent = error?.message || 'Unable to download deck. Please try again.';
    } finally {
      downloadButton.disabled = false;
      downloadButton.textContent = downloadButtonDefaultLabel;
    }
  });
}

function triggerBlobDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

async function openInGoogleSlides(event) {
  if (event) {
    event.preventDefault();
  }

  if (!slidesLink || slidesLink.getAttribute('aria-disabled') === 'true') {
    return;
  }

  try {
    const payload = buildExportPayload();
    slidesLink.setAttribute('aria-disabled', 'true');
    slidesLink.textContent = 'Opening in Google Slides…';
    status.textContent = 'Preparing Google Slides import…';

    const { blob, filename } = await requestDeckExport(payload);
    triggerBlobDownload(blob, filename);
    window.open('https://docs.google.com/presentation/u/0/import', '_blank', 'noopener');
    status.textContent = 'Downloaded the deck and opened Google Slides import in a new tab. Upload the file you just downloaded to finish.';
  } catch (error) {
    console.error('Failed to open deck in Google Slides:', error);
    status.textContent = error?.message || 'Unable to open in Google Slides. Please download the deck instead.';
  } finally {
    if (slidesLink) {
      slidesLink.setAttribute('aria-disabled', Array.isArray(deckState.slides) && deckState.slides.length > 0 ? 'false' : 'true');
      slidesLink.textContent = slidesLinkDefaultLabel;
    }
  }
}

if (slidesLink) {
  slidesLink.addEventListener('click', openInGoogleSlides);
}

function setGenerating(isGenerating) {
  submitButton.disabled = isGenerating;
  if (includeExampleBtn) {
    includeExampleBtn.disabled = isGenerating;
  }
  textarea.readOnly = isGenerating;
  form.classList.toggle('is-generating', isGenerating);
  if (!isGenerating) {
    submitButton.textContent = submitDefaultLabel;
    if (includeExampleBtn) {
      includeExampleBtn.disabled = false;
    }
    textarea.readOnly = false;
    if (previewMessageText) {
      previewMessageText.textContent = '';
    } else if (previewMessage) {
      previewMessage.textContent = '';
    }
  }
}

function setThemeUpload(theme) {
  uploadedTheme = theme || { name: null, data: null };
  useExampleTheme = Boolean(uploadedTheme?.data);
  if (includeExampleBtn) {
    updateThemeUi({ announce: true });
  }
}

function updateThemeUi({ announce = false } = {}) {
  if (!includeExampleBtn) {
    return;
  }
  const hasTheme = Boolean(uploadedTheme?.data);
  const label = hasTheme
    ? `Theme uploaded (${uploadedTheme.name})`
    : 'Upload custom pptx theme';

  includeExampleBtn.setAttribute('aria-pressed', String(hasTheme));
  includeExampleBtn.textContent = label;

  if (hasTheme) {
    preview.dataset.theme = 'example';
  } else {
    delete preview.dataset.theme;
  }

  if (announce) {
    status.textContent = hasTheme
      ? `Custom theme "${uploadedTheme.name}" uploaded. It will be used for the next deck.`
      : 'Custom theme cleared.';
  }
}

async function handleThemeUpload(event) {
  const file = event.target.files?.[0];

  if (!file) {
    return;
  }

  if (!file.name.toLowerCase().endsWith('.pptx')) {
    status.textContent = 'Please choose a .pptx file to use as your theme.';
    event.target.value = '';
    return;
  }

  try {
    const base64 = await readFileAsBase64(file);
    setThemeUpload({ name: file.name, data: base64 });
    event.target.value = '';
  } catch (error) {
    console.error('Failed to read theme file:', error);
    status.textContent = 'We could not read that file. Please try again with a different pptx.';
    event.target.value = '';
  }
}

function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result;
      if (typeof result === 'string') {
        const [, base64 = ''] = result.split(',', 2);
        resolve(base64);
      } else if (result instanceof ArrayBuffer) {
        const bytes = new Uint8Array(result);
        let binary = '';
        for (let i = 0; i < bytes.byteLength; i += 1) {
          binary += String.fromCharCode(bytes[i]);
        }
        resolve(btoa(binary));
      } else {
        reject(new Error('Unsupported file reader result'));
      }
    };
    reader.onerror = () => reject(reader.error || new Error('Failed reading file.'));
    reader.readAsDataURL(file);
  });
}

if (includeExampleBtn) {
  includeExampleBtn.addEventListener('click', () => {
    if (themeUploadInput) {
      themeUploadInput.click();
    }
  });
}

if (themeUploadInput) {
  themeUploadInput.addEventListener('change', handleThemeUpload);
}

function synthesizeBullets(text) {
  return text
    .split(/[\n\.]/)
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, 4);
}

function generateSlides(brief, outlineHints) {
  const keywords = brief
    .split(/[,\n]/)
    .map((part) => part.trim())
    .filter(Boolean);

  const slides = [];
  for (let i = 0; i < outlineHints.length; i += 1) {
    const title = outlineHints[i];
    const bullets = [];

    const start = i * 2;
    const segment = keywords.slice(start, start + 2);
    if (segment.length === 0) {
      bullets.push('Highlight the most important takeaways for this section.');
    } else {
      segment.forEach((keyword) => {
        bullets.push(`Focus on ${keyword.toLowerCase()}.`);
      });
    }

    slides.push({
      id: `slide-${i + 1}`,
      title,
      position: i + 1,
      displayTitle: `${i + 1}. ${title}`,
      bullets,
    });
  }

  return slides;
}

function quickDraftDeck(brief) {
  const insights = synthesizeBullets(brief);
  const slideCount = Math.max(6, Math.min(10, insights.length + 4));
  const outline = slideTitles.slice(0, slideCount);
  const slides = generateSlides(brief, outline);

  return { slides, insights, outline };
}

function setPreviewMessage(message, { showProgress = false } = {}) {
  if (!previewMessage) {
    return;
  }

  if (showProgress) {
    if (previewProgress) {
      previewProgress.hidden = false;
    }
    startProgressCountdown(message);
  } else {
    stopProgressCountdown();
    if (previewProgress) {
      previewProgress.hidden = true;
    }
    if (previewMessageText) {
      previewMessageText.textContent = message;
    } else {
      previewMessage.textContent = message;
    }
  }

  previewMessage.hidden = false;
  hideSuccessBanner();
  showOutputSection();

  if (deckActions) {
    deckActions.hidden = true;
  }

  if (downloadButton) {
    downloadButton.disabled = true;
    downloadButton.textContent = downloadButtonDefaultLabel;
  }

  if (slidesLink) {
    slidesLink.setAttribute('aria-disabled', 'true');
    slidesLink.textContent = slidesLinkDefaultLabel;
  }
}

function showDeckActions() {
  stopProgressCountdown();
  if (previewMessage) {
    previewMessage.hidden = true;
  }
  if (deckActions) {
    deckActions.hidden = false;
  }
  if (downloadButton) {
    downloadButton.disabled = false;
    downloadButton.textContent = downloadButtonDefaultLabel;
  }
  if (previewProgress) {
    previewProgress.hidden = true;
  }
  if (slidesLink) {
    const hasSlides = Array.isArray(deckState.slides) && deckState.slides.length > 0;
    slidesLink.setAttribute('aria-disabled', hasSlides ? 'false' : 'true');
    slidesLink.textContent = slidesLinkDefaultLabel;
  }
}

function renderDeck(slidesInput) {
  const hasSlides = Array.isArray(slidesInput) && slidesInput.length > 0;

  if (!hasSlides) {
    setPreviewMessage('No slides were generated. Please try again.');
    return;
  }
}

async function requestDeckFromServer(brief, { includeExampleTheme = false, customTheme = null, customThemeName = null } = {}) {
  const payload = {
    brief,
    includeExampleTheme,
    customTheme,
    customThemeName,
  };

  const response = await fetch(API_ENDPOINT, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const text = await response.text();
    const message = text ? `${response.status} ${response.statusText}: ${text}` : `${response.status} ${response.statusText}`;
    const error = new Error(`Server responded with ${message}`);
    error.code = 'SERVER_ERROR';
    throw error;
  }

  return response.json();
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  const brief = textarea.value.trim();

  if (!brief) {
    status.textContent = 'Please describe the deck before generating.';
    textarea.focus();
    return;
  }

  setGenerating(true);
  showOutputSection();
  setPreviewMessage('Drafting your deck…', { showProgress: true });

  let deckResult = { slides: [], insights: [], outline: [] };
  let provider = 'openai';
  let rawModel = null;
  let capturedError = null;
  let themeKey = useExampleTheme ? EXAMPLE_THEME_KEY : 'default';
  let modelUsed = null;

  try {
    const apiResponse = await requestDeckFromServer(brief, {
      includeExampleTheme: useExampleTheme,
      customTheme: uploadedTheme?.data || null,
      customThemeName: uploadedTheme?.name || null,
    });
    provider = apiResponse.provider || 'openai';
    deckResult = apiResponse.deck || deckResult;
    rawModel = apiResponse.rawModel || null;
    capturedError = apiResponse.error || null;
    themeKey = apiResponse.themeKey || themeKey;
    modelUsed = apiResponse.modelUsed || null;

    if (provider === 'openai') {
      status.textContent = 'OpenAI responded. Building your preview…';
    } else if (provider === 'fallback') {
      if (capturedError?.code === 'MISSING_OPENAI_API_KEY') {
        status.textContent = 'OpenAI API key missing on server. Generated a quick draft locally instead.';
      } else {
        const message = capturedError?.message || 'OpenAI request failed. Using local draft.';
        status.textContent = `${message} Using local draft instead.`;
      }
    }
  } catch (error) {
    console.error('Deck generation request failed:', error);
    provider = 'fallback-client';
    capturedError = {
      message: error?.message || 'Unknown client error',
      code: error?.code || 'CLIENT_FETCH_ERROR',
    };
    deckResult = quickDraftDeck(brief);
    rawModel = { error: capturedError };
    status.textContent = 'Server request failed. Generated a quick draft locally instead.';
  }

  try {
    renderDeck(deckResult.slides);

    persistDeckState({
      brief,
      slides: deckResult.slides,
      insights: deckResult.insights,
      outline: deckResult.outline,
      provider,
      modelUsed,
      rawModel,
      error: capturedError,
      themeKey,
      customThemeName: uploadedTheme?.name || null,
      customThemeData: uploadedTheme?.data || null,
    });

    // Notify any listeners (e.g., export module) that fresh deck data exists.
    const deckGeneratedDetail = {
      state: getDeckState(),
      payload: buildExportPayload(),
    };
    document.dispatchEvent(new CustomEvent('omega:deckGenerated', { detail: deckGeneratedDetail }));

    let finalMessage = 'Your deck outline is ready. Use the actions below to open Google Slides or download the file.';
    if (useExampleTheme && uploadedTheme?.name) {
      finalMessage += ` Custom theme "${uploadedTheme.name}" is applied to this outline and will be used for export.`;
    }
    if (provider === 'openai') {
      const modelDescriptor = modelUsed ? ` (${modelUsed})` : '';
      finalMessage += ` Outline powered by OpenAI${modelDescriptor}.`;
    } else if (provider === 'fallback') {
      finalMessage += ' OpenAI request failed; this outline came from the server fallback.';
    } else {
      finalMessage += ' Server unreachable; this outline was generated locally in your browser.';
    }
    status.textContent = finalMessage;
    showSuccessBanner('🎉 Congrats! Your deck outline is ready.');
    if (Array.isArray(deckResult.slides) && deckResult.slides.length > 0) {
      showDeckActions();
    }
  } catch (renderError) {
    console.error('Unexpected error while rendering deck:', renderError);
    status.textContent = 'An unexpected error occurred while preparing your deck. Please try again.';
    setPreviewMessage('We hit an error while preparing your deck. Please try again.');
  } finally {
    setGenerating(false);
  }
});

// Ensure the preview reflects the default button state on load.
updateThemeUi();
