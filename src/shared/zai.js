/* global __API_KEY__, __BASE_URL__, AbortController, clearTimeout, fetch, setTimeout */

const DEFAULT_BASE_URL = "https://openrouter.ai/api/v1";
const DEFAULT_MODEL = "anthropic/claude-3.5-haiku";
const DEFAULT_REPLY_MODEL = "anthropic/claude-3.5-haiku";
const MODEL_DISCOVERY_TIMEOUT_MS = 5000;
const CHAT_COMPLETION_TIMEOUT_MS = 120000;
const FALLBACK_MODELS = Object.freeze([
  "anthropic/claude-3.5-haiku",
  "anthropic/claude-3.5-sonnet",
  "anthropic/claude-3-opus",
  "openai/gpt-4o-mini",
  "openai/gpt-4o",
  "google/gemini-2.0-flash-001",
  "google/gemini-1.5-pro",
  "meta-llama/llama-3.3-70b-instruct",
  "mistralai/mistral-7b-instruct",
]);

function normalizeModelName(value) {
  return typeof value === "string" ? value.trim().toLowerCase() : "";
}

function dedupeModels(models) {
  const seen = new Set();
  return models.filter((model) => {
    const normalized = normalizeModelName(model);
    if (!normalized || seen.has(normalized)) {
      return false;
    }

    seen.add(normalized);
    return true;
  });
}

function normalizeTextContent(content) {
  if (typeof content === "string") {
    return content.trim();
  }

  if (Array.isArray(content)) {
    return content
      .map((item) => {
        if (typeof item === "string") {
          return item;
        }

        if (item && typeof item.text === "string") {
          return item.text;
        }

        return "";
      })
      .join("")
      .trim();
  }

  return "";
}

function getBaseUrl() {
  const configuredBaseUrl =
    typeof __BASE_URL__ === "string" && __BASE_URL__.trim()
      ? __BASE_URL__.trim()
      : DEFAULT_BASE_URL;

  return configuredBaseUrl.replace(/\/+$/, "");
}

function getApiKey() {
  return typeof __API_KEY__ === "string" ? __API_KEY__.trim() : "";
}

function hasApiKey() {
  return Boolean(getApiKey());
}

function getDefaultModels() {
  return [...FALLBACK_MODELS];
}

function getDefaultModel() {
  return DEFAULT_MODEL;
}

function getDefaultReplyModel() {
  return DEFAULT_REPLY_MODEL;
}

function createTimeoutSignal(timeoutMs) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

  return {
    signal: controller.signal,
    clear() {
      clearTimeout(timeoutId);
    },
  };
}

function extractModelNames(payload) {
  // OpenRouter returns { data: [{ id: "anthropic/claude-3.5-haiku", ... }] }
  if (Array.isArray(payload?.data)) {
    const discovered = payload.data
      .map((item) => (typeof item?.id === "string" ? item.id.trim() : ""))
      .filter(Boolean);
    return dedupeModels(discovered);
  }

  return [];
}

async function fetchJson(url, options = {}, timeoutMs = MODEL_DISCOVERY_TIMEOUT_MS) {
  const timeout = createTimeoutSignal(timeoutMs);

  try {
    const response = await fetch(url, { ...options, signal: timeout.signal });
    const payload = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(payload?.error?.message || `Request failed (${response.status})`);
    }

    return payload;
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error(`Request timed out after ${Math.ceil(timeoutMs / 1000)} seconds.`);
    }

    throw error;
  } finally {
    timeout.clear();
  }
}

async function generateText(prompt, options = {}) {
  const apiKey = options.apiKey || getApiKey();
  const model = normalizeModelName(options.model) || DEFAULT_MODEL;

  if (!apiKey) {
    throw new Error("API key is not configured.");
  }

  const payload = await fetchJson(
    `${getBaseUrl()}/chat/completions`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/AlanSynn/michael",
        "X-Title": "Sidekick for Outlook",
      },
      body: JSON.stringify({
        model,
        messages: [
          {
            role: "user",
            content: prompt,
          },
        ],
        temperature: options.temperature ?? 0.4,
        max_tokens: options.maxTokens ?? 4096,
        stream: false,
      }),
    },
    options.timeoutMs ?? CHAT_COMPLETION_TIMEOUT_MS
  );

  const content = normalizeTextContent(payload?.choices?.[0]?.message?.content);
  if (!content) {
    throw new Error("No content generated.");
  }

  return content;
}

async function fetchAvailableModels(options = {}) {
  const apiKey = options.apiKey || getApiKey();
  if (!apiKey) {
    throw new Error("API key is not configured.");
  }

  const payload = await fetchJson(
    `${getBaseUrl()}/models`,
    {
      method: "GET",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/AlanSynn/michael",
        "X-Title": "Sidekick for Outlook",
      },
    },
    options.timeoutMs
  );

  const discoveredModels = extractModelNames(payload);
  if (!discoveredModels.length) {
    throw new Error("No models were returned by the provider.");
  }

  return discoveredModels;
}

export {
  FALLBACK_MODELS,
  fetchAvailableModels,
  generateText,
  getDefaultModel,
  getDefaultModels,
  getDefaultReplyModel,
  getApiKey,
  getBaseUrl,
  hasApiKey,
};
