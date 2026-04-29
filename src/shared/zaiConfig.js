/* global __API_KEY__, module */

const BASE_URL = "https://openrouter.ai/api/v1";
const CHAT_COMPLETIONS_URL = `${BASE_URL}/chat/completions`;
const DEFAULT_MODEL = "anthropic/claude-3.5-haiku";

function getApiKey() {
  if (typeof __API_KEY__ !== "string") {
    return "";
  }

  return __API_KEY__.trim();
}

function hasApiKeyConfigured() {
  return getApiKey().length > 0;
}

function requireApiKey() {
  const apiKey = getApiKey();

  if (!apiKey) {
    throw new Error(
      "Missing OPENROUTER_API_KEY. Set the environment variable before starting the add-in."
    );
  }

  return apiKey;
}

module.exports = {
  BASE_URL,
  CHAT_COMPLETIONS_URL,
  DEFAULT_MODEL,
  getApiKey,
  hasApiKeyConfigured,
  requireApiKey,
};
