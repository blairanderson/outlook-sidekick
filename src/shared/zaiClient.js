/* global fetch, module, require */

const { CHAT_COMPLETIONS_URL, DEFAULT_MODEL, requireApiKey } = require("./zaiConfig");

function buildChatMessages(systemPrompt, userPrompt) {
  const messages = [];

  if (systemPrompt) {
    messages.push({ role: "system", content: systemPrompt });
  }

  messages.push({ role: "user", content: userPrompt });

  return messages;
}

function buildChatCompletionRequest({
  systemPrompt = "",
  userPrompt,
  model = DEFAULT_MODEL,
  temperature = 0.3,
}) {
  if (!userPrompt || !userPrompt.trim()) {
    throw new Error("Requests require a non-empty user prompt.");
  }

  return {
    model,
    messages: buildChatMessages(systemPrompt, userPrompt.trim()),
    temperature,
    stream: false,
  };
}

async function parseJsonResponse(response) {
  const body = await response.text();

  if (!body) {
    return {};
  }

  try {
    return JSON.parse(body);
  } catch {
    throw new Error(`Non-JSON response received (${response.status}).`);
  }
}

function extractMessageText(data) {
  const content = data?.choices?.[0]?.message?.content;

  if (typeof content !== "string" || !content.trim()) {
    throw new Error("No message content returned.");
  }

  return content.trim();
}

async function executeChatCompletion(request) {
  const apiKey =
    typeof request?.apiKey === "string" && request.apiKey.trim()
      ? request.apiKey.trim()
      : requireApiKey();
  const payload = buildChatCompletionRequest(request);
  const response = await fetch(CHAT_COMPLETIONS_URL, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
      "HTTP-Referer": "https://github.com/AlanSynn/michael",
      "X-Title": "Sidekick for Outlook",
    },
    body: JSON.stringify(payload),
  });
  const data = await parseJsonResponse(response);
  const errorMessage = data?.error?.message || data?.message;

  if (!response.ok || errorMessage) {
    throw new Error(errorMessage || `Request failed (${response.status}).`);
  }

  return {
    text: extractMessageText(data),
    data,
  };
}

module.exports = {
  buildChatCompletionRequest,
  executeChatCompletion,
};
