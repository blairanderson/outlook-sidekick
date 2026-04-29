/**
 * OpenRouter defaults for Sidekick taskpane integration.
 */

/**
 * @typedef {Object} ProviderConfig
 * @property {string} providerName
 * @property {string} apiKeyEnvVar
 * @property {string} baseUrl
 * @property {string} defaultModel
 * @property {string} replyModel
 * @property {readonly string[]} modelSuggestions
 */

/**
 * @typedef {Object} PromptTemplates
 * @property {string} summarize
 * @property {string} reply
 * @property {string} tldrPrompt
 * @property {string} calendarParse
 * @property {string} calendarCheck
 */

/**
 * @typedef {Object} SettingsDefaults
 * @property {string} model
 * @property {string} replyModel
 * @property {string} theme
 * @property {string} fontSize
 * @property {string} tldrMode
 * @property {string} showReply
 * @property {string} autorun
 * @property {string} autorunOption
 * @property {string} devMode
 * @property {string} devServer
 */

/** @type {readonly string[]} */
export const TEMPLATE_KEYS = Object.freeze([
  "summarize",
  "reply",
  "tldrPrompt",
  "calendarParse",
  "calendarCheck",
]);

/** @type {ProviderConfig} */
export const PROVIDER_CONFIG = Object.freeze({
  providerName: "OpenRouter",
  apiKeyEnvVar: "OPENROUTER_API_KEY",
  baseUrl: "https://openrouter.ai/api/v1",
  defaultModel: "anthropic/claude-3.5-haiku",
  replyModel: "anthropic/claude-3.5-haiku",
  modelSuggestions: Object.freeze([
    "anthropic/claude-3.5-haiku",
    "anthropic/claude-3.5-sonnet",
    "openai/gpt-4o-mini",
  ]),
});

export const PROVIDER_MESSAGES = Object.freeze({
  apiKeyLabel: "OpenRouter API Key",
  modelLabel: "Model",
  missingApiKey: "Please enter your OpenRouter API key in Settings > General and save it.",
  exportTitle: "Sidekick Prompt Templates",
  templatesReset: "Templates reset to defaults",
  templatesCleared: "Prompt templates cleared from Outlook add-in settings.",
  sessionDefaultsSaved: "Current prompt templates saved as Outlook defaults.",
});

/** @type {PromptTemplates} */
export const DEFAULT_PROMPT_TEMPLATES = Object.freeze({
  summarize: `You are Sidekick, an Outlook email assistant.

Review the email below and respond in Markdown with these sections only:
1. TL;DR — 1-3 bullet points with the most important takeaways.
2. Summary — a concise summary of the message.
3. Action Items — concrete asks, decisions, deadlines, or follow-ups.
4. Risks / Open Questions — blockers, ambiguities, or missing information.

Keep the answer factual, compact, and grounded in the email. If the email does not contain enough information for a section, say "None".

Subject: {subject}
Content:
{content}`,
  reply: `You are Sidekick, an Outlook email assistant.

Draft a clear, professional reply to the email below.

Output format:
Subject: <reply subject>

<body>

Requirements:
- Keep the reply actionable and courteous.
- Preserve any important dates, commitments, or questions from the source email.
- If the original email is missing context needed for a confident reply, say so briefly in the body.
- Do not include code fences or commentary outside the reply.

Subject: {subject}
Content:
{content}`,
  tldrPrompt: `Provide a very concise TL;DR for the email below.

Rules:
- Focus on the main point, key actions, and deadlines.
- Keep it short and direct.
- Do not add introductions or commentary.

Subject: {subject}
Content:
{content}`,
  calendarParse: `Analyze the following email content and extract information needed to create a calendar event for Microsoft Graph API.
The response must be valid JSON and follow the required schema exactly.
Return only JSON with no extra commentary.

Event title language instructions:
{languageInstructions}

Required JSON format:
{
  "subject": "Meeting title",
  "body": {
    "contentType": "HTML",
    "content": "Meeting description"
  },
  "start": {
    "dateTime": "YYYY-MM-DDTHH:mm:ss",
    "timeZone": "America/New_York"
  },
  "end": {
    "dateTime": "YYYY-MM-DDTHH:mm:ss",
    "timeZone": "America/New_York"
  },
  "location": {
    "displayName": "Location name"
  },
  "attendees": [
    {
      "emailAddress": {
        "address": "attendee@email.com",
        "name": "Attendee Name"
      },
      "type": "required"
    }
  ],
  "isOnlineMeeting": true,
  "onlineMeetingProvider": "teamsForBusiness"
}

Important notes:
1. Convert dates and times to ISO 8601 format (YYYY-MM-DDTHH:mm:ss)
2. Email addresses must be valid
3. Use "America/New_York" as the default timezone
4. Set isOnlineMeeting to true if Teams or video conference details are present
5. Mark unknown values as null

Email content:
{content}`,
  calendarCheck: `Check whether the following email content is a calendar event, meeting invite, appointment, or schedule-related notice.

Return "true" only if the content clearly includes one or more of these:
- Date and time information
- Meeting or appointment language
- Attendee information
- Location information
- Calendar or RSVP actions

Return "false" otherwise.

Email content:
{content}`,
});

/** @type {SettingsDefaults} */
export const DEFAULT_SETTINGS = Object.freeze({
  model: "",
  replyModel: "",
  theme: "system",
  fontSize: "medium",
  tldrMode: "true",
  showReply: "true",
  autorun: "false",
  autorunOption: "summarize",
  devMode: "false",
  devServer: "",
});

export const BLANK_PROMPT_TEMPLATES = Object.freeze({
  summarize: "",
  reply: "",
  tldrPrompt: "",
  calendarParse: "",
  calendarCheck: "",
});

/**
 * @returns {PromptTemplates}
 */
export function createDefaultPromptTemplates() {
  return {
    summarize: DEFAULT_PROMPT_TEMPLATES.summarize,
    reply: DEFAULT_PROMPT_TEMPLATES.reply,
    tldrPrompt: DEFAULT_PROMPT_TEMPLATES.tldrPrompt,
    calendarParse: DEFAULT_PROMPT_TEMPLATES.calendarParse,
    calendarCheck: DEFAULT_PROMPT_TEMPLATES.calendarCheck,
  };
}

export function createBlankPromptTemplates() {
  return {
    summarize: BLANK_PROMPT_TEMPLATES.summarize,
    reply: BLANK_PROMPT_TEMPLATES.reply,
    tldrPrompt: BLANK_PROMPT_TEMPLATES.tldrPrompt,
    calendarParse: BLANK_PROMPT_TEMPLATES.calendarParse,
    calendarCheck: BLANK_PROMPT_TEMPLATES.calendarCheck,
  };
}

/**
 * @returns {SettingsDefaults & { templates: PromptTemplates }}
 */
export function createDefaultSettings() {
  return {
    ...DEFAULT_SETTINGS,
    templates: createDefaultPromptTemplates(),
  };
}

export function createBlankSettings() {
  return {
    ...DEFAULT_SETTINGS,
    templates: createBlankPromptTemplates(),
  };
}
