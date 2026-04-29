# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

"Michael" is an Outlook add-in that uses Z.AI GLM models to summarize, translate, draft replies, and create calendar events from emails. It is a vanilla JavaScript + Webpack project built on the Office Add-ins platform.

## Commands

```bash
npm start          # Start dev server + sideload add-in (requires Outlook desktop)
npm run build      # Production build → dist/
npm run build:dev  # Development build
npm run watch      # Webpack watch mode
npm run lint       # Lint check
npm run lint:fix   # Auto-fix lint issues
npm run validate   # Validate manifest.xml against Office schema
npm run stop       # Stop the debugging session
```

Benchmark Z.AI models (requires `ZAI_API_KEY` env var):
```bash
ZAI_API_KEY=... node scripts/benchmark-zai-models.mjs
```

Environment variables (passed into the bundle via `webpack.DefinePlugin`):
- `ZAI_API_KEY` — bakes a fallback API key into the build; runtime always prefers the key saved in Outlook roaming settings
- `ZAI_CODING_BASE_URL` — overrides the default `https://api.z.ai/api/coding/paas/v4`

The dev server runs on `https://localhost:3000` with self-signed certs managed by `office-addin-dev-certs`.

## Architecture

### Two webpack entry points

| Entry | File | Module system | Purpose |
|---|---|---|---|
| `taskpane` | `src/taskpane/taskpane.js` | ES modules (`import`/`export`) | Main sidebar UI |
| `commands` | `src/commands/commands.js` | CommonJS (`require`) | Ribbon button (quick-translate command) |

Both entries share `src/shared/` helpers but use different import styles — keep them consistent with their respective files.

### Shared modules (`src/shared/`)

- **`zai.js`** — ES-module API client used by `taskpane.js`. Exports `generateText`, `fetchAvailableModels`, `getDefaultZaiModels`. Handles timeout via `AbortController`, normalises model names, and extracts text from `choices[0].message.content`.
- **`zaiConfig.js`** — CommonJS constants (`ZAI_CODING_BASE_URL`, `ZAI_DEFAULT_MODEL`) and `requireZaiApiKey()` used by `commands.js`.
- **`zaiClient.js`** — CommonJS thin wrapper over `fetch` used by `commands.js` for chat completions.

### Settings persistence

All settings are stored in `Office.context.roamingSettings` (roams across devices via Exchange). The key `michael_settings` holds a JSON blob with the full settings object. The key `michael_template_defaults` holds saved prompt defaults. `michael_zai_model_catalog` caches the live model list.

On load, `migrateSettingsKeys()` promotes any legacy `sessionStorage`/`localStorage` keys into roaming settings, then cleans them up.

Settings shape (defined in `src/taskpane/prompt-templates.js`):
- `apiKey`, `model`, `replyModel` — Z.AI credentials and model selection
- `defaultLanguage`, `eventTitleLanguage` — language codes (`ko`, `en`, `ja`, `zh_cn`, etc.)
- `theme` (`light` | `dark` | `system`), `fontSize`, `tldrMode`, `showReply`, `autorun`, `autorunOption`, `devMode`, `devServer`
- `templates` — object with keys: `summarize`, `translate`, `translateSummarize`, `reply`, `commandTranslate`, `tldrPrompt`, `calendarParse`, `calendarCheck`

### Prompt templates

All built-in defaults live in `src/taskpane/prompt-templates.js` (`DEFAULT_PROMPT_TEMPLATES`). Templates use `{subject}`, `{content}`, `{language}`, `{languageInstructions}` as placeholders replaced at call time. When a template is empty in settings, `requireTemplate()` throws an error directing the user to Settings > Templates.

### TL;DR mode

When `tldrMode` is enabled (default), the add-in fires two sequential Z.AI calls: a fast short-token TL;DR first (shown immediately), then the full-length generation in the background. The "Show Full Content" button is enabled once the second call completes.

### Webpack production build

In production mode, `webpack.config.js` copies `manifest.xml` and replaces `https://localhost:3000/` with `https://alansynn.com/michael/` and `"Michael [Local]"` with `"Michael"` in the output manifest.

### Z.AI API

- Endpoint: `https://api.z.ai/api/coding/paas/v4/chat/completions`
- Auth: `Authorization: Bearer <key>`
- Default model: `glm-4.5-air`; reply model: `glm-4.5-air`
- Chat completion timeout: 120 s; model discovery timeout: 5 s
- `thinking` is always disabled (`{ type: "disabled" }`)
