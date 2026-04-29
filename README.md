# Sidekick for Outlook

AI email assistant powered by [OpenRouter](https://openrouter.ai). Summarize emails, draft replies, and create calendar events directly inside Outlook.

## Features

- **Summarize** — TL;DR + full summary with action items and open questions
- **Generate Reply** — draft a professional reply ready to copy and send
- **Calendar Event** — detect meeting invites and pre-fill a new Outlook appointment
- **Copy as…** — copy any result as Markdown, Plaintext, or HTML
- **TL;DR mode** — shows a quick preview immediately while the full response loads in the background
- **Customizable prompts** — edit every prompt template in Settings > Templates

## Installation

### 1. Get an OpenRouter API key

Sign up at [openrouter.ai/keys](https://openrouter.ai/keys) and create an API key.

### 2. Sideload the add-in

Download [`manifest.prod.xml`](https://blairanderson.github.io/outlook-sidekick/manifest.prod.xml) and sideload it into Outlook:

- **Outlook on the web** — Settings → Add-ins → My add-ins → Add a custom add-in → Add from URL or file
- **Outlook desktop (Mac/Windows)** — Home ribbon → Get Add-ins → My add-ins → Add a custom add-in → Add from file

Microsoft docs: [Sideload Outlook add-ins for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing)

### 3. Enter your API key

Open the Sidekick pane in Outlook → click **Settings** → paste your OpenRouter API key → **Save Settings**.

### 4. Select a model

In Settings > General, choose a model from the dropdown (populated from your OpenRouter account). Recommended starting points:

| Use case | Model |
|---|---|
| Fast + cheap | `anthropic/claude-3.5-haiku` |
| Best quality | `anthropic/claude-3.5-sonnet` |
| Budget option | `openai/gpt-4o-mini` |

## Local Development

```bash
git clone https://github.com/blairanderson/outlook-sidekick.git
cd outlook-sidekick
bun install
bun run start        # starts dev server + sideloads into Outlook desktop
```

Requires [Bun](https://bun.sh) and Outlook desktop. The dev server runs at `https://localhost:3000` with auto-generated HTTPS certs.

Other commands:

```bash
bun run build        # production build → dist/
bun run build:dev    # development build
bun run watch        # webpack watch mode
bun run validate     # validate manifest.xml
bun run lint         # lint check
```

### Environment variables

| Variable | Description |
|---|---|
| `OPENROUTER_API_KEY` | Bakes a fallback key into the bundle at build time (optional — users can enter their key in Settings instead) |
| `OPENROUTER_BASE_URL` | Override the API base URL (default: `https://openrouter.ai/api/v1`) |

## Tech stack

- Office Add-ins Platform (Outlook)
- Vanilla JavaScript (ES6+) + HTML/CSS
- Webpack 5 + Babel
- Bun (package manager + script runner)
- OpenRouter API (OpenAI-compatible chat completions)

## License

MIT
