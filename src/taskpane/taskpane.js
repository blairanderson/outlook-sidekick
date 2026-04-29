/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Blob, console, document, localStorage, marked, navigator, Office, sessionStorage, setTimeout, URL, window */

import { fetchAvailableModels, generateText, getDefaultModels } from "../shared/zai";
import {
  createBlankPromptTemplates,
  createBlankSettings,
  createDefaultPromptTemplates,
  DEFAULT_SETTINGS,
  PROVIDER_MESSAGES,
  TEMPLATE_KEYS,
} from "./prompt-templates";

const TYPES = {
  SUMMARIZE: 0,
  REPLY: 1,
  CALENDAR: 2,
};

const MODEL_CATALOG_CACHE_KEY = "michael_model_catalog";
const TEMPLATE_DEFAULTS_KEY = "michael_template_defaults";
const SETTINGS_KEY = "michael_settings";
let availableModelCatalog = getDefaultModels();

// Stores the raw markdown string for the last result so "Copy as…" can use it
let lastRawContent = "";

const PROMPT_FIELD_MAP = Object.freeze({
  summarize: "dropdown-summarize-template",
  reply: "dropdown-reply-template",
  tldrPrompt: "dropdown-tldr-template",
  calendarParse: "dropdown-calendar-parse-template",
  calendarCheck: "dropdown-calendar-check-template",
});

// ---------------------------------------------------------------------------
// Roaming settings helpers
// ---------------------------------------------------------------------------

function getRoamingStore() {
  return Office?.context?.roamingSettings || null;
}

function saveRoamingStoreAsync() {
  const store = getRoamingStore();
  if (!store) {
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    store.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error?.message || "Failed to save Outlook add-in settings."));
        return;
      }

      resolve();
    });
  });
}

const roamingStorage = Object.freeze({
  getItem(key) {
    try {
      const store = getRoamingStore();
      const value = store ? store.get(key) : null;
      return typeof value === "string" ? value : null;
    } catch (error) {
      console.error(`Error reading Outlook add-in setting ${key}:`, error);
      return null;
    }
  },
  setItem(key, value) {
    try {
      const store = getRoamingStore();
      if (!store) {
        return;
      }

      store.set(key, value);
    } catch (error) {
      console.error(`Error writing Outlook add-in setting ${key}:`, error);
    }
  },
  removeItem(key) {
    try {
      const store = getRoamingStore();
      if (!store) {
        return;
      }

      store.remove(key);
    } catch (error) {
      console.error(`Error removing Outlook add-in setting ${key}:`, error);
    }
  },
});

/**
 * Migrate legacy browser storage keys into Outlook add-in settings.
 */
function migrateSettingsKeys() {
  try {
    const legacySettings = sessionStorage.getItem("my_sidekick_michael_settings");
    const currentSettings = roamingStorage.getItem(SETTINGS_KEY);
    if (!currentSettings && legacySettings) {
      roamingStorage.setItem(SETTINGS_KEY, legacySettings);
    }
    sessionStorage.removeItem("my_sidekick_michael_settings");

    const previousLocalSettings = localStorage.getItem("michael_settings");
    if (!currentSettings && previousLocalSettings) {
      roamingStorage.setItem(SETTINGS_KEY, previousLocalSettings);
    }

    const legacyTemplateDefaults = sessionStorage.getItem("michael_session_template_defaults");
    const currentTemplateDefaults = roamingStorage.getItem(TEMPLATE_DEFAULTS_KEY);
    if (!currentTemplateDefaults && legacyTemplateDefaults) {
      roamingStorage.setItem(TEMPLATE_DEFAULTS_KEY, legacyTemplateDefaults);
    }

    const legacyModelCatalog = sessionStorage.getItem(MODEL_CATALOG_CACHE_KEY);
    const currentModelCatalog = roamingStorage.getItem(MODEL_CATALOG_CACHE_KEY);
    if (!currentModelCatalog && legacyModelCatalog) {
      roamingStorage.setItem(MODEL_CATALOG_CACHE_KEY, legacyModelCatalog);
    }

    saveRoamingStoreAsync().catch((error) => {
      console.error("Error saving migrated Outlook add-in settings:", error);
    });
  } catch {
    // no-op
  }
}

// ---------------------------------------------------------------------------
// Settings view
// ---------------------------------------------------------------------------

function toggleSettingsView() {
  const settingsView = document.getElementById("settings-view");
  const appBody = document.getElementById("app-body");

  if (settingsView && appBody) {
    const isSettingsVisible = settingsView.style.display === "block";
    if (isSettingsVisible) {
      settingsView.style.display = "none";
      appBody.style.display = "flex";
    } else {
      settingsView.style.display = "block";
      appBody.style.display = "none";
      loadDropdownSettings();
      refreshModelCatalog({ silent: true });
    }
  }
}

function initializeSettingsTabs() {
  const tabButtons = document.querySelectorAll(".settings-tabs .settings-tab-button");
  const tabContents = document.querySelectorAll(".settings-content .settings-tab-content");

  if (!tabButtons.length || !tabContents.length) return;

  tabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const targetTabId = button.getAttribute("data-tab");
      tabButtons.forEach((btn) => btn.classList.remove("active"));
      tabContents.forEach((content) => content.classList.remove("active"));
      button.classList.add("active");
      const targetContent = document.getElementById(targetTabId);
      if (targetContent) {
        targetContent.classList.add("active");
      }
    });
  });

  const firstTabButton = tabButtons[0];
  const firstTabContentId = firstTabButton?.getAttribute("data-tab");
  const firstTabContent = document.getElementById(firstTabContentId);
  tabButtons.forEach((btn) => btn.classList.remove("active"));
  tabContents.forEach((content) => content.classList.remove("active"));
  if (firstTabButton && firstTabContent) {
    firstTabButton.classList.add("active");
    firstTabContent.classList.add("active");
  }
}

// ---------------------------------------------------------------------------
// JSON storage helpers
// ---------------------------------------------------------------------------

function readJsonStorage(key, fallbackValue) {
  try {
    const rawValue = roamingStorage.getItem(key);
    return rawValue ? JSON.parse(rawValue) : fallbackValue;
  } catch (error) {
    console.error(`Error reading ${key}:`, error);
    return fallbackValue;
  }
}

function writeJsonStorage(key, value) {
  roamingStorage.setItem(key, JSON.stringify(value));
}

function getSettings() {
  return readJsonStorage(SETTINGS_KEY, createBlankSettings());
}

function saveSettingsToStore(settings) {
  writeJsonStorage(SETTINGS_KEY, settings);
}

function getSavedTemplateDefaults() {
  return readJsonStorage(TEMPLATE_DEFAULTS_KEY, createBlankPromptTemplates());
}

function saveTemplateDefaults(templates) {
  writeJsonStorage(TEMPLATE_DEFAULTS_KEY, templates);
}

// ---------------------------------------------------------------------------
// Prompt template form helpers
// ---------------------------------------------------------------------------

function getPromptFieldValue(templateKey) {
  const fieldId = PROMPT_FIELD_MAP[templateKey];
  const field = fieldId ? document.getElementById(fieldId) : null;
  return field ? field.value : "";
}

function setPromptFieldValue(templateKey, value) {
  const fieldId = PROMPT_FIELD_MAP[templateKey];
  const field = fieldId ? document.getElementById(fieldId) : null;
  if (field) {
    field.value = value || "";
  }
}

function collectPromptTemplatesFromForm() {
  return TEMPLATE_KEYS.reduce((templates, key) => {
    templates[key] = getPromptFieldValue(key);
    return templates;
  }, {});
}

function applyPromptTemplatesToForm(templates) {
  TEMPLATE_KEYS.forEach((key) => {
    setPromptFieldValue(key, templates[key] || "");
  });
}

// ---------------------------------------------------------------------------
// Model catalog
// ---------------------------------------------------------------------------

function getCachedModelCatalog() {
  const cachedModels = readJsonStorage(MODEL_CATALOG_CACHE_KEY, []);
  return Array.isArray(cachedModels) && cachedModels.length ? cachedModels : getDefaultModels();
}

function persistModelCatalog(models) {
  availableModelCatalog = Array.isArray(models) && models.length ? [...models] : getDefaultModels();
  roamingStorage.setItem(MODEL_CATALOG_CACHE_KEY, JSON.stringify(availableModelCatalog));
  saveRoamingStoreAsync().catch((error) => {
    console.error("Error saving model catalog:", error);
  });
}

function setSettingsStatus(elementId, message) {
  const element = document.getElementById(elementId);
  if (element) {
    element.textContent = message;
  }
}

function updateAuthenticationStatus() {
  if (getApiKey()) {
    setSettingsStatus(
      "dropdown-auth-status",
      "Authentication source: saved Outlook add-in settings."
    );
    return;
  }

  setSettingsStatus(
    "dropdown-auth-status",
    "Authentication source: empty. Enter the API key in this screen and save settings."
  );
}

function updateModelSelectOptions(selectId, models, selectedValue) {
  const select = document.getElementById(selectId);
  if (!select) {
    return;
  }

  const catalog = Array.isArray(models) && models.length ? models : getDefaultModels();
  const normalizedSelection = typeof selectedValue === "string" ? selectedValue.trim() : "";
  const nextValue =
    normalizedSelection && catalog.includes(normalizedSelection) ? normalizedSelection : "";

  select.innerHTML = "";
  const placeholderOption = document.createElement("option");
  placeholderOption.value = "";
  placeholderOption.textContent = "Select a model";
  select.appendChild(placeholderOption);
  catalog.forEach((modelName) => {
    const option = document.createElement("option");
    option.value = modelName;
    option.textContent = modelName;
    select.appendChild(option);
  });

  select.value = nextValue || "";
}

function syncModelDropdowns() {
  const settings = getSettings();
  const models = availableModelCatalog.length ? availableModelCatalog : getCachedModelCatalog();
  updateModelSelectOptions("dropdown-model", models, settings.model || "");
  updateModelSelectOptions("dropdown-reply-model", models, settings.replyModel || "");
}

async function refreshModelCatalog(options = {}) {
  const silent = options.silent === true;
  const apiKey = getApiKey();

  updateAuthenticationStatus();

  if (!apiKey) {
    availableModelCatalog = getCachedModelCatalog();
    syncModelDropdowns();
    setSettingsStatus(
      "dropdown-model-status",
      "Using cached/fallback models. Enter an API key in Settings > General to enable live refresh."
    );
    return availableModelCatalog;
  }

  setSettingsStatus("dropdown-model-status", "Refreshing available models...");

  try {
    const liveModels = await fetchAvailableModels({ apiKey });
    persistModelCatalog(liveModels);
    syncModelDropdowns();
    setSettingsStatus(
      "dropdown-model-status",
      `Loaded ${liveModels.length} models from OpenRouter.`
    );
    return availableModelCatalog;
  } catch (error) {
    console.error("Error refreshing model catalog:", error);
    availableModelCatalog = getCachedModelCatalog();
    syncModelDropdowns();
    setSettingsStatus(
      "dropdown-model-status",
      `Live refresh failed. Using cached/fallback models (${error.message}).`
    );
    if (!silent) {
      showNotification(`Model refresh failed: ${error.message}`, "warning");
    }
    return availableModelCatalog;
  }
}

// ---------------------------------------------------------------------------
// Office.onReady
// ---------------------------------------------------------------------------

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    migrateSettingsKeys();
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    initializeApp();

    let autorunEnabled = false;
    let selectedOption = null;
    try {
      const settings = getSettings();
      autorunEnabled = settings.autorun === "true";
      selectedOption = settings.autorunOption;
    } catch (error) {
      console.error("Error getting autorun settings:", error);
    }

    if (autorunEnabled && selectedOption) {
      runAutorunOption(selectedOption);
    }

    // Main action buttons
    document.getElementById("summarize").addEventListener("click", summarizeEmail);
    document.getElementById("calendar-event").addEventListener("click", handleCalendarEvent);
    document.getElementById("settings-toggle").addEventListener("click", toggleSettingsView);
    document.getElementById("close-settings-view").addEventListener("click", toggleSettingsView);
    document
      .getElementById("dropdown-save-settings")
      .addEventListener("click", saveDropdownSettings);
    document.getElementById("dropdown-reset-all").addEventListener("click", resetAllSettings);

    const resetTemplatesBtn = document.getElementById("dropdown-reset-templates");
    if (resetTemplatesBtn) resetTemplatesBtn.addEventListener("click", resetTemplates);

    const loadBuiltInsBtn = document.getElementById("dropdown-load-builtins");
    if (loadBuiltInsBtn) loadBuiltInsBtn.addEventListener("click", loadBuiltInTemplateDefaults);

    const saveTemplateDefaultsBtn = document.getElementById("dropdown-save-template-defaults");
    if (saveTemplateDefaultsBtn)
      saveTemplateDefaultsBtn.addEventListener("click", saveCurrentTemplatesAsDefaults);

    const clearTemplatesBtn = document.getElementById("dropdown-clear-templates");
    if (clearTemplatesBtn) clearTemplatesBtn.addEventListener("click", clearTemplates);

    const copyTemplatesBtn = document.getElementById("dropdown-copy-templates");
    if (copyTemplatesBtn) copyTemplatesBtn.addEventListener("click", copyAllTemplatesToClipboard);

    const exportMarkdownBtn = document.getElementById("dropdown-export-markdown");
    if (exportMarkdownBtn) exportMarkdownBtn.addEventListener("click", exportTemplatesAsMarkdown);

    const refreshModelsBtn = document.getElementById("dropdown-refresh-models");
    if (refreshModelsBtn)
      refreshModelsBtn.addEventListener("click", () => {
        refreshModelCatalog();
      });

    const apiKeyInput = document.getElementById("dropdown-api-key");
    if (apiKeyInput) apiKeyInput.addEventListener("input", updateAuthenticationStatus);

    document.getElementById("dropdown-dev-mode").addEventListener("change", function () {
      const devServerGroup = document.getElementById("dev-server-group");
      devServerGroup.style.display = this.value === "true" ? "block" : "none";
    });

    document.getElementById("expand-content").addEventListener("click", expandContent);
    document.getElementById("generate-reply").addEventListener("click", generateReply);

    const copyReplyBtn = document.getElementById("copy-reply");
    if (copyReplyBtn) copyReplyBtn.addEventListener("click", copyReply);

    // Copy as dropdown
    initCopyAsDropdown();

    loadDropdownSettings();
    updateAuthenticationStatus();
    refreshModelCatalog({ silent: true });
    applyCurrentTheme();

    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.SettingsChanged, onSettingsChanged);
    }

    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
      const settings = getSettings();
      if (settings.autorun === "true" && settings.autorunOption) {
        runAutorunOption(settings.autorunOption);
      }
      updateCalendarButtonState();
    });

    updateCalendarButtonState();
    initializeSettingsTabs();
  }
});

function runAutorunOption(option) {
  switch (option) {
    case "summarize":
      summarizeEmail();
      break;
    case "reply":
      generateReply();
      break;
  }
}

// ---------------------------------------------------------------------------
// Copy as dropdown
// ---------------------------------------------------------------------------

function initCopyAsDropdown() {
  const toggle = document.getElementById("copy-as-toggle");
  const dropdown = document.getElementById("copy-as-dropdown");

  if (!toggle || !dropdown) return;

  toggle.addEventListener("click", (e) => {
    e.stopPropagation();
    dropdown.classList.toggle("open");
  });

  document.addEventListener("click", () => {
    dropdown.classList.remove("open");
  });

  document.getElementById("copy-as-markdown").addEventListener("click", copyAsMarkdown);
  document.getElementById("copy-as-plaintext").addEventListener("click", copyAsPlaintext);
  document.getElementById("copy-as-html").addEventListener("click", copyAsHtml);
}

async function copyAsMarkdown() {
  document.getElementById("copy-as-dropdown").classList.remove("open");
  try {
    await navigator.clipboard.writeText(lastRawContent);
    showCopyStatus("copy-status", "Copied as Markdown!");
  } catch {
    showNotification("Failed to copy", "error");
  }
}

async function copyAsPlaintext() {
  document.getElementById("copy-as-dropdown").classList.remove("open");
  try {
    // Render to a temp element and grab innerText for clean plaintext
    const temp = document.createElement("div");
    temp.innerHTML = marked.parse(lastRawContent);
    await navigator.clipboard.writeText(temp.innerText);
    showCopyStatus("copy-status", "Copied as Plaintext!");
  } catch {
    showNotification("Failed to copy", "error");
  }
}

async function copyAsHtml() {
  document.getElementById("copy-as-dropdown").classList.remove("open");
  try {
    const html =
      document.getElementById("result-content").innerHTML ||
      document.getElementById("tldr-content").innerHTML;
    await navigator.clipboard.writeText(html);
    showCopyStatus("copy-status", "Copied as HTML!");
  } catch {
    showNotification("Failed to copy", "error");
  }
}

function showCopyStatus(elementId, message) {
  const el = document.getElementById(elementId);
  if (!el) return;
  el.textContent = message;
  setTimeout(() => {
    el.textContent = "";
  }, 2000);
}

// ---------------------------------------------------------------------------
// Theme
// ---------------------------------------------------------------------------

function applyCurrentTheme() {
  const savedTheme = getSettings().theme || "dark";

  if (savedTheme === "light") {
    document.body.setAttribute("data-theme", "light");
    document.body.classList.remove("dark-theme");
  } else if (savedTheme === "dark") {
    document.body.setAttribute("data-theme", "dark");
    document.body.classList.add("dark-theme");
  } else {
    if (Office.context.officeTheme) {
      const bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
      if (bodyBackgroundColor) {
        if (isDarkTheme(bodyBackgroundColor)) {
          document.body.setAttribute("data-theme", "dark");
          document.body.classList.add("dark-theme");
        } else {
          document.body.setAttribute("data-theme", "light");
          document.body.classList.remove("dark-theme");
        }
      } else {
        document.body.setAttribute("data-theme", "dark");
        document.body.classList.add("dark-theme");
      }
    }
  }
}

function isDarkTheme(color) {
  if (!color || typeof color !== "string") {
    return false;
  }

  try {
    color = color.replace("#", "");
    const r = parseInt(color.substr(0, 2), 16);
    const g = parseInt(color.substr(2, 2), 16);
    const b = parseInt(color.substr(4, 2), 16);
    if (isNaN(r) || isNaN(g) || isNaN(b)) {
      return false;
    }
    return 0.299 * r + 0.587 * g + 0.114 * b < 128;
  } catch (error) {
    console.error("Error calculating theme brightness:", error);
    return false;
  }
}

function onSettingsChanged() {
  const savedTheme = getSettings().theme || "system";
  if (savedTheme === "system") {
    applyCurrentTheme();
  }
}

// ---------------------------------------------------------------------------
// Notifications
// ---------------------------------------------------------------------------

function showNotification(message, type = "info") {
  const existingNotification = document.getElementById("notification");
  if (existingNotification) {
    existingNotification.remove();
  }

  const notification = document.createElement("div");
  notification.id = "notification";
  notification.className = `notification ${type}`;
  notification.textContent = message;
  document.body.appendChild(notification);
  notification.offsetHeight;
  notification.style.animation = "slideInFromTop 0.3s ease-out";

  setTimeout(() => {
    notification.style.animation = "slideOutToTop 0.3s ease-out";
    setTimeout(() => {
      notification.remove();
    }, 300);
  }, 3000);
}

// ---------------------------------------------------------------------------
// Settings save / load / reset
// ---------------------------------------------------------------------------

async function saveDropdownSettings() {
  try {
    const settings = getSettings();

    const apiKey = document.getElementById("dropdown-api-key").value.trim();
    const model = document.getElementById("dropdown-model").value;
    const theme = document.getElementById("dropdown-theme").value;
    const fontSize = document.getElementById("dropdown-font-size").value;
    const tldrMode = document.getElementById("dropdown-tldr-mode").value;
    const showReply = document.getElementById("dropdown-show-reply").value;
    const replyModel = document.getElementById("dropdown-reply-model")
      ? document.getElementById("dropdown-reply-model").value
      : undefined;
    const autorun = document.getElementById("dropdown-autorun").value;
    const autorunOption = document.getElementById("dropdown-autorun-option").value;
    const devMode = document.getElementById("dropdown-dev-mode").value;
    const devServer = document.getElementById("dropdown-dev-server").value;
    const templates = collectPromptTemplatesFromForm();

    settings.apiKey = apiKey;
    settings.model = model;
    settings.theme = theme;
    settings.fontSize = fontSize;
    settings.tldrMode = tldrMode;
    settings.showReply = showReply;
    settings.replyModel = replyModel || "";
    settings.autorun = autorun;
    settings.autorunOption = autorunOption;
    settings.devMode = devMode;
    settings.devServer = devServer;
    settings.templates = {
      ...createBlankPromptTemplates(),
      ...templates,
    };

    saveSettingsToStore(settings);
    roamingStorage.setItem("theme", theme);
    await saveRoamingStoreAsync();

    applyCurrentTheme();
    applyFontSize(fontSize);
    updateReplyButtonVisibility(showReply === "true");
    updateDevBadges(devMode === "true");
    updateAuthenticationStatus();

    showNotification("All settings saved successfully");
    toggleSettingsView();
  } catch (error) {
    console.error("Error saving Outlook settings:", error);
    showNotification(`Failed to save settings: ${error.message}`, "error");
  }
}

function loadDropdownSettings() {
  try {
    const settings = getSettings();
    const sessionDefaults = getSavedTemplateDefaults();

    document.getElementById("dropdown-api-key").value = settings.apiKey || "";
    if (settings.theme) document.getElementById("dropdown-theme").value = settings.theme;
    if (settings.fontSize) document.getElementById("dropdown-font-size").value = settings.fontSize;
    if (settings.tldrMode) document.getElementById("dropdown-tldr-mode").value = settings.tldrMode;
    if (settings.showReply)
      document.getElementById("dropdown-show-reply").value = settings.showReply;
    if (settings.devServer !== undefined)
      document.getElementById("dropdown-dev-server").value = settings.devServer;
    if (settings.autorun) document.getElementById("dropdown-autorun").value = settings.autorun;
    if (settings.autorunOption)
      document.getElementById("dropdown-autorun-option").value = settings.autorunOption;
    if (settings.devMode) document.getElementById("dropdown-dev-mode").value = settings.devMode;

    const devServerGroup = document.getElementById("dev-server-group");
    if (devServerGroup) {
      devServerGroup.style.display = settings.devMode === "true" ? "block" : "none";
    }

    updateDevBadges(settings.devMode === "true");
    applyFontSize(settings.fontSize || DEFAULT_SETTINGS.fontSize);
    updateReplyButtonVisibility(settings.showReply === "true");

    const templates =
      settings.templates && Object.keys(settings.templates).length
        ? settings.templates
        : sessionDefaults;
    applyPromptTemplatesToForm({
      ...createBlankPromptTemplates(),
      ...templates,
    });

    availableModelCatalog = getCachedModelCatalog();
    syncModelDropdowns();
    updateAuthenticationStatus();
  } catch (error) {
    console.error("Error loading dropdown settings:", error);
    applyPromptTemplatesToForm(createBlankPromptTemplates());
  }
}

async function resetAllSettings() {
  try {
    const blankSettings = createBlankSettings();

    document.getElementById("dropdown-api-key").value = "";
    document.getElementById("dropdown-model").value = blankSettings.model;
    document.getElementById("dropdown-theme").value = blankSettings.theme;
    document.getElementById("dropdown-font-size").value = blankSettings.fontSize;
    document.getElementById("dropdown-tldr-mode").value = blankSettings.tldrMode;
    document.getElementById("dropdown-show-reply").value = blankSettings.showReply;
    document.getElementById("dropdown-reply-model").value = blankSettings.replyModel;
    document.getElementById("dropdown-autorun").value = blankSettings.autorun;
    document.getElementById("dropdown-autorun-option").value = blankSettings.autorunOption;
    document.getElementById("dropdown-dev-mode").value = blankSettings.devMode;
    document.getElementById("dropdown-dev-server").value = blankSettings.devServer;
    document.getElementById("dev-server-group").style.display = "none";

    applyPromptTemplatesToForm(blankSettings.templates);

    roamingStorage.removeItem(SETTINGS_KEY);
    roamingStorage.removeItem(TEMPLATE_DEFAULTS_KEY);
    roamingStorage.removeItem(MODEL_CATALOG_CACHE_KEY);
    roamingStorage.removeItem("theme");
    await saveRoamingStoreAsync();
    persistModelCatalog(getDefaultModels());
    syncModelDropdowns();
    updateAuthenticationStatus();
    setSettingsStatus(
      "dropdown-model-status",
      "Saved settings cleared. Model catalog reset to defaults."
    );

    applyCurrentTheme();
    applyFontSize(blankSettings.fontSize);
    updateReplyButtonVisibility(blankSettings.showReply === "true");
    updateDevBadges(false);

    showNotification("All saved Outlook settings cleared", "success");
    initializeSettingsTabs();
  } catch (error) {
    console.error("Error resetting Outlook settings:", error);
    showNotification(`Failed to reset settings: ${error.message}`, "error");
  }
}

// ---------------------------------------------------------------------------
// Template actions
// ---------------------------------------------------------------------------

function resetTemplates() {
  const currentSettings = getSettings();
  const sessionDefaults = getSavedTemplateDefaults();
  currentSettings.templates = { ...createBlankPromptTemplates(), ...sessionDefaults };
  saveSettingsToStore(currentSettings);
  saveRoamingStoreAsync().catch((error) => {
    console.error("Error saving template reset:", error);
  });
  applyPromptTemplatesToForm(currentSettings.templates);
  showNotification(PROVIDER_MESSAGES.templatesReset);
}

function clearTemplates() {
  const currentSettings = getSettings();
  currentSettings.templates = createBlankPromptTemplates();
  saveSettingsToStore(currentSettings);
  saveRoamingStoreAsync().catch((error) => {
    console.error("Error saving cleared templates:", error);
  });
  applyPromptTemplatesToForm(currentSettings.templates);
  showNotification(PROVIDER_MESSAGES.templatesCleared);
}

function loadBuiltInTemplateDefaults() {
  applyPromptTemplatesToForm(createDefaultPromptTemplates());
  showNotification("Built-in prompt defaults loaded into the form.");
}

function saveCurrentTemplatesAsDefaults() {
  const templates = collectPromptTemplatesFromForm();
  const currentSettings = getSettings();
  currentSettings.templates = { ...createBlankPromptTemplates(), ...templates };
  saveTemplateDefaults(templates);
  saveSettingsToStore(currentSettings);
  saveRoamingStoreAsync().catch((error) => {
    console.error("Error saving template defaults:", error);
  });
  showNotification(PROVIDER_MESSAGES.sessionDefaultsSaved);
}

async function copyAllTemplatesToClipboard() {
  try {
    const templates = collectPromptTemplatesFromForm();
    const lines = [
      "# Sidekick Prompt Templates",
      "",
      ...Object.entries(templates).flatMap(([key, value]) => [
        `## ${key}`,
        "```",
        value || "",
        "```",
        "",
      ]),
    ];
    await navigator.clipboard.writeText(lines.join("\n"));
    showNotification("All prompt templates copied to clipboard.", "success");
  } catch (error) {
    console.error("Error copying prompt templates:", error);
    showNotification("Failed to copy prompt templates.", "error");
  }
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------

function applyFontSize(size) {
  document.documentElement.setAttribute("data-font-size", size || "medium");
}

function updateReplyButtonVisibility(show) {
  const replyButton = document.getElementById("generate-reply");
  if (replyButton) {
    replyButton.style.display = show ? "inline-block" : "none";
  }
}

function updateDevBadges(show) {
  const devBadge = document.getElementById("dev-badge");
  const footerDevBadge = document.getElementById("footer-dev-badge");
  if (devBadge) devBadge.style.display = show ? "block" : "none";
  if (footerDevBadge) footerDevBadge.style.display = show ? "block" : "none";
}

function setTheme(theme) {
  const root = document.documentElement;
  if (theme === "system") {
    if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) {
      root.setAttribute("data-theme", "dark");
    } else {
      root.setAttribute("data-theme", "light");
    }
  } else {
    root.setAttribute("data-theme", theme);
  }
}

function setFontSize(size) {
  document.documentElement.style.setProperty("--result-font-size", getFontSizeValue(size));
}

function getFontSizeValue(size) {
  switch (size) {
    case "small":
      return "0.875rem";
    case "large":
      return "1.125rem";
    default:
      return "1rem";
  }
}

// ---------------------------------------------------------------------------
// Email access
// ---------------------------------------------------------------------------

function getApiKey() {
  const input = document.getElementById("dropdown-api-key");
  if (input && typeof input.value === "string" && input.value.trim()) {
    return input.value.trim();
  }

  const settings = getSettings();
  return typeof settings.apiKey === "string" ? settings.apiKey.trim() : "";
}

async function getEmailContent() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to get email content"));
      }
    });
  });
}

function getTemplateValue(templateKey) {
  const settings = getSettings();
  return settings.templates && typeof settings.templates[templateKey] === "string"
    ? settings.templates[templateKey].trim()
    : "";
}

function requireTemplate(templateKey, label) {
  const template = getTemplateValue(templateKey);
  if (!template) {
    throw new Error(`${label} prompt is empty. Configure it in Settings > Templates.`);
  }

  return template;
}

function requireModel(settingKey, label) {
  const settings = getSettings();
  const model = typeof settings[settingKey] === "string" ? settings[settingKey].trim() : "";
  if (!model) {
    throw new Error(`${label} is empty. Select a model in Settings > General.`);
  }

  return model;
}

// ---------------------------------------------------------------------------
// LLM calls
// ---------------------------------------------------------------------------

async function generateContent(prompt, apiKey, modelOverride = null, isTldr = false) {
  const model = modelOverride || requireModel("model", "Primary model");

  try {
    return await generateText(prompt, {
      apiKey,
      model,
      maxTokens: isTldr ? 800 : 8192,
      temperature: 0.4,
    });
  } catch (error) {
    console.error("Error generating content:", error);
    throw error;
  } finally {
    if (!isTldr) {
      document.getElementById("loading").style.display = "none";
    }
  }
}

async function generateTldrContent(prompt, apiKey, modelOverride = null) {
  const subject = Office.context.mailbox.item.subject;
  const emailContent = await getEmailContent();

  const tldrPrompt = requireTemplate("tldrPrompt", "TL;DR")
    .replace("{subject}", subject)
    .replace("{content}", emailContent);

  return generateContent(tldrPrompt, apiKey, modelOverride, true);
}

// ---------------------------------------------------------------------------
// Summarize
// ---------------------------------------------------------------------------

async function summarizeEmail() {
  const apiKey = getApiKey();
  if (!apiKey) {
    showNotification(PROVIDER_MESSAGES.missingApiKey, "error");
    toggleSettingsView();
    return;
  }

  showLoading("Summarizing email...");

  try {
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;
    const template = requireTemplate("summarize", "Summarize");
    const prompt = template.replace("{subject}", subject).replace("{content}", emailContent);

    const tldrMode = getSettings().tldrMode === "true";

    if (tldrMode) {
      const tldrContent = await generateTldrContent(prompt, apiKey);
      hideLoading();
      showResults(tldrContent, TYPES.SUMMARIZE);

      const fullContent = await generateContent(prompt, apiKey);
      updateResults(fullContent);
      updateExpandButton(true);
    } else {
      const summary = await generateContent(prompt, apiKey);
      showResults(summary, TYPES.SUMMARIZE);
    }
  } catch (error) {
    hideLoading();
    showNotification(`Error: ${error.message}`, "error");
  }
}

// ---------------------------------------------------------------------------
// Reply
// ---------------------------------------------------------------------------

async function generateReply() {
  const apiKey = getApiKey();
  if (!apiKey) {
    showNotification(PROVIDER_MESSAGES.missingApiKey, "error");
    toggleSettingsView();
    return;
  }

  showLoading("Generating reply...");

  try {
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;
    const template = requireTemplate("reply", "Reply");
    const prompt = template.replace("{subject}", subject).replace("{content}", emailContent);

    let replyModelOverride = null;
    try {
      replyModelOverride = requireModel("replyModel", "Reply model");
    } catch {
      // fall back to primary model
    }

    const result = await generateContent(prompt, apiKey, replyModelOverride);
    lastRawContent = result;
    const formattedReply = formatReplyOutput(result);

    document.getElementById("tldr-content").innerHTML = formattedReply.html;
    document.getElementById("result-content").innerHTML = formattedReply.html;
    document.getElementById("result-section").style.display = "block";
    document.getElementById("copy-reply").style.display = "inline-block";
  } catch (error) {
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    hideLoading();
  }
}

function formatReplyOutput(replyText) {
  let subject = "";
  let body = "";

  const subjectMatch = replyText.match(/^(?:SUBJECT:|Subject:)\s*(.+?)(?:\n|$)/m);
  if (subjectMatch) {
    subject = subjectMatch[1].trim();
    body = replyText.replace(/^(?:SUBJECT:|Subject:)\s*.+?\n+/m, "").trim();
  } else {
    const lines = replyText.trim().split("\n");
    if (lines.length > 0) {
      subject = lines[0].trim();
      body = lines.slice(1).join("\n").trim();
    } else {
      subject = "Re: Your email";
      body = replyText.trim();
    }
  }

  const formattedHtml = `
    <div class="reply-container">
      <div class="reply-subject">
        <span class="reply-label">Subject:</span>${subject}
      </div>
      <div class="reply-body">${body}</div>
    </div>
  `;

  return {
    html: formattedHtml,
    subject: subject,
    body: body,
    raw: `Subject: ${subject}\n\n${body}`,
  };
}

function copyReply() {
  const tldrContent = document.getElementById("tldr-content").innerText;
  const subjectMatch = tldrContent.match(/Subject:\s*(.+?)(?:\n|$)/i);
  const subject = subjectMatch ? subjectMatch[1].trim() : "";
  const bodyStart = tldrContent.indexOf(subject) + subject.length;
  const body = tldrContent.substring(bodyStart).trim();
  const replyContent = `Subject: ${subject}\n\n${body}`;

  navigator.clipboard
    .writeText(replyContent)
    .then(() => {
      showCopyStatus("copy-reply-status", "Copied!");
    })
    .catch(() => {
      showNotification("Failed to copy reply", "error");
    });
}

// ---------------------------------------------------------------------------
// Result display
// ---------------------------------------------------------------------------

function showLoading(message = "Loading...") {
  document.getElementById("loading").style.display = "block";
  document.getElementById("loading-message").textContent = message;
  document.getElementById("landing-screen").style.display = "none";
  document.getElementById("result-section").style.display = "none";
}

function hideLoading() {
  const loadingSection = document.getElementById("loading");
  if (loadingSection) {
    loadingSection.style.display = "none";
  }
}

function showResults(content, type) {
  lastRawContent = content;

  document.getElementById("result-content").innerHTML = "";
  document.getElementById("tldr-content").innerHTML = "";

  const expandButton = document.getElementById("expand-content");
  if (expandButton) {
    expandButton.disabled = true;
    expandButton.classList.add("ms-Button--disabled");
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
    expandButton.classList.remove("ms-Button--primary");
  }

  document.getElementById("loading").style.display = "none";
  document.getElementById("result-section").style.display = "block";
  document.getElementById("landing-screen").style.display = "none";
  document.getElementById("app-body").style.display = "block";

  const tldrMode = getSettings().tldrMode === "true";

  if (tldrMode) {
    document.getElementById("tldr-content").innerHTML = marked.parse(content);
    if (expandButton) {
      expandButton.disabled = true;
      expandButton.classList.add("ms-Button--disabled");
      expandButton.innerHTML = '<span class="ms-Button-label">Loading Full Content...</span>';
    }
  } else {
    document.getElementById("tldr-content").innerHTML = marked.parse(content);
    document.getElementById("result-content").innerHTML = marked.parse(content);
    document.getElementById("full-content-container").style.display = "block";
    if (expandButton) {
      expandButton.disabled = false;
      expandButton.classList.remove("ms-Button--disabled");
      expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
      expandButton.classList.add("ms-Button--primary");
    }
  }

  // Show/hide reply-specific buttons
  const copyReplyButton = document.getElementById("copy-reply");
  if (copyReplyButton) {
    copyReplyButton.style.display = type === TYPES.REPLY ? "inline-block" : "none";
  }

  const generateReplyButton = document.getElementById("generate-reply");
  if (generateReplyButton) {
    const showReply = getSettings().showReply === "true";
    generateReplyButton.style.display =
      type === TYPES.REPLY ? "none" : showReply ? "inline-block" : "none";
  }

  applyFontSize(getSettings().fontSize || "medium");
}

function updateResults(content) {
  lastRawContent = content;

  const expandButton = document.getElementById("expand-content");
  if (expandButton) {
    expandButton.disabled = false;
    expandButton.classList.remove("ms-Button--disabled");
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
  }

  const loadingContainer = document.getElementById("loading-container");
  if (loadingContainer) loadingContainer.remove();

  const resultContent = document.getElementById("result-content");
  if (resultContent) {
    resultContent.innerHTML = marked.parse(content);
  }
}

function updateExpandButton(isFullContentVisible) {
  const expandButton = document.getElementById("expand-content");
  if (!expandButton) return;
  expandButton.disabled = !isFullContentVisible;
  expandButton.classList.toggle("ms-Button--disabled", !isFullContentVisible);
  if (!isFullContentVisible) {
    expandButton.innerHTML = '<span class="ms-Button-label">Loading Full Content...</span>';
    expandButton.classList.remove("ms-Button--primary");
  } else {
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
    expandButton.classList.add("ms-Button--primary");
  }
}

function expandContent() {
  const expandButton = document.getElementById("expand-content");
  if (expandButton.disabled) return;

  const fullContentContainer = document.getElementById("full-content-container");
  if (fullContentContainer.style.display === "none") {
    fullContentContainer.style.display = "block";
    expandButton.innerHTML = '<span class="ms-Button-label">Hide Full Content</span>';
    expandButton.classList.remove("ms-Button--primary");
  } else {
    fullContentContainer.style.display = "none";
    expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
    expandButton.classList.add("ms-Button--primary");
  }
}

// ---------------------------------------------------------------------------
// Calendar
// ---------------------------------------------------------------------------

async function parseEventDetailsWithZai(emailContent) {
  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error(PROVIDER_MESSAGES.missingApiKey);
  }

  const langInstructions =
    "Event title should be in English. If the event has a type or category, include it in square brackets ([]) at the beginning.";

  const prompt = requireTemplate("calendarParse", "Calendar parse")
    .replace("{languageInstructions}", langInstructions)
    .replace("{content}", emailContent);

  const result = await generateContent(prompt, apiKey, null, false);

  const jsonMatch = result.match(/\{[\s\S]*\}/);
  const jsonText = jsonMatch ? jsonMatch[0] : result;

  try {
    const eventDetails = JSON.parse(jsonText);

    if (!eventDetails.subject) throw new Error("Event title not found.");
    if (!eventDetails.start?.dateTime) throw new Error("Event start time not found.");
    if (!eventDetails.end?.dateTime) throw new Error("Event end time not found.");

    const dateRegex = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/;
    if (!dateRegex.test(eventDetails.start.dateTime)) {
      throw new Error("Invalid start time format. (Should be YYYY-MM-DDTHH:mm:ss)");
    }
    if (!dateRegex.test(eventDetails.end.dateTime)) {
      throw new Error("Invalid end time format. (Should be YYYY-MM-DDTHH:mm:ss)");
    }

    return eventDetails;
  } catch (parseError) {
    throw new Error("Failed to extract event information: " + parseError.message);
  }
}

async function createCalendarEvent(eventDetails) {
  if (!eventDetails.subject || !eventDetails.start?.dateTime || !eventDetails.end?.dateTime) {
    throw new Error("Required event information is missing.");
  }

  const startDate = new Date(eventDetails.start.dateTime);
  const endDate = new Date(eventDetails.end.dateTime);
  const requiredAttendees = [];
  const optionalAttendees = [];

  if (eventDetails.attendees && eventDetails.attendees.length > 0) {
    eventDetails.attendees.forEach((attendee) => {
      if (attendee.emailAddress && attendee.emailAddress.address) {
        if (attendee.type === "optional") {
          optionalAttendees.push(attendee.emailAddress.address);
        } else {
          requiredAttendees.push(attendee.emailAddress.address);
        }
      }
    });
  }

  Office.context.mailbox.displayNewAppointmentForm({
    requiredAttendees,
    optionalAttendees,
    start: startDate,
    end: endDate,
    location: eventDetails.location?.displayName || "",
    body: eventDetails.body?.content || "",
    subject: eventDetails.subject,
  });

  showNotification(`Event '${eventDetails.subject}' has been created.`, "info");
  return true;
}

async function checkIfCalendarEvent(emailContent) {
  try {
    const apiKey = getApiKey();
    if (!apiKey) return false;

    const prompt = requireTemplate("calendarCheck", "Calendar check").replace(
      "{content}",
      emailContent
    );

    const result = await generateContent(prompt, apiKey, null, true);
    return result.toLowerCase().trim() === "true";
  } catch {
    return false;
  }
}

async function handleCalendarEvent() {
  const apiKey = getApiKey();
  if (!apiKey) {
    showNotification(PROVIDER_MESSAGES.missingApiKey, "error");
    toggleSettingsView();
    return;
  }

  showLoading("Creating calendar event...");

  try {
    const emailContent = await getEmailContent();

    document.getElementById("landing-screen").style.display = "none";
    document.getElementById("result-section").style.display = "block";

    document.getElementById("tldr-content").innerHTML = `
      <div class="email-content-header">
        <h3>Email Content</h3>
        <button id="copy-email-content" class="ms-Button ms-Button--primary">
          <span class="ms-Button-label">Copy to Clipboard</span>
        </button>
      </div>
      <div class="email-content-body">
        <pre style="white-space: pre-wrap; word-break: break-word;">${emailContent}</pre>
      </div>
    `;

    document.getElementById("copy-email-content").addEventListener("click", function () {
      navigator.clipboard
        .writeText(emailContent)
        .then(() => {
          showNotification("Email content copied to clipboard", "info");
        })
        .catch(() => {
          showNotification("Failed to copy to clipboard", "error");
        });
    });

    try {
      const eventDetails = await parseEventDetailsWithZai(emailContent);

      document.getElementById("result-content").innerHTML = `
        <div class="event-details-header"><h3>Extracted Event Details</h3></div>
        <div class="event-details-body">
          <p><strong>Subject:</strong> ${eventDetails.subject || "Not found"}</p>
          <p><strong>Start:</strong> ${eventDetails.start?.dateTime || "Not found"}</p>
          <p><strong>End:</strong> ${eventDetails.end?.dateTime || "Not found"}</p>
          <p><strong>Location:</strong> ${eventDetails.location?.displayName || "Not found"}</p>
        </div>
      `;

      await createCalendarEvent(eventDetails);
    } catch (extractionError) {
      const errorMessage = extractionError.message;
      const cleanedMessage = errorMessage.includes("Failed to extract event information:")
        ? errorMessage.split("Failed to extract event information:")[1].trim()
        : errorMessage;

      document.getElementById("result-content").innerHTML = `
        <div class="event-details-header"><h3>Event Extraction Failed</h3></div>
        <div class="event-details-body">
          <p class="error-message">${cleanedMessage}</p>
        </div>
      `;

      showNotification(`Event extraction failed: ${cleanedMessage}`, "error");
    }
  } catch (error) {
    showNotification(`Error: ${error.message}`, "error");
  } finally {
    hideLoading();
    updateCalendarButtonState();

    const expandButton = document.getElementById("expand-content");
    if (expandButton) {
      expandButton.disabled = false;
      expandButton.classList.remove("ms-Button--disabled");
      expandButton.innerHTML = '<span class="ms-Button-label">Show Full Content</span>';
      expandButton.classList.add("ms-Button--primary");
      document.getElementById("full-content-container").style.display = "block";
    }
  }
}

async function updateCalendarButtonState() {
  try {
    const emailContent = await getEmailContent();
    const isCalendarEvent = await checkIfCalendarEvent(emailContent);
    const calendarBtn = document.getElementById("calendar-event");
    if (calendarBtn) {
      calendarBtn.disabled = !isCalendarEvent;
      calendarBtn.classList.toggle("action-button--disabled", !isCalendarEvent);
      calendarBtn.classList.toggle("action-button--primary", isCalendarEvent);
    }
  } catch (error) {
    console.error("Error updating calendar button state:", error);
  }
}

// ---------------------------------------------------------------------------
// Export templates
// ---------------------------------------------------------------------------

function exportTemplatesAsMarkdown() {
  try {
    const settings = getSettings();
    const templates = settings.templates || createBlankPromptTemplates();
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10);

    let userInfo = "";
    try {
      if (Office.context.mailbox && Office.context.mailbox.userProfile) {
        const user = Office.context.mailbox.userProfile;
        userInfo = `\n\n*Exported by: ${user.displayName} (${user.emailAddress})*`;
      }
    } catch {
      // no-op
    }

    let markdownContent = `# Sidekick Prompt Templates\n\n`;
    markdownContent += `*Exported on: ${now.toLocaleString()}*${userInfo}\n\n`;
    markdownContent += `## General Settings\n\n`;
    markdownContent += `- **Provider**: OpenRouter\n`;
    markdownContent += `- **Model**: ${settings.model || "(empty)"}\n`;
    markdownContent += `- **Reply Model**: ${settings.replyModel || "(empty)"}\n\n`;
    markdownContent += `## Prompt Templates\n\n`;

    const exportedSections = [
      ["Summarize Template", templates.summarize],
      ["Reply Template", templates.reply],
      ["TL;DR Template", templates.tldrPrompt],
      ["Calendar Parse Template", templates.calendarParse],
      ["Calendar Check Template", templates.calendarCheck],
    ];

    exportedSections.forEach(([title, value]) => {
      markdownContent += `### ${title}\n\n\`\`\`\n${value || ""}\n\`\`\`\n\n`;
    });

    const blob = new Blob([markdownContent], { type: "text/markdown" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `sidekick-templates-${dateStr}.md`;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);

    showNotification("Templates exported successfully", "success");
  } catch (error) {
    console.error("Error exporting templates:", error);
    showNotification("Failed to export templates", "error");
  }
}

// ---------------------------------------------------------------------------
// App initialization
// ---------------------------------------------------------------------------

async function initializeApp() {
  const settingsSection = document.getElementById("settings-section");
  const resultSection = document.getElementById("result-section");

  if (settingsSection) settingsSection.style.display = "none";
  if (resultSection) resultSection.style.display = "none";

  const settings = loadSettings();
  if (settings) {
    if (settings.theme) setTheme(settings.theme);
    if (settings.fontSize) setFontSize(settings.fontSize);
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    showNotification(PROVIDER_MESSAGES.missingApiKey, "warning");
  }
}

function loadSettings() {
  try {
    return getSettings();
  } catch (error) {
    console.error("Error loading settings:", error);
    return createBlankSettings();
  }
}
