/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console, require */

const { DEFAULT_MODEL } = require("../shared/zaiConfig");
const { executeChatCompletion } = require("../shared/zaiClient");

const SETTINGS_KEY = "michael_settings";

Office.onReady(() => {
  // Office is ready
});

/**
 * Handles the Add-in Command button click.
 * Summarizes the current email body and inserts the result at the cursor.
 * @param {Office.AddinCommands.Event} event The event object.
 */
async function action(event) {
  const settings = getSavedSettings();
  const apiKey = getSavedApiKey(settings);
  if (!apiKey) {
    showErrorNotification("Open Sidekick Settings and save an OpenRouter API key first.", event);
    return;
  }

  const model = getSavedModel(settings);

  showProcessingNotification("Summarizing email...", event);

  try {
    const emailContent = await getEmailContent();
    const subject = Office.context.mailbox.item.subject;
    const prompt = `Summarize the following email briefly.\n\nSubject: ${subject}\nContent:\n${emailContent}`;
    const summary = await generateContent(prompt, model, apiKey);

    await replaceSelectionWithText(summary);

    showSuccessNotification("Summary inserted at cursor.", event);
  } catch (error) {
    console.error("Error during summarize command:", error);
    showErrorNotification(`Summarize failed: ${error.message}`, event);
  }
}

Office.actions.associate("action", action);

async function replaceSelectionWithText(textToInsert) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSelectedDataAsync(
      textToInsert,
      { coercionType: Office.CoercionType.Text },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to set selected data:", asyncResult.error);
          reject(new Error(`Failed to insert text: ${asyncResult.error.message}`));
          return;
        }

        resolve();
      }
    );
  });
}

async function getEmailContent() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.trim());
        return;
      }

      console.error("Failed to get email body as text:", result.error);
      reject(new Error("Failed to get email content."));
    });
  });
}

async function generateContent(prompt, modelName, apiKey) {
  const result = await executeChatCompletion({
    apiKey,
    userPrompt: prompt,
    model: modelName,
    temperature: 0.3,
  });

  return result.text;
}

function getSavedSettings() {
  try {
    const rawValue = Office.context?.roamingSettings?.get(SETTINGS_KEY);
    if (!rawValue || typeof rawValue !== "string") {
      return {};
    }

    const parsed = JSON.parse(rawValue);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (error) {
    console.error("Failed to read saved Outlook settings:", error);
    return {};
  }
}

function getSavedApiKey(settings) {
  return typeof settings?.apiKey === "string" ? settings.apiKey.trim() : "";
}

function getSavedModel(settings) {
  const configuredModel = typeof settings?.model === "string" ? settings.model.trim() : "";
  return configuredModel || DEFAULT_MODEL;
}

function showProcessingNotification(message) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("ProcessingNotification", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message,
    icon: "Icon.80x80",
    persistent: false,
  });
}

function showSuccessNotification(message, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionCompleteNotification",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "Icon.80x80",
      persistent: true,
    },
    (asyncResult) => {
      handleNotificationResult(asyncResult, event, "success");
    }
  );
}

function showErrorNotification(message, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionErrorNotification",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message,
      icon: "Icon.80x80",
      persistent: true,
    },
    (asyncResult) => {
      handleNotificationResult(asyncResult, event, "error");
    }
  );
}

function handleNotificationResult(asyncResult, event, type) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.error(`Failed to show ${type} notification: `, asyncResult.error);
  }

  event.completed();
}
