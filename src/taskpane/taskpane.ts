/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, console, HTMLInputElement */

const EMPTY_SELECTION_MESSAGE = "(no text selected)";
const EMPTY_REPLACEMENT_MESSAGE = "(enter replacement text first)";

async function getSelectionText(): Promise<string> {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    return range.text;
  });
}

async function replaceSelectionText(newText: string): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(newText, Word.InsertLocation.replace);
    await context.sync();
  });
}

function setSelectedTextDisplay(message: string): void {
  const el = document.getElementById("selected-text-display");
  if (el) {
    el.textContent = message;
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    return;
  }

  const sideloadMsg = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");
  if (sideloadMsg) {
    sideloadMsg.style.display = "none";
  }
  if (appBody) {
    appBody.style.display = "flex";
  }

  const getSelectionBtn = document.getElementById("get-selection");
  const replaceSelectionBtn = document.getElementById("replace-selection");
  const replacementInput = document.getElementById("replacement-input");

  if (getSelectionBtn) {
    getSelectionBtn.onclick = async () => {
      try {
        const text = await getSelectionText();
        setSelectedTextDisplay(text.length > 0 ? text : EMPTY_SELECTION_MESSAGE);
      } catch (error) {
        console.error("getSelectionText failed:", error);
        setSelectedTextDisplay(EMPTY_SELECTION_MESSAGE);
      }
    };
  }

  if (replaceSelectionBtn && replacementInput instanceof HTMLInputElement) {
    replaceSelectionBtn.onclick = async () => {
      const newText = replacementInput.value;
      if (newText.trim().length === 0) {
        console.warn("Replace skipped: empty replacement input");
        setSelectedTextDisplay(EMPTY_REPLACEMENT_MESSAGE);
        return;
      }

      try {
        await replaceSelectionText(newText);
        const updated = await getSelectionText();
        setSelectedTextDisplay(updated.length > 0 ? updated : EMPTY_SELECTION_MESSAGE);
      } catch (error) {
        console.error("replaceSelectionText failed:", error);
      }
    };
  }
});
