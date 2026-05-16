/* global document, console, HTMLInputElement */

const EMPTY_SELECTION_MESSAGE = "(no text selected)";
const EMPTY_REPLACEMENT_MESSAGE = "(enter replacement text first)";

export interface WordOfficeService {
  getSelectionText(): Promise<string>;
  replaceSelectionText(newText: string): Promise<void>;
}

function setSelectedTextDisplay(message: string): void {
  const el = document.getElementById("selected-text-display");
  if (el) {
    el.textContent = message;
  }
}

function showAppBody(): void {
  const sideloadMsg = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");
  if (sideloadMsg) {
    sideloadMsg.style.display = "none";
  }
  if (appBody) {
    appBody.style.display = "flex";
  }
}

export function initApp(officeService: WordOfficeService): void {
  showAppBody();

  const getSelectionBtn = document.getElementById("get-selection");
  const replaceSelectionBtn = document.getElementById("replace-selection");
  const replacementInput = document.getElementById("replacement-input");

  if (getSelectionBtn) {
    getSelectionBtn.onclick = async () => {
      try {
        const text = await officeService.getSelectionText();
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
        await officeService.replaceSelectionText(newText);
        const updated = await officeService.getSelectionText();
        setSelectedTextDisplay(updated.length > 0 ? updated : EMPTY_SELECTION_MESSAGE);
      } catch (error) {
        console.error("replaceSelectionText failed:", error);
      }
    };
  }
}
