/* global Office, SignaturePreferences, AttensamSignatureRuntime, DOMParser */

const AUTO_RENDER_DATA_KEY = "attensam.signature.render-data.v1";
const SETTINGS_SIGNATURE_MARKER_ID = "attensam-signature-root";
const SETTINGS_SIGNATURE_MARKER_TEXT = "ATTENSAM-SIGNATURE-V2";

const phoneModeSelect = document.getElementById("phone-mode");
const edvHotlineOption = document.getElementById("edv-hotline-option");
const greetingModeSelect = document.getElementById("greeting-mode");
const customGreetingField = document.getElementById("custom-greeting-field");
const customGreetingInput = document.getElementById("custom-greeting");
const titleBeforeField = document.getElementById("title-before-field");
const insertTitleBeforeCheckbox = document.getElementById("insert-title-before");
const titleAfterField = document.getElementById("title-after-field");
const insertTitleAfterCheckbox = document.getElementById("insert-title-after");
const mobileUsageCheckbox = document.getElementById("mobile-usage");
const confidentialityCheckbox = document.getElementById("confidentiality");
const autoInsertCheckbox = document.getElementById("auto-insert");
const autoInsertModeField = document.getElementById("auto-insert-mode-field");
const autoInsertModeSelect = document.getElementById("auto-insert-mode");
const settingsStatus = document.getElementById("settings-status");
let currentSettings;

function setSettingsStatus(message) {
  settingsStatus.textContent = message;
}

function updateAutoInsertVisibility() {
  autoInsertModeField.hidden = !autoInsertCheckbox.checked;
}

function updateGreetingVisibility() {
  const isCustom = greetingModeSelect.value === "MfGCustom";
  customGreetingField.hidden = !isCustom;
  customGreetingInput.disabled = greetingModeSelect.disabled || !isCustom;
}

function setControlsDisabled(disabled) {
  phoneModeSelect.disabled = disabled;
  greetingModeSelect.disabled = disabled;
  customGreetingInput.disabled = disabled || greetingModeSelect.value !== "MfGCustom";
  insertTitleBeforeCheckbox.disabled = disabled;
  insertTitleAfterCheckbox.disabled = disabled;
  mobileUsageCheckbox.disabled = disabled;
  confidentialityCheckbox.disabled = disabled;
  autoInsertCheckbox.disabled = disabled;
  autoInsertModeSelect.disabled = disabled;
}

function getBodyHtml(body) {
  return new Promise((resolve, reject) => {
    body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(String(result.value || ""));
      else reject(new Error(result.error?.message || "Nachrichtentext konnte nicht gelesen werden."));
    });
  });
}

function setCurrentSignature(body, html) {
  return new Promise((resolve, reject) => {
    body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error?.message || "Signatur konnte nicht aktualisiert werden."));
      },
    );
  });
}

function setBodyHtml(body, html) {
  return new Promise((resolve, reject) => {
    body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error?.message || "Nachrichtentext konnte nicht aktualisiert werden."));
      },
    );
  });
}

function replaceMarkedSignature(bodyHtml, signatureHtml) {
  const document = new DOMParser().parseFromString(bodyHtml, "text/html");
  const hiddenMarker = Array.from(document.querySelectorAll("span"))
    .find((element) => element.textContent?.includes(SETTINGS_SIGNATURE_MARKER_TEXT));
  const existingSignature = document.getElementById(SETTINGS_SIGNATURE_MARKER_ID)
    || document.querySelector('[data-attensam-signature="v1"]')
    || document.querySelector('[data-attensam-signature="v2"]')
    || hiddenMarker?.parentElement;
  if (!existingSignature) return null;
  existingSignature.outerHTML = signatureHtml;
  return document.body.innerHTML;
}

async function updateInsertedSignature() {
  const body = Office.context.mailbox.item?.body;
  if (!body?.getAsync) return false;
  const renderData = Office.context.roamingSettings?.get(AUTO_RENDER_DATA_KEY);
  if (!renderData?.profile || typeof renderData.template !== "string") return false;
  const delegation = await new Promise((resolve) => {
    AttensamSignatureRuntime.resolveDelegation(renderData, resolve);
  });
  const html = AttensamSignatureRuntime.renderSignature(
    renderData,
    currentSettings,
    delegation,
  );
  const bodyHtml = await getBodyHtml(body);
  if (body.setSignatureAsync) {
    await setCurrentSignature(body, html);
    return true;
  }
  const replacedBodyHtml = replaceMarkedSignature(bodyHtml, html);
  if (replacedBodyHtml !== null && body.setAsync) {
    await setBodyHtml(body, replacedBodyHtml);
    return true;
  }
  return false;
}

async function initializeSettings() {
  try {
    currentSettings = await SignaturePreferences.getSettings();
    const department = SignaturePreferences.getDepartment();
    const titleAttributes = SignaturePreferences.getTitleAttributes();
    const canUseEdvHotline = department.trim().toLocaleUpperCase("de-AT") === "IT";
    edvHotlineOption.hidden = !canUseEdvHotline;
    edvHotlineOption.disabled = !canUseEdvHotline;
    phoneModeSelect.value = currentSettings.Nummer === "EDVHotline" && !canUseEdvHotline
      ? "Alles"
      : currentSettings.Nummer;
    greetingModeSelect.value = currentSettings.MfG;
    customGreetingInput.value = currentSettings.CustomGreeting;
    updateGreetingVisibility();
    titleBeforeField.hidden = !titleAttributes.customAttribute10;
    titleAfterField.hidden = !titleAttributes.customAttribute11;
    insertTitleBeforeCheckbox.checked = currentSettings.InsertTitleBefore;
    insertTitleAfterCheckbox.checked = currentSettings.InsertTitleAfter;
    mobileUsageCheckbox.checked = currentSettings.MobileUsage;
    confidentialityCheckbox.checked = currentSettings.Confidentiality;
    autoInsertCheckbox.checked = currentSettings.AutoInsert;
    autoInsertModeSelect.value = currentSettings.AutoInsertMode;
    updateAutoInsertVisibility();
    setControlsDisabled(false);
    setSettingsStatus("Einstellungen geladen.");
  } catch (error) {
    setSettingsStatus(error.message || "Einstellungen konnten nicht geladen werden.");
  }
}

async function saveSettings() {
  setControlsDisabled(true);
  setSettingsStatus("Einstellungen werden gespeichert …");
  try {
    currentSettings = await SignaturePreferences.saveSettings({
      Nummer: phoneModeSelect.value,
      MfG: greetingModeSelect.value,
      CustomGreeting: customGreetingInput.value,
      AutoInsert: autoInsertCheckbox.checked,
      AutoInsertMode: autoInsertModeSelect.value,
      InsertTitleBefore: insertTitleBeforeCheckbox.checked,
      InsertTitleAfter: insertTitleAfterCheckbox.checked,
      MobileUsage: mobileUsageCheckbox.checked,
      Confidentiality: confidentialityCheckbox.checked,
    });
    try {
      const signatureUpdated = await updateInsertedSignature();
      setSettingsStatus(signatureUpdated
        ? "Einstellungen gespeichert und Signatur aktualisiert."
        : "Einstellungen gespeichert.");
    } catch (updateError) {
      console.error("Die Einstellungen wurden gespeichert, aber die Signatur konnte nicht aktualisiert werden.", updateError);
      setSettingsStatus("Einstellungen gespeichert; Signatur konnte nicht aktualisiert werden.");
    }
  } catch (error) {
    setSettingsStatus(error.message || "Einstellungen konnten nicht gespeichert werden.");
  } finally {
    setControlsDisabled(false);
  }
}

phoneModeSelect.addEventListener("change", saveSettings);
greetingModeSelect.addEventListener("change", () => {
  updateGreetingVisibility();
  saveSettings();
});
customGreetingInput.addEventListener("change", saveSettings);
insertTitleBeforeCheckbox.addEventListener("change", saveSettings);
insertTitleAfterCheckbox.addEventListener("change", saveSettings);
mobileUsageCheckbox.addEventListener("change", saveSettings);
confidentialityCheckbox.addEventListener("change", saveSettings);
autoInsertCheckbox.addEventListener("change", () => {
  updateAutoInsertVisibility();
  saveSettings();
});
autoInsertModeSelect.addEventListener("change", saveSettings);

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) initializeSettings();
  else setSettingsStatus("Diese Seite muss als Outlook-Add-In geöffnet werden.");
});
