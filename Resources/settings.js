/* global Office, SignaturePreferences, AttensamSignatureRuntime */

const AUTO_RENDER_DATA_KEY = "attensam.signature.render-data.v1";
const SETTINGS_SIGNATURE_MARKER_ID = "attensam-signature-root";

const phoneModeSelect = document.getElementById("phone-mode");
const edvHotlineOption = document.getElementById("edv-hotline-option");
const greetingModeSelect = document.getElementById("greeting-mode");
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

function setControlsDisabled(disabled) {
  phoneModeSelect.disabled = disabled;
  greetingModeSelect.disabled = disabled;
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

function hasAttensamSignature(bodyHtml) {
  return bodyHtml.includes(`id="${SETTINGS_SIGNATURE_MARKER_ID}"`)
    || bodyHtml.includes(`id='${SETTINGS_SIGNATURE_MARKER_ID}'`)
    || bodyHtml.includes('data-attensam-signature="v1"')
    || bodyHtml.includes("data-attensam-signature='v1'");
}

async function updateInsertedSignature() {
  const body = Office.context.mailbox.item?.body;
  if (!body?.getAsync || !body?.setSignatureAsync) return false;
  const bodyHtml = await getBodyHtml(body);
  if (!hasAttensamSignature(bodyHtml)) return false;
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
  await setCurrentSignature(body, html);
  return true;
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
greetingModeSelect.addEventListener("change", saveSettings);
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
