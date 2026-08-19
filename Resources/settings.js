/* global Office, SignaturePreferences */

const phoneModeSelect = document.getElementById("phone-mode");
const edvHotlineOption = document.getElementById("edv-hotline-option");
const greetingModeSelect = document.getElementById("greeting-mode");
const titleBeforeField = document.getElementById("title-before-field");
const insertTitleBeforeCheckbox = document.getElementById("insert-title-before");
const titleAfterField = document.getElementById("title-after-field");
const insertTitleAfterCheckbox = document.getElementById("insert-title-after");
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
  autoInsertCheckbox.disabled = disabled;
  autoInsertModeSelect.disabled = disabled;
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
    });
    setSettingsStatus("Einstellungen gespeichert.");
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
autoInsertCheckbox.addEventListener("change", () => {
  updateAutoInsertVisibility();
  saveSettings();
});
autoInsertModeSelect.addEventListener("change", saveSettings);

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) initializeSettings();
  else setSettingsStatus("Diese Seite muss als Outlook-Add-In geöffnet werden.");
});
