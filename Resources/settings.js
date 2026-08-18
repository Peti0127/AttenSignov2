/* global Office, SignaturePreferences */

const phoneModeSelect = document.getElementById("phone-mode");
const greetingModeSelect = document.getElementById("greeting-mode");
const settingsStatus = document.getElementById("settings-status");
let currentSettings;

function setSettingsStatus(message) {
  settingsStatus.textContent = message;
}

async function initializeSettings() {
  try {
    currentSettings = await SignaturePreferences.getSettings();
    phoneModeSelect.value = currentSettings.Nummer;
    greetingModeSelect.value = currentSettings.MfG;
    phoneModeSelect.disabled = false;
    greetingModeSelect.disabled = false;
    setSettingsStatus("Einstellungen geladen.");
  } catch (error) {
    setSettingsStatus(error.message || "Einstellungen konnten nicht geladen werden.");
  }
}

async function saveSettings() {
  phoneModeSelect.disabled = true;
  greetingModeSelect.disabled = true;
  setSettingsStatus("Einstellungen werden gespeichert …");
  try {
    currentSettings = await SignaturePreferences.saveSettings({
      Nummer: phoneModeSelect.value,
      MfG: greetingModeSelect.value,
    });
    setSettingsStatus("Einstellungen gespeichert.");
  } catch (error) {
    setSettingsStatus(error.message || "Einstellungen konnten nicht gespeichert werden.");
  } finally {
    phoneModeSelect.disabled = false;
    greetingModeSelect.disabled = false;
  }
}

phoneModeSelect.addEventListener("change", saveSettings);
greetingModeSelect.addEventListener("change", saveSettings);

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) initializeSettings();
  else setSettingsStatus("Diese Seite muss als Outlook-Add-In geöffnet werden.");
});
