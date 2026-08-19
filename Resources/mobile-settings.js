/* global Office, msal, SignaturePreferences */

const CONFIG = {
  clientId: "asd",
  tenantId: "https://login.microsoftonline.com/asd",
  officeNumber: "YOUR_FIXED_OFFICE_NUMBER",
};
const AUTO_RENDER_DATA_KEY = "attensam.signature.render-data.v1";

const phoneModeSelect = document.getElementById("phone-mode");
const edvHotlineOption = document.getElementById("edv-hotline-option");
const greetingModeSelect = document.getElementById("greeting-mode");
const autoInsertCheckbox = document.getElementById("auto-insert");
const autoInsertModeField = document.getElementById("auto-insert-mode-field");
const autoInsertModeSelect = document.getElementById("auto-insert-mode");
const settingsStatus = document.getElementById("settings-status");
const closeButton = document.getElementById("close-button");

let currentSettings;
let currentProfile;
let signatureTemplate = "";
let msalInstance;

function setSettingsStatus(message) {
  settingsStatus.textContent = message;
}

function setControlsDisabled(disabled) {
  phoneModeSelect.disabled = disabled;
  greetingModeSelect.disabled = disabled;
  autoInsertCheckbox.disabled = disabled;
  autoInsertModeSelect.disabled = disabled;
}

function updateAutoInsertVisibility() {
  autoInsertModeField.hidden = !autoInsertCheckbox.checked;
}

function updateDepartmentOption(department) {
  const canUseEdvHotline = String(department || "").trim().toLocaleUpperCase("de-AT") === "IT";
  edvHotlineOption.hidden = !canUseEdvHotline;
  edvHotlineOption.disabled = !canUseEdvHotline;
  if (!canUseEdvHotline && phoneModeSelect.value === "EDVHotline") {
    phoneModeSelect.value = "Alles";
  }
  return canUseEdvHotline;
}

function showSettings(settings, department) {
  const canUseEdvHotline = updateDepartmentOption(department);
  phoneModeSelect.value = settings.Nummer === "EDVHotline" && !canUseEdvHotline
    ? "Alles"
    : settings.Nummer;
  greetingModeSelect.value = settings.MfG;
  autoInsertCheckbox.checked = settings.AutoInsert;
  autoInsertModeSelect.value = settings.AutoInsertMode;
  updateAutoInsertVisibility();
}

async function acquireGraphToken() {
  if (CONFIG.clientId.includes("YOUR_")) {
    throw new Error("Bitte die Entra Client-ID in mobile-settings.js eintragen.");
  }
  if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
    throw new Error("Dieser Outlook-Client unterstützt Nested App Authentication 1.1 nicht.");
  }
  if (!msalInstance) {
    const authority = CONFIG.tenantId.startsWith("https://")
      ? CONFIG.tenantId
      : `https://login.microsoftonline.com/${CONFIG.tenantId}`;
    msalInstance = await msal.createNestablePublicClientApplication({
      auth: { clientId: CONFIG.clientId, authority },
      cache: { cacheLocation: "localStorage" },
    });
  }
  const request = { scopes: ["User.Read"] };
  try {
    return (await msalInstance.acquireTokenSilent(request)).accessToken;
  } catch (error) {
    if (!(error instanceof msal.InteractionRequiredAuthError)) throw error;
    return (await msalInstance.acquireTokenPopup(request)).accessToken;
  }
}

async function loadProfile() {
  const token = await acquireGraphToken();
  const select = [
    "givenName", "surname", "displayName", "mail", "userPrincipalName",
    "companyName", "city", "streetAddress", "postalCode", "jobTitle",
    "department", "mobilePhone", "businessPhones", "onPremisesExtensionAttributes",
  ].join(",");
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me?$select=${encodeURIComponent(select)}`,
    { headers: { Authorization: `Bearer ${token}` } },
  );
  if (!response.ok) throw new Error(`Microsoft Graph: ${response.status}`);
  const user = await response.json();
  return {
    firstName: user.givenName || "",
    lastName: user.surname || "",
    email: user.mail || user.userPrincipalName || "",
    company: user.companyName || "",
    city: user.city || "",
    street: user.streetAddress || "",
    postalCode: user.postalCode || "",
    jobTitle: user.jobTitle || "",
    department: user.department || "",
    mobile: user.mobilePhone || "",
    phone: user.businessPhones?.[0] || "",
    customAttribute10: user.onPremisesExtensionAttributes?.extensionAttribute10 || "",
    customAttribute11: user.onPremisesExtensionAttributes?.extensionAttribute11 || "",
  };
}

async function saveAutomaticRenderData() {
  const roamingSettings = Office.context.roamingSettings;
  if (!roamingSettings) throw new Error("Outlook RoamingSettings ist nicht verfügbar.");
  const now = new Date().toISOString();
  roamingSettings.set(AUTO_RENDER_DATA_KEY, {
    profile: { ...currentProfile },
    template: signatureTemplate,
    officeNumber: CONFIG.officeNumber,
    settings: { ...currentSettings },
    settingsUpdatedAt: now,
    updatedAt: now,
  });
  await new Promise((resolve, reject) => {
    roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error?.message || "Signaturdaten konnten nicht gespeichert werden."));
    });
  });
}

async function initializeSettings() {
  setControlsDisabled(true);
  try {
    currentSettings = await SignaturePreferences.getSettings();
    showSettings(currentSettings, SignaturePreferences.getDepartment());
    setSettingsStatus("Microsoft-365-Profil wird geladen …");
    [signatureTemplate, currentProfile] = await Promise.all([
      fetch("template.html", { cache: "no-store" }).then((response) => {
        if (!response.ok) throw new Error("template.html konnte nicht geladen werden.");
        return response.text();
      }),
      loadProfile(),
    ]);
    SignaturePreferences.setDepartment(currentProfile.department);
    showSettings(currentSettings, currentProfile.department);
    await saveAutomaticRenderData();
    setControlsDisabled(false);
    setSettingsStatus("Profil und Einstellungen sind bereit.");
  } catch (error) {
    setSettingsStatus(error.message || "Profil und Einstellungen konnten nicht geladen werden.");
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
autoInsertCheckbox.addEventListener("change", () => {
  updateAutoInsertVisibility();
  saveSettings();
});
autoInsertModeSelect.addEventListener("change", saveSettings);
closeButton.addEventListener("click", () => {
  if (Office.context.ui?.closeContainer) Office.context.ui.closeContainer();
  else window.history.back();
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) initializeSettings();
  else setSettingsStatus("Diese Seite muss als Outlook-Add-In geöffnet werden.");
});
