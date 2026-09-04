/* Attensam compact UI bundle. Generated from the tested UI modules. */
const ATTENSAM_CONFIG = Object.freeze({
  clientId: "89659501-37e7-4916-abeb-4dc5178e3034",
  tenantId: "https://login.microsoftonline.com/1333c2c2-fdf6-4fdc-8559-3dc12559d264",
  officeNumber: "05 7999 100",
  feedbackEmail: "pnov@attensam.at",
});

function hasConfiguredEntraApp() {
  const clientId = String(ATTENSAM_CONFIG.clientId || "").trim();
  const tenantId = String(ATTENSAM_CONFIG.tenantId || "").trim();
  return Boolean(
    clientId
    && tenantId
    && clientId !== "asd"
    && !clientId.includes("YOUR_")
    && !tenantId.includes("YOUR_")
    && !tenantId.endsWith("/asd")
  );
}

function readableError(error) {
  if (typeof error === "string" && error.trim()) return error.trim();
  if (typeof error?.message === "string" && error.message.trim()) return error.message.trim();
  if (typeof error?.errorMessage === "string" && error.errorMessage.trim()) return error.errorMessage.trim();
  if (typeof error?.errorCode === "string" && error.errorCode.trim()) return error.errorCode.trim();
  return "Unbekannter Fehler";
}

(function followClientTheme() {
  const root = document.documentElement;
  const media = window.matchMedia?.("(prefers-color-scheme: dark)");
  let officeThemeDetected = false;
  const exactColorVariables = ["--paper", "--surface", "--control-bg", "--control-ink", "--footer-bg"];

  function clearExactColors() {
    exactColorVariables.forEach((name) => root.style.removeProperty(name));
  }

  function isDarkColor(value) {
    const hex = String(value || "").trim().replace(/^#/, "");
    if (!/^[0-9a-f]{6}$/i.test(hex)) return false;
    const red = Number.parseInt(hex.slice(0, 2), 16);
    const green = Number.parseInt(hex.slice(2, 4), 16);
    const blue = Number.parseInt(hex.slice(4, 6), 16);
    return (red * 0.299 + green * 0.587 + blue * 0.114) < 128;
  }

  function applySystemPreference() {
    if (officeThemeDetected) return;
    clearExactColors();
    root.dataset.clientTheme = media?.matches ? "dark" : "light";
  }

  function applyOfficeTheme(theme) {
    if (!theme?.bodyBackgroundColor) {
      officeThemeDetected = false;
      applySystemPreference();
      return;
    }
    officeThemeDetected = true;
    const dark = isDarkColor(theme.bodyBackgroundColor);
    root.dataset.clientTheme = dark ? "dark" : "light";
    clearExactColors();
    if (!dark) return;
    const bodyBackground = theme.bodyBackgroundColor;
    const bodyForeground = theme.bodyForegroundColor || "#f2f2f2";
    const controlBackground = theme.controlBackgroundColor || "#222222";
    const controlForeground = theme.controlForegroundColor || bodyForeground;
    root.style.setProperty("--paper", bodyBackground);
    root.style.setProperty("--surface", controlBackground);
    root.style.setProperty("--control-bg", controlBackground);
    root.style.setProperty("--control-ink", controlForeground);
    root.style.setProperty("--footer-bg", controlBackground);
  }

  applySystemPreference();
  if (media?.addEventListener) media.addEventListener("change", applySystemPreference);
  else if (media?.addListener) media.addListener(applySystemPreference);

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) return;
    applyOfficeTheme(Office.context.officeTheme);
    const supportsThemeEvents = Office.context.requirements?.isSetSupported?.("Mailbox", "1.14");
    if (!supportsThemeEvents || !Office.EventType.OfficeThemeChanged || !Office.context.mailbox?.addHandlerAsync) return;
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.OfficeThemeChanged,
      (event) => applyOfficeTheme(event.officeTheme),
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn("Outlook-Designänderungen konnten nicht automatisch überwacht werden.", result.error);
        }
      },
    );
  });
})();
/* global Office */

(function exposeSignaturePreferences(global) {
  const ROAMING_KEY = "attensam.signature.settings.v2";
  const RENDER_DATA_KEY = "attensam.signature.render-data.v1";
  const CUSTOM_SIGNATURES_KEY = "attensam.signature.custom-signatures.v1";
  const CUSTOM_SIGNATURES_CACHE_PREFIX = "attensam.signature.custom-signatures.v1";
  const REQUIRED_ROLE = "ATS.Signature";
  const VIP_ROLE = "ATS.Signature.VIP";
  const CITY_CHANGE_ROLE = "CityChange";
  const MAX_CUSTOM_SIGNATURES = 3;
  const PROFILE_CACHE_MAX_AGE_MS = 14 * 24 * 60 * 60 * 1000;
  const CACHE_PREFIX = "attensam.signature.settings.v2";
  const DEPARTMENT_CACHE_PREFIX = "attensam.signature.department.v1";
  const TITLE_ATTRIBUTES_CACHE_PREFIX = "attensam.signature.title-attributes.v1";
  const ACCESS_CACHE_PREFIX = "attensam.signature.access-role.v1";
  const VIP_CACHE_PREFIX = "attensam.signature.vip-role.v1";
  const CITY_CHANGE_CACHE_PREFIX = "attensam.signature.city-change-role.v1";
  const LEGACY_PHONE_PREFIX = "attensam.signature.phone-mode";
  const DEFAULT_SETTINGS = Object.freeze({
    Nummer: "Alles",
    MfG: "MfG1",
    CustomGreeting: "",
    GreetingLines: 1,
    CityOverride: "Standard",
    AutoInsert: true,
    AutoInsertMode: "NewMail",
    AutoInsertReplies: false,
    AutoInsertForwards: false,
    AutoInsertMeetings: false,
    SkipInternalOnly: false,
    SkipInternalOnNewMail: false,
    InsertTitleBefore: false,
    InsertTitleAfter: false,
    MobileUsage: false,
    MobileUsageText: "",
    Confidentiality: false,
  });
  const ALLOWED_NUMBERS = new Set(["Alles", "Handy", "Festnetz", "Office", "EDVHotline"]);
  const ALLOWED_GREETINGS = new Set(["MfG0", "MfG1", "MfG2", "MfG3", "MfGCustom"]);
  const ALLOWED_GREETING_LINES = new Set([1, 2, 3]);
  const ALLOWED_CITY_OVERRIDES = new Set(["Standard", "Neusiedl am See", "Oberwart", "Wr. Neustadt"]);
  const LEGACY_NUMBER_MAP = Object.freeze({
    both: "Alles",
    mobile: "Handy",
    landline: "Festnetz",
    office: "Office",
  });

  function currentUserKey() {
    try {
      const email = Office.context.mailbox.userProfile.emailAddress;
      return String(email || "unknown-user").trim().toLocaleLowerCase("de-AT");
    } catch {
      return "unknown-user";
    }
  }

  function getValidRenderData() {
    const cached = Office.context.roamingSettings?.get(RENDER_DATA_KEY);
    if (!cached?.profile || typeof cached.template !== "string" || !cached.template.trim()) return null;
    const cachedAt = Date.parse(cached.profileUpdatedAt || cached.updatedAt || "");
    if (!Number.isFinite(cachedAt) || Date.now() - cachedAt > PROFILE_CACHE_MAX_AGE_MS) return null;
    const mailboxEmail = currentUserKey();
    const cachedEmails = [cached.mailboxEmail, cached.profile.email]
      .map((value) => String(value || "").trim().toLocaleLowerCase("de-AT"))
      .filter(Boolean);
    return mailboxEmail && cachedEmails.includes(mailboxEmail) ? cached : null;
  }

  function storageKey() {
    return `${CACHE_PREFIX}:${currentUserKey()}`;
  }

  function departmentStorageKey() {
    return `${DEPARTMENT_CACHE_PREFIX}:${currentUserKey()}`;
  }

  function titleAttributesStorageKey() {
    return `${TITLE_ATTRIBUTES_CACHE_PREFIX}:${currentUserKey()}`;
  }

  function customSignaturesStorageKey() {
    return `${CUSTOM_SIGNATURES_CACHE_PREFIX}:${currentUserKey()}`;
  }

  function vipStorageKey() {
    return `${VIP_CACHE_PREFIX}:${currentUserKey()}`;
  }

  function accessStorageKey() {
    return `${ACCESS_CACHE_PREFIX}:${currentUserKey()}`;
  }

  function setAccessAuthorized(value) {
    localStorage.setItem(accessStorageKey(), JSON.stringify({
      authorized: value === true,
      role: REQUIRED_ROLE,
      updatedAt: new Date().toISOString(),
    }));
  }

  function getAccessAuthorizationState() {
    try {
      const value = JSON.parse(localStorage.getItem(accessStorageKey()) || "null");
      return typeof value?.authorized === "boolean" ? value.authorized : null;
    } catch {
      localStorage.removeItem(accessStorageKey());
      return null;
    }
  }

  function getAccessAuthorized() {
    return getAccessAuthorizationState() === true;
  }

  function cityChangeStorageKey() {
    return `${CITY_CHANGE_CACHE_PREFIX}:${currentUserKey()}`;
  }

  function setVipAuthorized(value) {
    localStorage.setItem(vipStorageKey(), JSON.stringify({
      authorized: value === true,
      updatedAt: new Date().toISOString(),
    }));
  }

  function getVipAuthorizationState() {
    try {
      const value = JSON.parse(localStorage.getItem(vipStorageKey()) || "null");
      return typeof value?.authorized === "boolean" ? value.authorized : null;
    } catch {
      localStorage.removeItem(vipStorageKey());
      return null;
    }
  }

  function getVipAuthorized() {
    return getVipAuthorizationState() === true;
  }

  function setCityChangeAuthorized(value) {
    localStorage.setItem(cityChangeStorageKey(), JSON.stringify({
      authorized: value === true,
      role: CITY_CHANGE_ROLE,
      updatedAt: new Date().toISOString(),
    }));
  }

  function getCityChangeAuthorizationState() {
    try {
      const value = JSON.parse(localStorage.getItem(cityChangeStorageKey()) || "null");
      return typeof value?.authorized === "boolean" ? value.authorized : null;
    } catch {
      localStorage.removeItem(cityChangeStorageKey());
      return null;
    }
  }

  function getCityChangeAuthorized() {
    return getCityChangeAuthorizationState() === true;
  }

  function normalizeCustomSignatures(value) {
    const items = Array.isArray(value?.items) ? value.items.slice(0, MAX_CUSTOM_SIGNATURES) : [];
    const normalizedItems = items.map((item) => ({
      id: String(item?.id || "").replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 64),
      title: String(item?.title || "").replace(/\s+/g, " ").trim().slice(0, 80),
      html: String(item?.html || "").trim().slice(0, 7000),
      settings: item?.settings && typeof item.settings === "object"
        ? publicSettings(normalizeRecord(item.settings) || normalizeRecord(DEFAULT_SETTINGS))
        : null,
      updatedAt: Number.isFinite(Date.parse(item?.updatedAt)) ? item.updatedAt : "",
    })).filter((item) => item.id && item.title && item.html);
    const allowedIds = new Set(normalizedItems.map((item) => item.id));
    const requestedDefault = String(value?.defaultId || "standard");
    return {
      requiredRole: VIP_ROLE,
      defaultId: allowedIds.has(requestedDefault) ? requestedDefault : "standard",
      items: normalizedItems,
      updatedAt: Number.isFinite(Date.parse(value?.updatedAt)) ? value.updatedAt : "",
    };
  }

  async function getCustomSignatures() {
    let cached = null;
    try {
      cached = normalizeCustomSignatures(JSON.parse(localStorage.getItem(customSignaturesStorageKey()) || "null"));
    } catch {
      localStorage.removeItem(customSignaturesStorageKey());
    }
    const roaming = normalizeCustomSignatures(Office.context.roamingSettings?.get(CUSTOM_SIGNATURES_KEY));
    const newest = recordTime(cached) > recordTime(roaming) ? cached : roaming;
    localStorage.setItem(customSignaturesStorageKey(), JSON.stringify(newest));
    return newest;
  }

  async function saveCustomSignatures(value) {
    if (Array.isArray(value?.items) && value.items.length > MAX_CUSTOM_SIGNATURES) {
      throw new Error("Maximal drei Signaturen sind erlaubt.");
    }
    const record = normalizeCustomSignatures({ ...value, updatedAt: new Date().toISOString() });
    const bytes = new TextEncoder().encode(JSON.stringify(record)).length;
    if (bytes > 23000) throw new Error("Die benutzerdefinierten Signaturen sind für Outlook RoamingSettings zu groß.");
    const roamingSettings = Office.context.roamingSettings;
    if (!roamingSettings) throw new Error("Outlook RoamingSettings ist nicht verfügbar.");
    roamingSettings.set(CUSTOM_SIGNATURES_KEY, record);
    await new Promise((resolve, reject) => {
      roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error?.message || "Signaturen konnten nicht gespeichert werden."));
      });
    });
    localStorage.setItem(customSignaturesStorageKey(), JSON.stringify(record));
    return record;
  }

  function normalizeCustomGreeting(value) {
    return String(value || "").replace(/\s+/g, " ").trim().slice(0, 200);
  }

  function normalizeMobileUsageText(value) {
    return String(value || "").replace(/\s+/g, " ").trim().slice(0, 300);
  }

  function setDepartment(department) {
    localStorage.setItem(departmentStorageKey(), String(department || "").trim());
  }

  function getDepartment() {
    const cached = localStorage.getItem(departmentStorageKey());
    if (cached) return cached;
    const renderData = Office.context.roamingSettings?.get(RENDER_DATA_KEY);
    const department = String(renderData?.profile?.department || "").trim();
    if (department) setDepartment(department);
    return department;
  }

  function setTitleAttributes(customAttribute10, customAttribute11) {
    localStorage.setItem(titleAttributesStorageKey(), JSON.stringify({
      customAttribute10: String(customAttribute10 || "").trim(),
      customAttribute11: String(customAttribute11 || "").trim(),
    }));
  }

  function getTitleAttributes() {
    const cached = localStorage.getItem(titleAttributesStorageKey());
    if (cached) {
      try {
        const value = JSON.parse(cached);
        return {
          customAttribute10: String(value?.customAttribute10 || "").trim(),
          customAttribute11: String(value?.customAttribute11 || "").trim(),
        };
      } catch {
        localStorage.removeItem(titleAttributesStorageKey());
      }
    }
    const renderData = Office.context.roamingSettings?.get(RENDER_DATA_KEY);
    const result = {
      customAttribute10: String(renderData?.profile?.customAttribute10 || "").trim(),
      customAttribute11: String(renderData?.profile?.customAttribute11 || "").trim(),
    };
    if (result.customAttribute10 || result.customAttribute11) {
      setTitleAttributes(result.customAttribute10, result.customAttribute11);
    }
    return result;
  }

  function normalizeRecord(value) {
    if (!value || typeof value !== "object") return null;
    const legacyAllMail = value.AutoInsertMode === "AllMail";
    const autoInsertReplies = typeof value.AutoInsertReplies === "boolean"
      ? value.AutoInsertReplies
      : legacyAllMail;
    const autoInsertForwards = typeof value.AutoInsertForwards === "boolean"
      ? value.AutoInsertForwards
      : legacyAllMail;
    return {
      Nummer: ALLOWED_NUMBERS.has(value?.Nummer) ? value.Nummer : DEFAULT_SETTINGS.Nummer,
      MfG: ALLOWED_GREETINGS.has(value?.MfG) ? value.MfG : DEFAULT_SETTINGS.MfG,
      CustomGreeting: normalizeCustomGreeting(value?.CustomGreeting),
      GreetingLines: ALLOWED_GREETING_LINES.has(Number(value?.GreetingLines))
        ? Number(value.GreetingLines)
        : DEFAULT_SETTINGS.GreetingLines,
      CityOverride: ALLOWED_CITY_OVERRIDES.has(value?.CityOverride)
        ? value.CityOverride
        : DEFAULT_SETTINGS.CityOverride,
      AutoInsert: true,
      AutoInsertMode: autoInsertReplies && autoInsertForwards ? "AllMail" : "NewMail",
      AutoInsertReplies: autoInsertReplies,
      AutoInsertForwards: autoInsertForwards,
      AutoInsertMeetings: value.AutoInsertMeetings === true,
      SkipInternalOnly: value.SkipInternalOnly === true || value.InternalRecipientsOnly === true,
      SkipInternalOnNewMail: value.SkipInternalOnNewMail === true,
      InsertTitleBefore: value.InsertTitleBefore === true,
      InsertTitleAfter: value.InsertTitleAfter === true,
      MobileUsage: value.MobileUsage === true,
      MobileUsageText: normalizeMobileUsageText(value?.MobileUsageText),
      Confidentiality: value.Confidentiality === true,
      updatedAt: Number.isFinite(Date.parse(value.updatedAt)) ? value.updatedAt : "",
    };
  }

  function readCachedRecord() {
    const stored = localStorage.getItem(storageKey());
    if (!stored) return null;
    try {
      return normalizeRecord(JSON.parse(stored));
    } catch {
      localStorage.removeItem(storageKey());
      return null;
    }
  }

  function recordTime(record) {
    const timestamp = Date.parse(record?.updatedAt || "");
    return Number.isFinite(timestamp) ? timestamp : 0;
  }

  function publicSettings(record) {
    return {
      Nummer: record.Nummer,
      MfG: record.MfG,
      CustomGreeting: record.CustomGreeting,
      GreetingLines: record.GreetingLines,
      CityOverride: record.CityOverride,
      AutoInsert: record.AutoInsert,
      AutoInsertMode: record.AutoInsertMode,
      AutoInsertReplies: record.AutoInsertReplies,
      AutoInsertForwards: record.AutoInsertForwards,
      AutoInsertMeetings: record.AutoInsertMeetings,
      SkipInternalOnly: record.SkipInternalOnly,
      SkipInternalOnNewMail: record.SkipInternalOnNewMail,
      InsertTitleBefore: record.InsertTitleBefore,
      InsertTitleAfter: record.InsertTitleAfter,
      MobileUsage: record.MobileUsage,
      MobileUsageText: record.MobileUsageText,
      Confidentiality: record.Confidentiality,
    };
  }

  async function getSettings() {
    const cached = readCachedRecord();
    const roaming = normalizeRecord(
      Office.context.roamingSettings?.get(ROAMING_KEY),
    );

    // Outlook can expose an older in-memory roaming snapshot immediately after
    // navigating between pages. Timestamps let the local cache bridge that gap.
    const newest = recordTime(cached) > recordTime(roaming) ? cached : (roaming || cached);
    if (newest) {
      localStorage.setItem(storageKey(), JSON.stringify(newest));
      return publicSettings(newest);
    }

    // Preserve an existing phone choice from the previous add-in version.
    const legacy = localStorage.getItem(`${LEGACY_PHONE_PREFIX}:${currentUserKey()}`);
    return {
      Nummer: LEGACY_NUMBER_MAP[legacy] || DEFAULT_SETTINGS.Nummer,
      MfG: DEFAULT_SETTINGS.MfG,
      CustomGreeting: DEFAULT_SETTINGS.CustomGreeting,
      GreetingLines: DEFAULT_SETTINGS.GreetingLines,
      CityOverride: DEFAULT_SETTINGS.CityOverride,
      AutoInsert: DEFAULT_SETTINGS.AutoInsert,
      AutoInsertMode: DEFAULT_SETTINGS.AutoInsertMode,
      AutoInsertReplies: DEFAULT_SETTINGS.AutoInsertReplies,
      AutoInsertForwards: DEFAULT_SETTINGS.AutoInsertForwards,
      AutoInsertMeetings: DEFAULT_SETTINGS.AutoInsertMeetings,
      SkipInternalOnly: DEFAULT_SETTINGS.SkipInternalOnly,
      SkipInternalOnNewMail: DEFAULT_SETTINGS.SkipInternalOnNewMail,
      InsertTitleBefore: DEFAULT_SETTINGS.InsertTitleBefore,
      InsertTitleAfter: DEFAULT_SETTINGS.InsertTitleAfter,
      MobileUsage: DEFAULT_SETTINGS.MobileUsage,
      MobileUsageText: DEFAULT_SETTINGS.MobileUsageText,
      Confidentiality: DEFAULT_SETTINGS.Confidentiality,
    };
  }

  async function saveSettings(settings) {
    if (
      !ALLOWED_NUMBERS.has(settings?.Nummer)
      || !ALLOWED_GREETINGS.has(settings?.MfG)
      || typeof settings?.CustomGreeting !== "string"
      || !ALLOWED_GREETING_LINES.has(Number(settings?.GreetingLines))
      || !ALLOWED_CITY_OVERRIDES.has(settings?.CityOverride)
      || typeof settings?.AutoInsert !== "boolean"
      || typeof settings?.AutoInsertReplies !== "boolean"
      || typeof settings?.AutoInsertForwards !== "boolean"
      || typeof settings?.AutoInsertMeetings !== "boolean"
      || typeof settings?.SkipInternalOnly !== "boolean"
      || typeof settings?.SkipInternalOnNewMail !== "boolean"
      || typeof settings?.InsertTitleBefore !== "boolean"
      || typeof settings?.InsertTitleAfter !== "boolean"
      || typeof settings?.MobileUsage !== "boolean"
      || typeof settings?.MobileUsageText !== "string"
      || typeof settings?.Confidentiality !== "boolean"
    ) {
      throw new Error("Ungültige Einstellung.");
    }
    const record = {
      Nummer: settings.Nummer,
      MfG: settings.MfG,
      CustomGreeting: normalizeCustomGreeting(settings.CustomGreeting),
      GreetingLines: Number(settings.GreetingLines),
      CityOverride: settings.CityOverride,
      AutoInsert: true,
      AutoInsertMode: settings.AutoInsertReplies && settings.AutoInsertForwards ? "AllMail" : "NewMail",
      AutoInsertReplies: settings.AutoInsertReplies,
      AutoInsertForwards: settings.AutoInsertForwards,
      AutoInsertMeetings: settings.AutoInsertMeetings,
      SkipInternalOnly: settings.SkipInternalOnly,
      SkipInternalOnNewMail: settings.SkipInternalOnNewMail,
      InsertTitleBefore: settings.InsertTitleBefore,
      InsertTitleAfter: settings.InsertTitleAfter,
      MobileUsage: settings.MobileUsage,
      MobileUsageText: normalizeMobileUsageText(settings.MobileUsageText),
      Confidentiality: settings.Confidentiality,
      updatedAt: new Date().toISOString(),
    };
    const roamingSettings = Office.context.roamingSettings;
    if (!roamingSettings) throw new Error("Outlook RoamingSettings ist nicht verfügbar.");

    roamingSettings.set(ROAMING_KEY, record);
    const existingRenderData = roamingSettings.get(RENDER_DATA_KEY);
    if (existingRenderData && typeof existingRenderData === "object") {
      roamingSettings.set(RENDER_DATA_KEY, {
        ...existingRenderData,
        settings: publicSettings(record),
        settingsUpdatedAt: record.updatedAt,
      });
    }
    await new Promise((resolve, reject) => {
      roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error?.message || "Einstellungen konnten nicht gespeichert werden."));
      });
    });

    localStorage.setItem(storageKey(), JSON.stringify(record));
    return publicSettings(record);
  }

  async function getSettingsForSignature(signatureId = "standard") {
    if (!signatureId || signatureId === "standard") return getSettings();
    const record = await getCustomSignatures();
    const item = record.items.find((entry) => entry.id === signatureId);
    if (!item) throw new Error("Die ausgewählte Signatur wurde nicht gefunden.");
    return item.settings || getSettings();
  }

  async function saveSettingsForSignature(signatureId = "standard", settings) {
    if (!signatureId || signatureId === "standard") return saveSettings(settings);
    const normalized = normalizeRecord({ ...settings, updatedAt: new Date().toISOString() });
    if (!normalized) throw new Error("Ungültige Einstellung.");
    const record = await getCustomSignatures();
    const item = record.items.find((entry) => entry.id === signatureId);
    if (!item) throw new Error("Die ausgewählte Signatur wurde nicht gefunden.");
    const savedSettings = publicSettings(normalized);
    await saveCustomSignatures({
      ...record,
      items: record.items.map((entry) => entry.id === signatureId
        ? { ...entry, settings: savedSettings, updatedAt: new Date().toISOString() }
        : entry),
    });
    return savedSettings;
  }

  global.SignaturePreferences = Object.freeze({
    getSettings,
    saveSettings,
    getDepartment,
    setDepartment,
    getTitleAttributes,
    setTitleAttributes,
    getCustomSignatures,
    saveCustomSignatures,
    getSettingsForSignature,
    saveSettingsForSignature,
    getValidRenderData,
    getAccessAuthorized,
    getAccessAuthorizationState,
    setAccessAuthorized,
    getVipAuthorized,
    getVipAuthorizationState,
    setVipAuthorized,
    getCityChangeAuthorized,
    getCityChangeAuthorizationState,
    setCityChangeAuthorized,
  });
})(window);

(function compactRoute(){
  const activeView = new URLSearchParams(window.location.search).get("view");
  if (activeView === "settings" || activeView === "feedback") return;
/* global Office, msal, SignaturePreferences */

const CONFIG = ATTENSAM_CONFIG;
const AUTO_RENDER_DATA_KEY = "attensam.signature.render-data.v1";
const DELEGATED_PROFILE_LOCAL_CACHE_KEY = "attensam.signature.delegated-profiles.v1";
const REQUIRED_ROLE = "ATS.Signature";
const VIP_ROLE = "ATS.Signature.VIP";
const CITY_CHANGE_ROLE = "CityChange";
const EXCLUDED_SUBJECT_PREFIX = "Ihre Objektinformation - ";
const MAX_CUSTOM_SIGNATURES = 3;
const SIGNATURE_MARKER_ID = "attensam-signature-root";
const SIGNATURE_MARKER_TEXT = "Attensam-Signatur";

const profile = {
  id: "", firstName: "", lastName: "", jobTitle: "", company: "",
  email: "", phone: "", mobile: "", street: "",
  postalCode: "", city: "", department: "", customAttribute10: "",
  customAttribute11: "",
};
let signatureTemplate = "";
let msalInstance;
let profileLoaded = false;
let usingCachedProfile = false;
let currentDelegation = null;
let currentDelegationAddress = "";
let userRoles = new Set();
let accessAuthorized = false;
let vipAuthorized = false;
let cityChangeAuthorized = false;
let customSignatures = { requiredRole: VIP_ROLE, defaultId: "standard", items: [] };
let contextSignatureId = "standard";
let editingCustomSignatureId = null;
let deleteConfirmationArmed = false;
let signatureSettings = {
  Nummer: "Alles",
  MfG: "MfG1",
  CustomGreeting: "",
  GreetingLines: 1,
  CityOverride: "Standard",
  InsertTitleBefore: false,
  InsertTitleAfter: false,
  MobileUsage: false,
  MobileUsageText: "",
  Confidentiality: false,
};

const statusElement = document.getElementById("status");
const signatureMain = document.getElementById("signature-main");
const taskpaneAccessDenied = document.getElementById("taskpane-access-denied");
const previewElement = document.getElementById("signature-preview");
const signatureButton = document.getElementById("signature-button");
const profileWarningsElement = document.getElementById("profile-warnings");
const mainSettingsLink = document.getElementById("main-settings-link");
const customAddButton = document.getElementById("custom-add-button");
const customEditor = document.getElementById("custom-editor");
const customTitleInput = document.getElementById("custom-signature-title");
const customHtmlInput = document.getElementById("custom-signature-html");
const customSaveButton = document.getElementById("custom-save-button");
const customCancelButton = document.getElementById("custom-cancel-button");
const customSignaturesElement = document.getElementById("custom-signatures");
const contextMenu = document.getElementById("signature-context-menu");
const openSignatureSettingsButton = document.getElementById("open-signature-settings");
const setDefaultButton = document.getElementById("set-default-signature");
const editCustomButton = document.getElementById("edit-custom-signature");
const deleteCustomButton = document.getElementById("delete-custom-signature");
const feedbackButton = document.getElementById("feedback-button");

const initialCachedVipState = SignaturePreferences.getVipAuthorizationState();
if (initialCachedVipState !== null) {
  vipAuthorized = initialCachedVipState;
}
const initialCachedCityChangeState = SignaturePreferences.getCityChangeAuthorizationState();
cityChangeAuthorized = initialCachedCityChangeState === true;

function setStatus(message) {
  statusElement.textContent = message;
}

function applyAccessView() {
  signatureMain.hidden = !accessAuthorized;
  taskpaneAccessDenied.hidden = accessAuthorized;
  mainSettingsLink.hidden = !accessAuthorized || vipAuthorized;
  feedbackButton.hidden = !accessAuthorized;
  feedbackButton.disabled = !accessAuthorized || !profileLoaded;
  if (!accessAuthorized) setStatus("Sie haben kein Zugriff auf dieses Add-In, bitte EDV kontaktieren!");
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;").replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;").replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function rememberAuthenticationRoles(result) {
  let tokenClaims = {};
  try {
    const encoded = String(result?.idToken || "").split(".")[1];
    if (encoded) {
      const normalized = encoded.replace(/-/g, "+").replace(/_/g, "/");
      tokenClaims = JSON.parse(atob(normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=")));
    }
  } catch {
    tokenClaims = {};
  }
  const roles = [
    ...(Array.isArray(result?.idTokenClaims?.roles) ? result.idTokenClaims.roles : []),
    ...(Array.isArray(result?.account?.idTokenClaims?.roles) ? result.account.idTokenClaims.roles : []),
    ...(Array.isArray(tokenClaims?.roles) ? tokenClaims.roles : []),
  ];
  userRoles = new Set(roles.map((role) => String(role).trim()).filter(Boolean));
  accessAuthorized = userRoles.has(REQUIRED_ROLE);
  vipAuthorized = userRoles.has(VIP_ROLE);
  cityChangeAuthorized = userRoles.has(CITY_CHANGE_ROLE) || userRoles.has("ATS.Signature.CityChange");
  SignaturePreferences.setAccessAuthorized(accessAuthorized);
  SignaturePreferences.setVipAuthorized(vipAuthorized);
  SignaturePreferences.setCityChangeAuthorized(cityChangeAuthorized);
  applyAccessView();
}

function sanitizeCustomSignatureHtml(value) {
  const documentValue = new DOMParser().parseFromString(`<div>${String(value || "")}</div>`, "text/html");
  const root = documentValue.body.firstElementChild;
  const allowedTags = new Set(["A", "B", "BR", "CENTER", "DIV", "EM", "FONT", "HR", "I", "IMG", "LI", "OL", "P", "SPAN", "STRONG", "SUB", "SUP", "TABLE", "TBODY", "TD", "TFOOT", "TH", "THEAD", "TR", "U", "UL"]);
  const allowedAttributes = {
    A: new Set(["href", "style", "title"]), IMG: new Set(["src", "alt", "width", "height", "border", "style"]),
    FONT: new Set(["face", "size", "color", "style"]),
    TABLE: new Set(["cellpadding", "cellspacing", "border", "width", "style"]),
    TD: new Set(["colspan", "rowspan", "width", "height", "style"]), TH: new Set(["colspan", "rowspan", "width", "height", "style"]),
  };
  const commonAttributes = new Set(["style", "title"]);
  [...root.querySelectorAll("*")].forEach((element) => {
    if (!allowedTags.has(element.tagName)) {
      element.replaceWith(...element.childNodes);
      return;
    }
    [...element.attributes].forEach((attribute) => {
      const name = attribute.name.toLowerCase();
      const allowed = allowedAttributes[element.tagName]?.has(name) || commonAttributes.has(name);
      if (!allowed || name.startsWith("on")) element.removeAttribute(attribute.name);
    });
    if (element.hasAttribute("style")) {
      const style = element.getAttribute("style").slice(0, 2000);
      if (/url\s*\(|expression\s*\(|javascript\s*:|@import|behavior\s*:|-moz-binding/i.test(style)) element.removeAttribute("style");
      else element.setAttribute("style", style);
    }
    ["href", "src"].forEach((name) => {
      if (!element.hasAttribute(name)) return;
      const url = element.getAttribute(name).trim();
      const permitted = name === "href" ? /^(https:|mailto:|tel:|#)/i.test(url) : /^https:/i.test(url);
      if (!permitted) element.removeAttribute(name);
    });
  });
  return root.innerHTML.trim();
}

function insertHtmlAtEditorCursor(html) {
  const selection = window.getSelection();
  if (!selection?.rangeCount || !customHtmlInput.contains(selection.anchorNode)) {
    customHtmlInput.insertAdjacentHTML("beforeend", html);
    return;
  }
  const range = selection.getRangeAt(0);
  range.deleteContents();
  const fragment = range.createContextualFragment(html);
  const lastNode = fragment.lastChild;
  range.insertNode(fragment);
  if (lastNode) {
    range.setStartAfter(lastNode);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  }
}

function resetCustomEditor() {
  editingCustomSignatureId = null;
  customEditor.hidden = true;
  customTitleInput.value = "";
  customHtmlInput.innerHTML = "";
  customSaveButton.textContent = "Speichern";
}

function openCustomEditor(item = null) {
  editingCustomSignatureId = item?.id || null;
  customTitleInput.value = item?.title || "";
  customHtmlInput.innerHTML = item?.html || "";
  customSaveButton.textContent = item ? "Änderungen speichern" : "Speichern";
  customEditor.hidden = false;
  customTitleInput.focus();
}

function phoneLine(profileValue = profile, settings = signatureSettings) {
  const phone = escapeHtml(String(profileValue.phone || "").trim());
  const mobile = escapeHtml(String(profileValue.mobile || "").trim());
  const officeNumber = CONFIG.officeNumber.includes("YOUR_")
    ? ""
    : escapeHtml(CONFIG.officeNumber.trim());

  switch (settings.Nummer) {
    case "Handy":
      return mobile ? `Mobil ${mobile}` : "";
    case "Festnetz":
      return phone ? `Tel. ${phone}` : "";
    case "Office":
      return officeNumber ? `Tel. ${officeNumber}` : "";
    case "EDVHotline":
      if (profile.department.trim().toLocaleUpperCase("de-AT") === "IT") {
        return mobile ? `Tel. 05 7999 9999 - Mobil ${mobile}` : "Tel. 05 7999 9999";
      }
      // If the department changed, fall back to the standard phone line.
    default:
      if (phone && mobile) return `Tel. ${phone} - Mobil ${mobile}`;
      if (mobile) return `Mobil ${mobile}`;
      if (phone) return `Tel. ${phone}`;
      return "";
  }
}

function greetingHtml(settings = signatureSettings) {
  if (settings.MfG === "MfG0") return "";
  let greeting = "";
  if (settings.MfG === "MfG1") {
    greeting = "Mit freundlichen Grüßen";
  } else if (settings.MfG === "MfG2") {
    greeting = "Freundliche Grüße";
  } else if (settings.MfG === "MfG3") {
    greeting = "LG";
  } else if (settings.MfG === "MfGCustom") {
    greeting = String(settings.CustomGreeting || "").trim();
  }
  if (!greeting) return "";
  const configuredLines = Number(settings.GreetingLines);
  const blankLines = [1, 2, 3].includes(configuredLines) ? configuredLines : 1;
  return `<p style="margin: 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">${escapeHtml(greeting)}${"<br>".repeat(blankLines + 1)}</p>`;
}

function isOutlookMobile() {
  const platform = Office.context?.platform;
  const platformTypes = Office.PlatformType || {};
  return platform === platformTypes.Android || platform === platformTypes.iOS;
}

function noticesHtml(settings = signatureSettings) {
  let html = "";
  if (settings.MobileUsage && isOutlookMobile()) {
    const customText = String(settings.MobileUsageText || "").replace(/\s+/g, " ").trim();
    const mobileNotice = `Wurde über Outlook Mobile versendet.${customText ? ` ${customText}` : ""}`;
    html += `<p style="margin: 12px 0 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">${escapeHtml(mobileNotice)}</p>`;
  }
  if (settings.Confidentiality) {
    html += '<p style="margin: 6px 0 0; font-family: Aptos, Arial, sans-serif; font-size: 9pt; color: rgb(0, 0, 0);">Diese E-Mail ist vertraulich.</p>';
  }
  return html;
}

function missingProfileFields(profileValue, ignoreMissingTitle = false) {
  const requiredFields = [
    ["firstName", "Vorname"],
    ["jobTitle", "Titel"],
    ["email", "E-Mail-Adresse"],
    ["city", "Ort"],
    ["postalCode", "Postleitzahl"],
    ["street", "Straße"],
  ];
  return requiredFields
    .filter(([key]) => key !== "jobTitle" || !ignoreMissingTitle)
    .filter(([key]) => !String(profileValue?.[key] || "").trim())
    .map(([, label]) => label);
}

function joinedFieldNames(fields) {
  if (fields.length < 2) return fields[0] || "";
  return `${fields.slice(0, -1).join(", ")} und ${fields[fields.length - 1]}`;
}

function profileWarningMessages(profileValue, phoneMode = null, ignoreMissingTitle = false) {
  const missing = missingProfileFields(profileValue, ignoreMissingTitle);
  const mobileMissing = !String(profileValue?.mobile || "").trim();
  const phoneMissing = !String(profileValue?.phone || "").trim();
  const needsMobile = phoneMode === null || phoneMode === "Handy" || phoneMode === "Alles";
  const needsPhone = phoneMode === null || phoneMode === "Festnetz" || phoneMode === "Alles";
  if (phoneMissing && needsPhone) missing.push("Festnetznummer");
  if (mobileMissing && needsMobile) missing.push("Mobilnummer");
  if (missing.length === 0) return [];
  if (missing.length === 1) {
    return [`Information über ${missing[0]} fehlt, bitte EDV kontaktieren!`];
  }
  return [`Informationen über ${joinedFieldNames(missing)} fehlen, bitte EDV kontaktieren!`];
}

function insertedProfileWarningsHtml(
  profileValue,
  phoneMode = signatureSettings.Nummer,
  ignoreMissingTitle = false,
) {
  return profileWarningMessages(profileValue, phoneMode, ignoreMissingTitle)
    .map((message) => `<p style="margin: 0 0 6px; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: #c00000; font-weight: bold;"><b>${escapeHtml(message)}</b></p>`)
    .join("");
}

function showProfileWarnings(
  profileValue,
  phoneMode = signatureSettings.Nummer,
  ignoreMissingTitle = false,
) {
  if (!profileLoaded) {
    profileWarningsElement.replaceChildren();
    profileWarningsElement.hidden = true;
    return;
  }
  const messages = profileWarningMessages(profileValue, phoneMode, ignoreMissingTitle);
  profileWarningsElement.replaceChildren(...messages.map((message) => {
    const paragraph = document.createElement("p");
    paragraph.textContent = message;
    return paragraph;
  }));
  profileWarningsElement.hidden = messages.length === 0;
}

function normalizeEmail(value) {
  return String(value || "").trim().toLocaleLowerCase("de-AT");
}

function emailDomain(value) {
  const email = normalizeEmail(value);
  const separator = email.lastIndexOf("@");
  if (separator <= 0 || separator === email.length - 1) return "";
  return email.slice(separator + 1);
}

function personalName(profileValue) {
  if (!profileValue) return "";
  return [profileValue.firstName, profileValue.lastName]
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .join(" ") || String(profileValue.displayName || profileValue.email || "").trim();
}

function delegatedName(profileValue, settings = signatureSettings) {
  const titleBefore = settings.InsertTitleBefore
    ? String(profileValue?.customAttribute10 || "").trim()
    : "";
  const titleAfter = settings.InsertTitleAfter
    ? String(profileValue?.customAttribute11 || "").trim()
    : "";
  return [titleBefore, personalName(profileValue), titleAfter].filter(Boolean).join(" ");
}

function isFirstNameOnlyProfile(profileValue) {
  if (!profileValue) return false;
  const firstName = String(profileValue.firstName || "").trim();
  const lastName = String(profileValue.lastName || "").trim();
  if (lastName) return false;
  if (firstName) return true;
  const displayName = String(profileValue.displayName || "").replace(/\s+/g, " ").trim();
  return Boolean(displayName && !displayName.includes("@") && !displayName.includes(" "));
}

const CITY_ADDRESS_OVERRIDES = Object.freeze({
  "Neusiedl am See": Object.freeze({
    city: "Neusiedl am See",
    postalCode: "7100",
    street: "Peter-Floridan-Gasse 4/Top 1",
  }),
  Oberwart: Object.freeze({
    city: "Oberwart",
    postalCode: "7400",
    street: "Schulgasse 42/2",
  }),
  "Wr. Neustadt": Object.freeze({
    city: "Wr. Neustadt",
    postalCode: "2700",
    street: "Badener Straße 16",
  }),
});

function applyCityOverride(profileValue, settings = signatureSettings) {
  if (!cityChangeAuthorized) return profileValue;
  const address = CITY_ADDRESS_OVERRIDES[String(settings.CityOverride || "Standard")];
  return address ? { ...profileValue, ...address } : profileValue;
}

function bannerForCity(profileValue = profile) {
  const city = String(profileValue.city || "").trim();
  if (city === "Wien") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_w" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_w.png" border="0" alt="Banner Wien"></a></p>';
  }
  if (city === "St. Pölten-Radlberg" || city === "Krems an der Donau") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_noe_nord" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_noe_nord.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  if (city === "Wr. Neustadt") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_noe_sued" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_noe_sued.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  if (city === "Neusiedl am See" || city === "Oberwart") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_bgld" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_bgld.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  if (city === "Klagenfurt") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_ktn" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_ktn.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  if (city === "Kalsdorf" || city === "Graz" || city == "Leoben") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_stmk" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_stmk.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  if (city === "Linz" || city === "Regau") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_ooe" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_ooe.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  if (city === "Salzburg" || city === "Bruck an der Großglocknerstraße") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_sbg" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_sbg.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  if (city === "Innsbruck") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_t" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_t.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  if (city === "Rankweil") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_vbg" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_vbg.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  return "";
}

function scalePreview(container, content) {
  content.style.transform = "none";
  const styles = getComputedStyle(container);
  const availableWidth = container.clientWidth
    - parseFloat(styles.paddingLeft)
    - parseFloat(styles.paddingRight);
  const availableHeight = container.clientHeight
    - parseFloat(styles.paddingTop)
    - parseFloat(styles.paddingBottom);
  const naturalWidth = content.scrollWidth;
  const naturalHeight = content.scrollHeight;

  if (!naturalWidth || !naturalHeight) return;

  const scale = Math.min(
    1,
    availableWidth / naturalWidth,
    availableHeight / naturalHeight,
  );
  content.style.transform = `scale(${scale})`;
}

function scaleSignaturePreview() {
  scalePreview(signatureButton, previewElement);
  customSignaturesElement.querySelectorAll(".preview").forEach((container) => {
    const content = container.querySelector(".signature-preview-content");
    if (content) scalePreview(container, content);
  });
}

function buildSignature(templateHtml = signatureTemplate, settings = signatureSettings, signatureId = "standard") {
  const sendAs = isFirstNameOnlyProfile(currentDelegation);
  const sendOnBehalf = Boolean(currentDelegation) && !sendAs;
  const selectedProfile = currentDelegation || profile;
  const baseSignatureProfile = sendAs && !String(selectedProfile.firstName || "").trim()
    ? { ...selectedProfile, firstName: String(selectedProfile.displayName || "").trim() }
    : selectedProfile;
  const signatureProfile = applyCityOverride(baseSignatureProfile, settings);
  const sendAsHasDirectNumber = Boolean(
    String(signatureProfile.phone || "").trim()
    || String(signatureProfile.mobile || "").trim(),
  );
  const renderSettings = sendAs
    ? {
        ...settings,
        Nummer: sendAsHasDirectNumber ? "Available" : "Office",
        Confidentiality: false,
        MobileUsage: false,
        MobileUsageText: "",
      }
    : settings;
  const titleBefore = !sendOnBehalf
    && settings.InsertTitleBefore && String(signatureProfile.customAttribute10 || "").trim()
    ? `${String(signatureProfile.customAttribute10).trim()} `
    : "";
  const titleAfter = !sendOnBehalf
    && settings.InsertTitleAfter && String(signatureProfile.customAttribute11 || "").trim()
    ? ` ${String(signatureProfile.customAttribute11).trim()}`
    : "";
  const senderName = personalName(profile);
  const fromName = delegatedName(currentDelegation, settings);
  const delegatedLastNameHtml = sendOnBehalf
    ? `<span style="font-weight: normal;">(im Auftrag von </span><span style="font-weight: bold;">${escapeHtml(fromName)}</span><span style="font-weight: normal;">)</span>`
    : "";
  const values = {
    FirstName: sendOnBehalf ? senderName : signatureProfile.firstName,
    LastName: signatureProfile.lastName,
    Company: signatureProfile.company, City: signatureProfile.city, Street: signatureProfile.street,
    PostalCode: signatureProfile.postalCode, JobTitle: signatureProfile.jobTitle,
    "E-mail": signatureProfile.email, Mobile: signatureProfile.mobile, Phone: signatureProfile.phone,
    TitelVor: titleBefore,
    TitelNach: titleAfter,
    CustomAttribute10: titleBefore,
    CustomAttribute11: titleAfter,
  };
  const signatureBody = templateHtml.replace(/\{([^{}]+)\}/g, (match, key) => {
    if (key === "Phone Mobile Office Number") return phoneLine(signatureProfile, renderSettings);
    if (key === "Banner") return bannerForCity(signatureProfile);
    if (sendOnBehalf && key === "LastName") return delegatedLastNameHtml;
    return Object.hasOwn(values, key) ? escapeHtml(values[key]) : match;
  });
  const signatureContent = greetingHtml(renderSettings) + signatureBody + noticesHtml(renderSettings);
  const safeSignatureId = escapeHtml(signatureId);
  const marker = `<span id="attensam-signature-marker-${safeSignatureId}" data-attensam-signature-id="${safeSignatureId}" style="display:none!important;mso-hide:all;max-height:0;overflow:hidden;font-size:0;line-height:0;color:transparent;">${SIGNATURE_MARKER_TEXT}</span>`;
  const previewHtml = `<div id="${SIGNATURE_MARKER_ID}" data-attensam-signature="v2" data-attensam-signature-id="${safeSignatureId}">${marker}${signatureContent}</div>`;
  const html = `<div id="${SIGNATURE_MARKER_ID}" data-attensam-signature="v2" data-attensam-signature-id="${safeSignatureId}">${marker}${insertedProfileWarningsHtml(signatureProfile, renderSettings.Nummer, sendAs)}${signatureContent}</div>`;
  return { html, previewHtml, signatureProfile, renderSettings, ignoreMissingTitle: sendAs };
}

function renderSignature() {
  const result = buildSignature();
  previewElement.innerHTML = result.previewHtml;
  showProfileWarnings(
    result.signatureProfile,
    result.renderSettings.Nummer,
    result.ignoreMissingTitle,
  );
  previewElement.querySelectorAll("img").forEach((image) => {
    if (!image.complete) image.addEventListener("load", scaleSignaturePreview, { once: true });
  });
  requestAnimationFrame(scaleSignaturePreview);
  const ready = Boolean(result.html && profileLoaded);
  signatureButton.setAttribute("aria-disabled", String(!ready));
  signatureButton.tabIndex = ready ? 0 : -1;
  signatureButton.classList.toggle("ready", ready);
  renderCustomSignatureCards();
  return result.html;
}

function defaultBadge(id) {
  return customSignatures.defaultId === id ? '<span class="default-badge">Standard</span>' : "";
}

function renderCustomSignatureCards() {
  customSignaturesElement.replaceChildren();
  if (!vipAuthorized) return;
  customSignatures.items.forEach((item) => {
    const card = document.createElement("article");
    card.className = "custom-signature-card";
    card.innerHTML = `<div class="custom-signature-title"><span>${escapeHtml(item.title)}</span>${defaultBadge(item.id)}</div><div class="preview ready" role="button" tabindex="0" aria-label="${escapeHtml(item.title)} einfügen"><div class="signature-preview-content"></div></div>`;
    const button = card.querySelector(".preview");
    const content = card.querySelector(".signature-preview-content");
    content.innerHTML = buildSignature(item.html, item.settings || signatureSettings, item.id).previewHtml;
    button.addEventListener("click", () => insertSignature(item.id));
    button.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        insertSignature(item.id);
      }
    });
    button.addEventListener("contextmenu", (event) => openSignatureMenu(event, item.id));
    content.querySelectorAll("img").forEach((image) => {
      if (!image.complete) image.addEventListener("load", () => scalePreview(button, content), { once: true });
    });
    customSignaturesElement.append(card);
    requestAnimationFrame(() => scalePreview(button, content));
  });
  customAddButton.disabled = customSignatures.items.length >= MAX_CUSTOM_SIGNATURES;
  customAddButton.title = customAddButton.disabled ? "Maximal drei Signaturen" : "Benutzerdefinierte Signatur hinzufügen";
}

function openSignatureMenu(event, id) {
  if (!vipAuthorized) return;
  event.preventDefault();
  contextSignatureId = id;
  deleteConfirmationArmed = false;
  deleteCustomButton.textContent = "Signatur löschen";
  editCustomButton.hidden = id === "standard";
  deleteCustomButton.hidden = id === "standard";
  setDefaultButton.disabled = customSignatures.defaultId === id;
  contextMenu.hidden = false;
  const width = 195;
  contextMenu.style.left = `${Math.min(event.clientX, window.innerWidth - width - 8)}px`;
  contextMenu.style.top = `${Math.min(event.clientY, window.innerHeight - 100)}px`;
  setDefaultButton.focus();
}

function closeSignatureMenu() {
  deleteConfirmationArmed = false;
  deleteCustomButton.textContent = "Signatur löschen";
  contextMenu.hidden = true;
}

function mergedDelegatedProfiles(existingProfiles, delegation, fromAddress, updatedAt) {
  const validAfter = Date.now() - (14 * 24 * 60 * 60 * 1000);
  const entries = Object.entries(existingProfiles || {}).filter(([, entry]) => {
    const cachedAt = Date.parse(entry?.updatedAt || "");
    return entry?.profile?.id && Number.isFinite(cachedAt) && cachedAt >= validAfter;
  });
  const result = Object.fromEntries(entries);
  const address = normalizeEmail(fromAddress);
  if (address && delegation?.id) {
    result[address] = { profile: { ...delegation }, updatedAt };
  }
  return Object.fromEntries(
    Object.entries(result)
      .sort((left, right) => Date.parse(right[1].updatedAt || "") - Date.parse(left[1].updatedAt || ""))
      .slice(0, 6),
  );
}

function saveCurrentDelegationLocally() {
  if (!currentDelegation?.id || !currentDelegationAddress) return;
  try {
    const existingRecord = JSON.parse(localStorage.getItem(DELEGATED_PROFILE_LOCAL_CACHE_KEY) || "null");
    const sameOwner = !existingRecord?.ownerId || !profile.id || existingRecord.ownerId === profile.id;
    const existingProfiles = sameOwner ? existingRecord?.profiles : null;
    localStorage.setItem(DELEGATED_PROFILE_LOCAL_CACHE_KEY, JSON.stringify({
      ownerId: profile.id || "",
      ownerMailbox: normalizeEmail(Office.context.mailbox?.userProfile?.emailAddress || profile.email),
      profiles: mergedDelegatedProfiles(
        existingProfiles,
        currentDelegation,
        currentDelegationAddress,
        new Date().toISOString(),
      ),
    }));
  } catch (error) {
    console.warn("Das lokale Absenderprofil konnte nicht zwischengespeichert werden.", error);
  }
}

async function saveCurrentDelegationCache() {
  if (!profileLoaded || !currentDelegation?.id || !currentDelegationAddress) return;
  const roamingSettings = Office.context.roamingSettings;
  const renderData = roamingSettings?.get(AUTO_RENDER_DATA_KEY);
  if (!roamingSettings || !renderData) return;
  const updatedAt = new Date().toISOString();
  roamingSettings.set(AUTO_RENDER_DATA_KEY, {
    ...renderData,
    delegatedProfiles: mergedDelegatedProfiles(
      renderData.delegatedProfiles,
      currentDelegation,
      currentDelegationAddress,
      updatedAt,
    ),
  });
  await new Promise((resolve, reject) => {
    roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error?.message || "Absenderprofil konnte nicht zwischengespeichert werden."));
    });
  });
}

async function saveAutoRenderData() {
  const roamingSettings = Office.context.roamingSettings;
  if (!roamingSettings || !signatureTemplate || !profileLoaded) return;
  const cachedAt = new Date().toISOString();
  const existingRenderData = roamingSettings.get(AUTO_RENDER_DATA_KEY);
  roamingSettings.set(AUTO_RENDER_DATA_KEY, {
    profile: { ...profile },
    mailboxEmail: Office.context.mailbox.userProfile?.emailAddress || profile.email,
    template: signatureTemplate,
    officeNumber: CONFIG.officeNumber,
    settings: { ...signatureSettings },
    accessAuthorized,
    cityChangeAuthorized,
    settingsUpdatedAt: cachedAt,
    graphAuth: {
      clientId: CONFIG.clientId,
      tenantId: CONFIG.tenantId,
    },
    delegatedProfiles: mergedDelegatedProfiles(
      existingRenderData?.delegatedProfiles,
      currentDelegation,
      currentDelegationAddress,
      cachedAt,
    ),
    profileUpdatedAt: cachedAt,
    updatedAt: cachedAt,
  });
  await new Promise((resolve, reject) => {
    roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error?.message || "Signaturdaten konnten nicht gespeichert werden."));
    });
  });
}

function showProfile() {
  renderSignature();
  feedbackButton.disabled = !accessAuthorized || !profileLoaded;
}

function applyMailboxBasics() {
  const mailbox = Office.context.mailbox.userProfile;
  const names = (mailbox.displayName || "").trim().split(/\s+/);
  profile.firstName = names[0] || "";
  profile.lastName = names.length > 1 ? names[names.length - 1] : "";
  profile.email = mailbox.emailAddress || "";
  showProfile();
}

function getValidCachedRenderData() {
  return SignaturePreferences.getValidRenderData();
}

function cachedProfileDate(cached) {
  const value = new Date(cached.profileUpdatedAt || cached.updatedAt || "");
  return Number.isFinite(value.getTime()) ? value.toLocaleDateString("de-AT") : "unbekannt";
}

async function restoreCachedProfile() {
  const cached = getValidCachedRenderData();
  if (!cached || !SignaturePreferences.getAccessAuthorized() || cached.accessAuthorized !== true) return false;
  Object.assign(profile, cached.profile);
  signatureTemplate = cached.template;
  SignaturePreferences.setDepartment(profile.department);
  SignaturePreferences.setTitleAttributes(profile.customAttribute10, profile.customAttribute11);
  customAddButton.hidden = !vipAuthorized;
  mainSettingsLink.hidden = vipAuthorized;
  customSignatures = vipAuthorized
    ? await SignaturePreferences.getCustomSignatures()
    : { requiredRole: VIP_ROLE, defaultId: "standard", items: [] };
  currentDelegation = null;
  currentDelegationAddress = "";
  usingCachedProfile = true;
  profileLoaded = true;
  showProfile();
  setStatus(`Live-Daten konnten nicht geladen werden. Gespeicherte Signaturdaten vom ${cachedProfileDate(cached)} werden verwendet.`);
  return true;
}

async function acquireGraphToken(scopes = ["User.Read"]) {
  if (!hasConfiguredEntraApp()) {
    throw new Error("Bitte Client-ID und Tenant-ID einmal im ATTENSAM_CONFIG-Block oben in taskpane.js eintragen.");
  }
  if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
    throw new Error("Dieser Outlook-Client unterstützt Nested App Authentication 1.1 nicht.");
  }
  if (!msalInstance) {
    const authority = CONFIG.tenantId.startsWith("https://")
      ? CONFIG.tenantId
      : `https://login.microsoftonline.com/${CONFIG.tenantId}`;
    msalInstance = await msal.createNestablePublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority,
      },
      cache: { cacheLocation: "localStorage" },
    });
  }
  const request = { scopes };
  try {
    const result = await msalInstance.acquireTokenSilent(request);
    rememberAuthenticationRoles(result);
    return result.accessToken;
  } catch (error) {
    if (!(error instanceof msal.InteractionRequiredAuthError)) throw error;
    const result = await msalInstance.acquireTokenPopup(request);
    rememberAuthenticationRoles(result);
    return result.accessToken;
  }
}

function getCurrentFrom() {
  return new Promise((resolve) => {
    const from = Office.context.mailbox.item?.from;
    if (!from?.getAsync) {
      resolve(null);
      return;
    }
    from.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else resolve(null);
    });
  });
}

async function loadDelegatedUser(fromDetails) {
  const fallback = {
    displayName: fromDetails.displayName || fromDetails.emailAddress || "",
    firstName: "",
    lastName: "",
    email: fromDetails.emailAddress || "",
    id: "",
    company: "",
    city: "",
    street: "",
    postalCode: "",
    jobTitle: "",
    department: "",
    mobile: "",
    phone: "",
    customAttribute10: "",
    customAttribute11: "",
  };
  try {
    const token = await acquireGraphToken(["User.Read.All"]);
    const select = [
      "id", "displayName", "givenName", "surname", "mail", "userPrincipalName",
      "companyName", "city", "streetAddress", "postalCode", "jobTitle",
      "department", "mobilePhone", "businessPhones", "onPremisesExtensionAttributes",
    ].join(",");
    let response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromDetails.emailAddress)}?$select=${encodeURIComponent(select)}`,
      { headers: { Authorization: `Bearer ${token}` } },
    );
    let user;
    if (response.ok) {
      user = await response.json();
    }
    if (!user) {
      const address = String(fromDetails.emailAddress || "").replaceAll("'", "''");
      const filter = `mail eq '${address}' or userPrincipalName eq '${address}'`;
      response = await fetch(
        `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(filter)}&$select=${encodeURIComponent(select)}&$top=1`,
        { headers: { Authorization: `Bearer ${token}` } },
      );
      if (response.ok) {
        const result = await response.json();
        user = result.value?.[0];
      }
    }
    if (!user) {
      const address = String(fromDetails.emailAddress || "").replaceAll("'", "''");
      const filter = `proxyAddresses/any(proxy:proxy eq 'smtp:${address}')`;
      response = await fetch(
        `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(filter)}&$select=${encodeURIComponent(select)}&$count=true&$top=1`,
        {
          headers: {
            Authorization: `Bearer ${token}`,
            ConsistencyLevel: "eventual",
          },
        },
      );
      if (response.ok) {
        const result = await response.json();
        user = result.value?.[0];
      }
    }
    if (!user) return fallback;
    return {
      displayName: user.displayName || fallback.displayName,
      id: user.id || "",
      firstName: user.givenName || "",
      lastName: user.surname || "",
      email: user.mail || user.userPrincipalName || fallback.email,
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
  } catch (error) {
    console.warn("Das Profil der abweichenden Absenderadresse konnte nicht vollständig geladen werden.", error);
    return fallback;
  }
}

async function refreshDelegationForCurrentFrom() {
  const fromDetails = await getCurrentFrom();
  const fromEmail = normalizeEmail(fromDetails?.emailAddress);
  const ownEmails = new Set([
    normalizeEmail(profile.email),
    normalizeEmail(Office.context.mailbox.userProfile.emailAddress),
  ].filter(Boolean));
  if (!fromEmail || ownEmails.has(fromEmail)) {
    currentDelegation = null;
    currentDelegationAddress = "";
    return;
  }
  currentDelegation = await loadDelegatedUser(fromDetails);
  if (
    !String(currentDelegation.firstName || "").trim()
    &&
    !String(currentDelegation.lastName || "").trim()
    && !String(currentDelegation.department || "").trim()
  ) {
    currentDelegation = null;
    currentDelegationAddress = "";
    return;
  }
  if (currentDelegation.id && profile.id && currentDelegation.id === profile.id) {
    currentDelegation = null;
    currentDelegationAddress = "";
    return;
  }
  currentDelegationAddress = fromEmail;
  saveCurrentDelegationLocally();
  if (profileLoaded) {
    try {
      await saveCurrentDelegationCache();
    } catch (error) {
      console.warn("Das Profil der abweichenden Absenderadresse konnte nicht zwischengespeichert werden.", error);
    }
  }
}

async function loadProfile() {
  setStatus("App wird geladen...");
  try {
    const token = await acquireGraphToken();
    if (!accessAuthorized) {
      await saveAccessDeniedState();
      profileLoaded = false;
      signatureButton.setAttribute("aria-disabled", "true");
      signatureButton.tabIndex = -1;
      signatureButton.classList.remove("ready");
      return;
    }
    setStatus("Signaturdaten werden geladen …");
    const select = [
      "id", "givenName", "surname", "displayName", "mail", "userPrincipalName",
      "companyName", "city", "streetAddress", "postalCode", "jobTitle",
      "department", "mobilePhone", "businessPhones", "onPremisesExtensionAttributes",
    ].join(",");
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me?$select=${encodeURIComponent(select)}`,
      { headers: { Authorization: `Bearer ${token}` } },
    );
    if (!response.ok) throw new Error(`Microsoft Graph: ${response.status}`);
    const user = await response.json();
    Object.assign(profile, {
      id: user.id || "",
      firstName: user.givenName || "",
      lastName: user.surname || "",
      email: user.mail || user.userPrincipalName || profile.email,
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
    });
    usingCachedProfile = false;
    SignaturePreferences.setDepartment(profile.department);
    SignaturePreferences.setTitleAttributes(profile.customAttribute10, profile.customAttribute11);
    customAddButton.hidden = !vipAuthorized;
    mainSettingsLink.hidden = vipAuthorized;
    customSignatures = vipAuthorized
      ? await SignaturePreferences.getCustomSignatures()
      : { requiredRole: VIP_ROLE, defaultId: "standard", items: [] };
    await refreshDelegationForCurrentFrom();
    profileLoaded = true;
    showProfile();
    try {
      await saveAutoRenderData();
    } catch (cacheError) {
      setStatus(`Profil geladen, aber automatische Signaturdaten konnten nicht gespeichert werden: ${cacheError.message}`);
      return;
    }
    setStatus("Signaturdaten erfolgreich geladen.");
  } catch (error) {
    try {
      if (await restoreCachedProfile()) return;
    } catch (cacheError) {
      console.error("Gespeicherte Signaturdaten konnten nicht geladen werden.", cacheError);
    }
    profileLoaded = false;
    if (accessAuthorized && SignaturePreferences.getVipAuthorizationState() === null) {
      mainSettingsLink.hidden = false;
    }
    signatureButton.setAttribute("aria-disabled", "true");
    signatureButton.tabIndex = -1;
    signatureButton.classList.remove("ready");
    setStatus(error.message || "Profildaten konnten nicht geladen werden.");
  }
}

async function saveAccessDeniedState() {
  const roamingSettings = Office.context.roamingSettings;
  if (!roamingSettings) return;
  const existing = roamingSettings.get(AUTO_RENDER_DATA_KEY) || {};
  roamingSettings.set(AUTO_RENDER_DATA_KEY, {
    ...existing,
    accessAuthorized: false,
    accessCheckedAt: new Date().toISOString(),
  });
  await new Promise((resolve) => {
    roamingSettings.saveAsync(() => resolve());
  });
}

async function subjectExcludesSignature() {
  if (Office.context?.platform !== Office.PlatformType?.PC) return false;
  const subject = Office.context.mailbox.item?.subject;
  if (typeof subject?.getAsync !== "function") return false;
  return new Promise((resolve) => {
    subject.getAsync((result) => {
      resolve(
        result.status === Office.AsyncResultStatus.Succeeded
        && String(result.value || "").startsWith(EXCLUDED_SUBJECT_PREFIX),
      );
    });
  });
}

async function insertSignature(customId = "standard") {
  if (!accessAuthorized || !profileLoaded || signatureButton.getAttribute("aria-disabled") === "true") return;
  const body = Office.context.mailbox.item?.body;
  if (!body) {
    setStatus("Bitte eine neue Nachricht öffnen.");
    return;
  }
  if (await subjectExcludesSignature()) {
    if (typeof body.setSignatureAsync === "function") {
      await new Promise((resolve) => {
        body.setSignatureAsync("", { coercionType: Office.CoercionType.Html }, () => resolve());
      });
    }
    setStatus("Für diese Objektinformation wird keine Signatur eingefügt.");
    return;
  }
  if (usingCachedProfile) {
    currentDelegation = null;
    currentDelegationAddress = "";
  }
  else await refreshDelegationForCurrentFrom();
  const item = vipAuthorized && customId !== "standard"
    ? customSignatures.items.find((entry) => entry.id === customId)
    : null;
  const html = item ? buildSignature(item.html, item.settings || signatureSettings, item.id).html : renderSignature();
  const callback = (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      setStatus("Signatur wurde eingefügt.");
    } else {
      setStatus(result.error?.message || "Signatur konnte nicht eingefügt werden.");
    }
  };
  if (Office.context.requirements.isSetSupported("Mailbox", "1.10") && body.setSignatureAsync) {
    body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, callback);
  } else {
    body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, callback);
  }
}

async function initialize() {
  try {
    signatureMain.hidden = true;
    taskpaneAccessDenied.hidden = true;
    mainSettingsLink.hidden = true;
    setStatus("App wird geladen...");
    const cachedAccessState = SignaturePreferences.getAccessAuthorizationState();
    accessAuthorized = cachedAccessState === true;
    const cachedVipState = SignaturePreferences.getVipAuthorizationState();
    vipAuthorized = cachedVipState === true;
    try {
      signatureTemplate = await fetch("template.html", { cache: "no-store" }).then((response) => {
        if (!response.ok) throw new Error("template.html konnte nicht geladen werden.");
        return response.text();
      });
    } catch (templateError) {
      const cached = getValidCachedRenderData();
      if (!cached) throw templateError;
      signatureTemplate = cached.template;
    }
    signatureSettings = await SignaturePreferences.getSettings();
    applyMailboxBasics();
    await loadProfile();
  } catch (error) {
    setStatus(error.message);
  }
}

async function openFeedbackPage() {
  if (!accessAuthorized || !profileLoaded) return;
  feedbackButton.disabled = true;
  setStatus("Feedback wird geöffnet …");
  try {
    await acquireGraphToken(["User.Read", "Mail.Send"]);
    window.location.href = "feedback.html?view=feedback";
  } catch (error) {
    feedbackButton.disabled = false;
    setStatus(`Feedback konnte nicht geöffnet werden: ${readableError(error)}`);
  }
}

signatureButton.addEventListener("click", () => insertSignature("standard"));
feedbackButton.addEventListener("click", openFeedbackPage);
signatureButton.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    insertSignature();
  }
});
signatureButton.addEventListener("contextmenu", (event) => openSignatureMenu(event, "standard"));
customAddButton.addEventListener("click", () => {
  if (!vipAuthorized || customSignatures.items.length >= MAX_CUSTOM_SIGNATURES) return;
  openCustomEditor();
});
customCancelButton.addEventListener("click", resetCustomEditor);
customHtmlInput.addEventListener("paste", (event) => {
  const clipboard = event.clipboardData;
  if (!clipboard) return;
  const clipboardHtml = clipboard.getData("text/html");
  const clipboardText = clipboard.getData("text/plain");
  const clipboardUrl = clipboard.getData("text/uri-list").split(/\r?\n/).find((value) => /^https:\/\//i.test(value.trim()));
  const containsLocalImage = Array.from(clipboard.items || []).some((item) => item.type.startsWith("image/"));
  event.preventDefault();
  if (clipboardHtml) {
    const sanitized = sanitizeCustomSignatureHtml(clipboardHtml);
    if (sanitized) insertHtmlAtEditorCursor(sanitized);
    return;
  }
  if (containsLocalImage && clipboardUrl) {
    insertHtmlAtEditorCursor(`<img src="${escapeHtml(clipboardUrl.trim())}" alt="">`);
    return;
  }
  if (clipboardText) {
    insertHtmlAtEditorCursor(escapeHtml(clipboardText).replace(/\r\n|\r|\n/g, "<br>"));
    return;
  }
  if (containsLocalImage) {
    setStatus("Das Bild besitzt keine öffentliche HTTPS-Quelladresse und kann deshalb nicht dauerhaft in der Signatur gespeichert werden.");
  }
});
customSaveButton.addEventListener("click", async () => {
  if (!vipAuthorized) return;
  const title = customTitleInput.value.replace(/\s+/g, " ").trim();
  const html = sanitizeCustomSignatureHtml(customHtmlInput.innerHTML);
  if (!title || !html) {
    setStatus("Bitte Titel und Signatur eingeben.");
    return;
  }
  if (html.length > 7000) {
    setStatus("Die Signatur ist zu groß. Bitte Inhalt oder Formatierungen reduzieren.");
    return;
  }
  if (!editingCustomSignatureId && customSignatures.items.length >= MAX_CUSTOM_SIGNATURES) {
    setStatus("Maximal drei benutzerdefinierte Signaturen sind erlaubt.");
    return;
  }
  const id = editingCustomSignatureId || `custom-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  const existingItem = customSignatures.items.find((item) => item.id === editingCustomSignatureId);
  const updatedItem = {
    id,
    title,
    html,
    settings: existingItem?.settings || { ...signatureSettings },
    updatedAt: new Date().toISOString(),
  };
  const items = editingCustomSignatureId
    ? customSignatures.items.map((item) => item.id === editingCustomSignatureId ? updatedItem : item)
    : [...customSignatures.items, updatedItem];
  try {
    customSignatures = await SignaturePreferences.saveCustomSignatures({
      ...customSignatures,
      items,
    });
    const wasEditing = Boolean(editingCustomSignatureId);
    resetCustomEditor();
    renderCustomSignatureCards();
    setStatus(wasEditing ? "Benutzerdefinierte Signatur wurde aktualisiert." : "Benutzerdefinierte Signatur wurde gespeichert.");
  } catch (error) {
    setStatus(error.message || "Signatur konnte nicht gespeichert werden.");
  }
});
setDefaultButton.addEventListener("click", async () => {
  try {
    customSignatures = await SignaturePreferences.saveCustomSignatures({ ...customSignatures, defaultId: contextSignatureId });
    renderCustomSignatureCards();
    setStatus(contextSignatureId === "standard" ? "Standard-Signatur wurde als Standard festgelegt." : "Benutzerdefinierte Signatur wurde als Standard festgelegt.");
  } catch (error) {
    setStatus(error.message || "Standard konnte nicht gespeichert werden.");
  } finally {
    closeSignatureMenu();
  }
});
openSignatureSettingsButton.addEventListener("click", () => {
  const signatureId = contextSignatureId || "standard";
  window.location.href = `taskpane.html?view=settings&signature=${encodeURIComponent(signatureId)}`;
});
editCustomButton.addEventListener("click", () => {
  const item = customSignatures.items.find((entry) => entry.id === contextSignatureId);
  if (item) openCustomEditor(item);
  closeSignatureMenu();
});
deleteCustomButton.addEventListener("click", async () => {
  const item = customSignatures.items.find((entry) => entry.id === contextSignatureId);
  if (!item) return;
  if (!deleteConfirmationArmed) {
    deleteConfirmationArmed = true;
    deleteCustomButton.textContent = "Wirklich löschen?";
    return;
  }
  try {
    const items = customSignatures.items.filter((entry) => entry.id !== contextSignatureId);
    customSignatures = await SignaturePreferences.saveCustomSignatures({
      ...customSignatures,
      defaultId: customSignatures.defaultId === contextSignatureId ? "standard" : customSignatures.defaultId,
      items,
    });
    renderCustomSignatureCards();
    if (editingCustomSignatureId === contextSignatureId) resetCustomEditor();
    setStatus("Benutzerdefinierte Signatur wurde gelöscht.");
  } catch (error) {
    setStatus(error.message || "Signatur konnte nicht gelöscht werden.");
  } finally {
    closeSignatureMenu();
  }
});
document.addEventListener("click", (event) => {
  if (!contextMenu.hidden && !contextMenu.contains(event.target)) closeSignatureMenu();
});
window.addEventListener("resize", scaleSignaturePreview);
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) initialize();
  else setStatus("Diese Seite muss als Outlook-Add-In geöffnet werden.");
});

})();

(function feedbackRoute(){
  if (new URLSearchParams(window.location.search).get("view") !== "feedback") return;
/* global Office, msal, SignaturePreferences */

const REQUIRED_ROLE = "ATS.Signature";
const form = document.getElementById("feedback-form");
const feedbackMain = document.getElementById("feedback-main");
const feedbackAccessDenied = document.getElementById("feedback-access-denied");
const categorySelect = document.getElementById("feedback-category");
const messageInput = document.getElementById("feedback-message");
const sendButton = document.getElementById("send-button");
const feedbackStatus = document.getElementById("feedback-status");
let feedbackMsalInstance;
let senderName = "Unbekannter Benutzer";

function setFeedbackStatus(message, isError = false) {
  feedbackStatus.textContent = message;
  feedbackStatus.classList.toggle("error", isError);
}

function setFeedbackControlsDisabled(disabled) {
  categorySelect.disabled = disabled;
  messageInput.disabled = disabled;
  sendButton.disabled = disabled;
}

function authenticationRoles(result) {
  let tokenClaims = {};
  try {
    const encoded = String(result?.idToken || "").split(".")[1];
    if (encoded) {
      const normalized = encoded.replace(/-/g, "+").replace(/_/g, "/");
      tokenClaims = JSON.parse(atob(normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=")));
    }
  } catch {
    tokenClaims = {};
  }
  return new Set([
    ...(Array.isArray(result?.idTokenClaims?.roles) ? result.idTokenClaims.roles : []),
    ...(Array.isArray(result?.account?.idTokenClaims?.roles) ? result.account.idTokenClaims.roles : []),
    ...(Array.isArray(tokenClaims?.roles) ? tokenClaims.roles : []),
  ].map((role) => String(role).trim()).filter(Boolean));
}

async function acquireFeedbackToken() {
  if (!hasConfiguredEntraApp()) {
    throw new Error("Bitte Client-ID und Tenant-ID einmal im ATTENSAM_CONFIG-Block oben in taskpane.js eintragen.");
  }
  if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
    throw new Error("Dieser Outlook-Client unterstützt Nested App Authentication 1.1 nicht.");
  }
  if (!feedbackMsalInstance) {
    const authority = ATTENSAM_CONFIG.tenantId.startsWith("https://")
      ? ATTENSAM_CONFIG.tenantId
      : `https://login.microsoftonline.com/${ATTENSAM_CONFIG.tenantId}`;
    feedbackMsalInstance = await msal.createNestablePublicClientApplication({
      auth: { clientId: ATTENSAM_CONFIG.clientId, authority },
      cache: { cacheLocation: "localStorage" },
    });
  }
  const request = { scopes: ["User.Read", "Mail.Send"] };
  try {
    return await feedbackMsalInstance.acquireTokenSilent(request);
  } catch (error) {
    if (!(error instanceof msal.InteractionRequiredAuthError)) throw error;
    return feedbackMsalInstance.acquireTokenPopup(request);
  }
}

function feedbackSubject(category) {
  if (category === "Fehler") return `[ATS Signatures] Fehlermeldung von ${senderName}`;
  if (category === "Wünsche") return `[ATS Signatures] Wunsch von ${senderName}`;
  if (category === "Feedback") return `[ATS Signatures] Feedback von ${senderName}`;
  if (category === "Hilfe") return `[ATS Signatures] ${senderName} braucht Hilfe`;
  throw new Error("Ungültige Feedback-Kategorie.");
}

async function initializeFeedback() {
  feedbackMain.hidden = true;
  feedbackAccessDenied.hidden = true;
  setFeedbackControlsDisabled(true);
  setFeedbackStatus("App wird geladen...");
  try {
    const authentication = await acquireFeedbackToken();
    const roles = authenticationRoles(authentication);
    const accessAuthorized = roles.has(REQUIRED_ROLE);
    SignaturePreferences.setAccessAuthorized(accessAuthorized);
    feedbackMain.hidden = !accessAuthorized;
    feedbackAccessDenied.hidden = accessAuthorized;
    if (!accessAuthorized) {
      setFeedbackStatus("Sie haben kein Zugriff auf dieses Add-In, bitte EDV kontaktieren!", true);
      return;
    }
    const select = "givenName,surname,displayName,mail,userPrincipalName";
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me?$select=${encodeURIComponent(select)}`,
      { headers: { Authorization: `Bearer ${authentication.accessToken}` } },
    );
    if (!response.ok) throw new Error(`Microsoft Graph: ${response.status}`);
    const user = await response.json();
    senderName = [user.givenName, user.surname].map((value) => String(value || "").trim()).filter(Boolean).join(" ")
      || user.displayName
      || user.mail
      || user.userPrincipalName
      || Office.context.mailbox?.userProfile?.displayName
      || Office.context.mailbox?.userProfile?.emailAddress
      || senderName;
    setFeedbackControlsDisabled(false);
    setFeedbackStatus("Feedback-Formular ist bereit.");
  } catch (error) {
    feedbackMain.hidden = false;
    setFeedbackControlsDisabled(true);
    setFeedbackStatus(`Feedback konnte nicht geladen werden: ${readableError(error)}`, true);
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const category = String(categorySelect.value || "");
  const message = String(messageInput.value || "").trim();
  if (!["Fehler", "Wünsche", "Feedback", "Hilfe"].includes(category) || !message || message.length > 5000) {
    setFeedbackStatus("Bitte Kategorie und Nachricht überprüfen.", true);
    return;
  }
  setFeedbackControlsDisabled(true);
  setFeedbackStatus("E-Mail wird gesendet …");
  let sent = false;
  try {
    const authentication = await acquireFeedbackToken();
    if (!authenticationRoles(authentication).has(REQUIRED_ROLE)) {
      throw new Error("Sie haben kein Zugriff auf dieses Add-In, bitte EDV kontaktieren!");
    }
    const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${authentication.accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message: {
          subject: feedbackSubject(category),
          body: { contentType: "Text", content: message },
          toRecipients: [{ emailAddress: { address: ATTENSAM_CONFIG.feedbackEmail } }],
        },
        saveToSentItems: true,
      }),
    });
    if (!response.ok) throw new Error(`Microsoft Graph: ${response.status}`);
    sent = true;
    messageInput.value = "";
    sendButton.textContent = "Gesendet ✅";
    setFeedbackStatus("Feedback wurde erfolgreich gesendet.");
    window.setTimeout(() => {
      sendButton.textContent = "E-Mail senden";
      setFeedbackControlsDisabled(false);
      setFeedbackStatus("Feedback-Formular ist bereit.");
    }, 3000);
  } catch (error) {
    setFeedbackStatus(`Feedback konnte nicht gesendet werden: ${readableError(error)}`, true);
  } finally {
    if (!sent) setFeedbackControlsDisabled(false);
  }
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) initializeFeedback();
  else setFeedbackStatus("Diese Seite muss als Outlook-Add-In geöffnet werden.", true);
});

})();

(function compactRoute(){
  if (!(new URLSearchParams(window.location.search).get("view") === "settings" && new URLSearchParams(window.location.search).get("mobile") !== "1")) return;
/* global Office, SignaturePreferences, AttensamSignatureRuntime, DOMParser */

const AUTO_RENDER_DATA_KEY = "attensam.signature.render-data.v1";
const SETTINGS_SIGNATURE_MARKER_ID = "attensam-signature-root";
const SETTINGS_SIGNATURE_MARKER_TEXT = "Attensam-Signatur";
const SETTINGS_SIGNATURE_MARKER_ID_PREFIX = "attensam-signature-marker-";
const SETTINGS_SIGNATURE_ID = String(new URLSearchParams(window.location.search).get("signature") || "standard").replace(/[^a-zA-Z0-9_-]/g, "") || "standard";

const phoneModeSelect = document.getElementById("phone-mode");
const cityChangeField = document.getElementById("city-change-field");
const cityChangeSelect = document.getElementById("city-change");
const edvHotlineOption = document.getElementById("edv-hotline-option");
const combinedPhoneWarning = document.getElementById("combined-phone-warning");
const mobilePhoneWarning = document.getElementById("mobile-phone-warning");
const landlinePhoneWarning = document.getElementById("landline-phone-warning");
const greetingModeSelect = document.getElementById("greeting-mode");
const customGreetingField = document.getElementById("custom-greeting-field");
const customGreetingInput = document.getElementById("custom-greeting");
const greetingLinesField = document.getElementById("greeting-lines-field");
const greetingLinesSelect = document.getElementById("greeting-lines");
const titleBeforeField = document.getElementById("title-before-field");
const insertTitleBeforeCheckbox = document.getElementById("insert-title-before");
const titleAfterField = document.getElementById("title-after-field");
const insertTitleAfterCheckbox = document.getElementById("insert-title-after");
const mobileUsageCheckbox = document.getElementById("mobile-usage");
const mobileUsageTextField = document.getElementById("mobile-usage-text-field");
const mobileUsageTextInput = document.getElementById("mobile-usage-text");
const confidentialityCheckbox = document.getElementById("confidentiality");
const autoInsertRepliesCheckbox = document.getElementById("auto-insert-replies");
const autoInsertForwardsCheckbox = document.getElementById("auto-insert-forwards");
const autoInsertMeetingsCheckbox = document.getElementById("auto-insert-meetings");
const skipInternalOnlyCheckbox = document.getElementById("skip-internal-only");
const skipInternalNewMailField = document.getElementById("skip-internal-new-mail-field");
const skipInternalNewMailCheckbox = document.getElementById("skip-internal-new-mail");
const settingsStatus = document.getElementById("settings-status");
const settingsHeading = document.getElementById("settings-heading");
const settingsMain = document.getElementById("settings-main");
const settingsAccessDenied = document.getElementById("settings-access-denied");
let currentSettings;
let settingsProfile = null;
let cityChangeAuthorized = false;

function setSettingsStatus(message) {
  settingsStatus.textContent = message;
}

function updateInternalInsertionVisibility() {
  skipInternalNewMailField.hidden = !skipInternalOnlyCheckbox.checked;
  skipInternalNewMailCheckbox.disabled = skipInternalOnlyCheckbox.disabled
    || !skipInternalOnlyCheckbox.checked;
}

function updateMobileUsageVisibility() {
  mobileUsageTextField.hidden = !mobileUsageCheckbox.checked;
  mobileUsageTextInput.disabled = mobileUsageCheckbox.disabled || !mobileUsageCheckbox.checked;
}

function updatePhoneWarnings() {
  if (!settingsProfile) {
    combinedPhoneWarning.hidden = true;
    mobilePhoneWarning.hidden = true;
    landlinePhoneWarning.hidden = true;
    return;
  }
  const mode = phoneModeSelect.value;
  const mobileMissing = !String(settingsProfile.mobile || "").trim();
  const phoneMissing = !String(settingsProfile.phone || "").trim();
  const bothNeeded = mode === "Alles";
  combinedPhoneWarning.hidden = !(bothNeeded && mobileMissing && phoneMissing);
  mobilePhoneWarning.hidden = !mobileMissing
    || (mode !== "Handy" && mode !== "Alles" && mode !== "EDVHotline")
    || (bothNeeded && phoneMissing);
  landlinePhoneWarning.hidden = !phoneMissing
    || (mode !== "Festnetz" && mode !== "Alles")
    || (bothNeeded && mobileMissing);
}

function updateGreetingVisibility() {
  const isCustom = greetingModeSelect.value === "MfGCustom";
  const hasGreeting = greetingModeSelect.value !== "MfG0";
  customGreetingField.hidden = !isCustom;
  greetingLinesField.hidden = !hasGreeting;
  customGreetingInput.disabled = greetingModeSelect.disabled || !isCustom;
  greetingLinesSelect.disabled = greetingModeSelect.disabled || !hasGreeting;
}

function setControlsDisabled(disabled) {
  phoneModeSelect.disabled = disabled;
  cityChangeSelect.disabled = disabled || !cityChangeAuthorized;
  greetingModeSelect.disabled = disabled;
  customGreetingInput.disabled = disabled || greetingModeSelect.value !== "MfGCustom";
  greetingLinesSelect.disabled = disabled || greetingModeSelect.value === "MfG0";
  insertTitleBeforeCheckbox.disabled = disabled;
  insertTitleAfterCheckbox.disabled = disabled;
  mobileUsageCheckbox.disabled = disabled;
  mobileUsageTextInput.disabled = disabled || !mobileUsageCheckbox.checked;
  confidentialityCheckbox.disabled = disabled;
  autoInsertRepliesCheckbox.disabled = disabled;
  autoInsertForwardsCheckbox.disabled = disabled;
  autoInsertMeetingsCheckbox.disabled = disabled;
  skipInternalOnlyCheckbox.disabled = disabled;
  skipInternalNewMailCheckbox.disabled = disabled || !skipInternalOnlyCheckbox.checked;
}

function settingsEmailDomain(value) {
  const email = String(value || "").trim().toLocaleLowerCase("de-AT");
  const separator = email.lastIndexOf("@");
  return separator > 0 && separator < email.length - 1 ? email.slice(separator + 1) : "";
}

function getSettingsRecipients(field) {
  if (typeof field?.getAsync !== "function") return Promise.resolve([]);
  return new Promise((resolve) => {
    field.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        resolve(null);
        return;
      }
      resolve((result.value || []).map((recipient) => String(recipient?.emailAddress || "").trim()));
    });
  });
}

async function settingsHasOnlyInternalRecipients(renderData) {
  const senderDomain = settingsEmailDomain(
    renderData?.mailboxEmail
    || renderData?.profile?.email
    || Office.context.mailbox?.userProfile?.emailAddress,
  );
  if (!senderDomain) return false;
  const item = Office.context.mailbox.item;
  const recipientGroups = await Promise.all([
    getSettingsRecipients(item?.to),
    getSettingsRecipients(item?.cc),
    getSettingsRecipients(item?.bcc),
  ]);
  if (recipientGroups.some((group) => group === null)) return false;
  const addresses = recipientGroups.flat();
  return Boolean(
    addresses.length
    && addresses.every((address) => settingsEmailDomain(address) === senderDomain),
  );
}

function getSettingsComposeType() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    if (typeof item?.getComposeTypeAsync !== "function") {
      resolve("");
      return;
    }
    item.getComposeTypeAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded
        ? String(result.value?.composeType || "")
        : "");
    });
  });
}

function settingsSubjectExcludesSignature() {
  if (Office.context?.platform !== Office.PlatformType?.PC) return Promise.resolve(false);
  const subject = Office.context.mailbox.item?.subject;
  if (typeof subject?.getAsync !== "function") return Promise.resolve(false);
  return new Promise((resolve) => {
    subject.getAsync((result) => {
      resolve(
        result.status === Office.AsyncResultStatus.Succeeded
        && String(result.value || "").startsWith("Ihre Objektinformation - "),
      );
    });
  });
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

function findMarkedSignature(document) {
  const hiddenMarker = Array.from(document.querySelectorAll("span"))
    .find((element) => element.textContent?.includes(SETTINGS_SIGNATURE_MARKER_TEXT));
  return document.getElementById(SETTINGS_SIGNATURE_MARKER_ID)
    || document.querySelector('[data-attensam-signature="v1"]')
    || document.querySelector('[data-attensam-signature="v2"]')
    || document.querySelector(`[id^="${SETTINGS_SIGNATURE_MARKER_ID_PREFIX}"]`)?.parentElement
    || hiddenMarker?.parentElement;
}

function replaceMarkedSignature(bodyHtml, signatureHtml) {
  const document = new DOMParser().parseFromString(bodyHtml, "text/html");
  const existingSignature = findMarkedSignature(document);
  if (!existingSignature) return null;
  existingSignature.outerHTML = signatureHtml;
  return document.body.innerHTML;
}

function readInsertedSignatureId(existingSignature) {
  const attributeId = String(existingSignature?.getAttribute("data-attensam-signature-id") || "")
    .replace(/[^a-zA-Z0-9_-]/g, "");
  if (attributeId) return attributeId;
  const marker = existingSignature?.querySelector(`[id^="${SETTINGS_SIGNATURE_MARKER_ID_PREFIX}"]`);
  return String(marker?.id || "")
    .slice(SETTINGS_SIGNATURE_MARKER_ID_PREFIX.length)
    .replace(/[^a-zA-Z0-9_-]/g, "");
}

async function updateInsertedSignature() {
  const body = Office.context.mailbox.item?.body;
  if (!body) return false;
  const cachedRenderData = SignaturePreferences.getValidRenderData();
  const renderData = cachedRenderData
    ? { ...cachedRenderData, cityChangeAuthorized: SignaturePreferences.getCityChangeAuthorized() }
    : null;
  if (!renderData) return false;
  const settingsComposeType = await getSettingsComposeType();
  if (!settingsComposeType) return false;
  const composeModeAllowsInsertion = settingsComposeType === "newMail"
    || (settingsComposeType === "reply" && currentSettings.AutoInsertReplies === true)
    || (settingsComposeType === "forward" && currentSettings.AutoInsertForwards === true);
  const internalSuppressionApplies = currentSettings.SkipInternalOnly === true
    && (settingsComposeType !== "newMail" || currentSettings.SkipInternalOnNewMail === true)
    && await settingsHasOnlyInternalRecipients(renderData);
  const subjectSuppressionApplies = await settingsSubjectExcludesSignature();
  if (!composeModeAllowsInsertion || internalSuppressionApplies || subjectSuppressionApplies) {
    if (typeof body.setSignatureAsync === "function") {
      await setCurrentSignature(body, "");
      return true;
    }
    if (typeof body.getAsync === "function" && typeof body.setAsync === "function") {
      const bodyHtml = await getBodyHtml(body);
      const bodyDocument = new DOMParser().parseFromString(bodyHtml, "text/html");
      const existingSignature = findMarkedSignature(bodyDocument);
      if (!existingSignature) return false;
      existingSignature.remove();
      await setBodyHtml(body, bodyDocument.body.innerHTML);
      return true;
    }
    return false;
  }
  const delegation = await new Promise((resolve) => {
    AttensamSignatureRuntime.resolveDelegation(renderData, resolve);
  });
  const verifiedDelegation = delegation?.id ? delegation : null;
  const customRecord = await SignaturePreferences.getCustomSignatures();
  const renderCustomRecord = SETTINGS_SIGNATURE_ID === "standard"
    ? customRecord
    : {
        ...customRecord,
        items: customRecord.items.map((item) => item.id === SETTINGS_SIGNATURE_ID
          ? { ...item, settings: { ...currentSettings } }
          : item),
      };
  const html = AttensamSignatureRuntime.renderSignature(
    renderData,
    currentSettings,
    verifiedDelegation,
    renderCustomRecord,
    SETTINGS_SIGNATURE_ID,
  );
  if (typeof body.setSignatureAsync === "function") {
    await setCurrentSignature(body, html);
    return true;
  }
  if (typeof body.getAsync !== "function") return false;
  const bodyHtml = await getBodyHtml(body);
  const bodyDocument = new DOMParser().parseFromString(bodyHtml, "text/html");
  const existingSignature = findMarkedSignature(bodyDocument);
  if (!existingSignature) return false;
  const insertedSignatureId = readInsertedSignatureId(existingSignature);
  if (insertedSignatureId && insertedSignatureId !== SETTINGS_SIGNATURE_ID) return false;
  const replacedBodyHtml = replaceMarkedSignature(bodyHtml, html);
  if (replacedBodyHtml !== null && body.setAsync) {
    await setBodyHtml(body, replacedBodyHtml);
    return true;
  }
  return false;
}

async function initializeSettings() {
  try {
    const accessAuthorized = SignaturePreferences.getAccessAuthorized();
    settingsMain.hidden = !accessAuthorized;
    settingsAccessDenied.hidden = accessAuthorized;
    if (!accessAuthorized) {
      setSettingsStatus("Sie haben kein Zugriff auf dieses Add-In, bitte EDV kontaktieren!");
      return;
    }
    const isVipUser = SignaturePreferences.getVipAuthorized();
    if (!isVipUser && SETTINGS_SIGNATURE_ID !== "standard") {
      throw new Error("Diese Signatur-Einstellungen sind nur für VIP-Benutzer verfügbar.");
    }
    const customRecord = await SignaturePreferences.getCustomSignatures();
    const selectedItem = customRecord.items.find((item) => item.id === SETTINGS_SIGNATURE_ID);
    if (SETTINGS_SIGNATURE_ID !== "standard" && !selectedItem) throw new Error("Die ausgewählte Signatur wurde nicht gefunden.");
    settingsHeading.textContent = !isVipUser
      ? "Einstellungen"
      : SETTINGS_SIGNATURE_ID === "standard"
      ? "Einstellungen: Standard"
      : `Einstellungen: ${selectedItem.title}`;
    currentSettings = await SignaturePreferences.getSettingsForSignature(SETTINGS_SIGNATURE_ID);
    cityChangeAuthorized = SignaturePreferences.getCityChangeAuthorized();
    cityChangeField.hidden = !cityChangeAuthorized;
    cityChangeSelect.value = currentSettings.CityOverride;
    settingsProfile = SignaturePreferences.getValidRenderData()?.profile || null;
    const department = SignaturePreferences.getDepartment();
    const titleAttributes = SignaturePreferences.getTitleAttributes();
    const canUseEdvHotline = department.trim().toLocaleUpperCase("de-AT") === "IT";
    edvHotlineOption.hidden = !canUseEdvHotline;
    edvHotlineOption.disabled = !canUseEdvHotline;
    phoneModeSelect.value = currentSettings.Nummer === "EDVHotline" && !canUseEdvHotline
      ? "Alles"
      : currentSettings.Nummer;
    updatePhoneWarnings();
    greetingModeSelect.value = currentSettings.MfG;
    customGreetingInput.value = currentSettings.CustomGreeting;
    greetingLinesSelect.value = String(currentSettings.GreetingLines);
    updateGreetingVisibility();
    titleBeforeField.hidden = !titleAttributes.customAttribute10;
    titleAfterField.hidden = !titleAttributes.customAttribute11;
    insertTitleBeforeCheckbox.checked = currentSettings.InsertTitleBefore;
    insertTitleAfterCheckbox.checked = currentSettings.InsertTitleAfter;
    mobileUsageCheckbox.checked = currentSettings.MobileUsage;
    mobileUsageTextInput.value = currentSettings.MobileUsageText;
    updateMobileUsageVisibility();
    confidentialityCheckbox.checked = currentSettings.Confidentiality;
    autoInsertRepliesCheckbox.checked = currentSettings.AutoInsertReplies;
    autoInsertForwardsCheckbox.checked = currentSettings.AutoInsertForwards;
    autoInsertMeetingsCheckbox.checked = currentSettings.AutoInsertMeetings;
    skipInternalOnlyCheckbox.checked = currentSettings.SkipInternalOnly;
    skipInternalNewMailCheckbox.checked = currentSettings.SkipInternalOnNewMail;
    updateInternalInsertionVisibility();
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
    currentSettings = await SignaturePreferences.saveSettingsForSignature(SETTINGS_SIGNATURE_ID, {
      Nummer: phoneModeSelect.value,
      MfG: greetingModeSelect.value,
      CustomGreeting: customGreetingInput.value,
      GreetingLines: Number(greetingLinesSelect.value),
      CityOverride: cityChangeAuthorized ? cityChangeSelect.value : currentSettings.CityOverride,
      AutoInsert: true,
      AutoInsertReplies: autoInsertRepliesCheckbox.checked,
      AutoInsertForwards: autoInsertForwardsCheckbox.checked,
      AutoInsertMeetings: autoInsertMeetingsCheckbox.checked,
      SkipInternalOnly: skipInternalOnlyCheckbox.checked,
      SkipInternalOnNewMail: skipInternalNewMailCheckbox.checked,
      InsertTitleBefore: insertTitleBeforeCheckbox.checked,
      InsertTitleAfter: insertTitleAfterCheckbox.checked,
      MobileUsage: mobileUsageCheckbox.checked,
      MobileUsageText: mobileUsageTextInput.value,
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

phoneModeSelect.addEventListener("change", () => {
  updatePhoneWarnings();
  saveSettings();
});
cityChangeSelect.addEventListener("change", saveSettings);
greetingModeSelect.addEventListener("change", () => {
  updateGreetingVisibility();
  saveSettings();
});
customGreetingInput.addEventListener("change", saveSettings);
greetingLinesSelect.addEventListener("change", saveSettings);
insertTitleBeforeCheckbox.addEventListener("change", saveSettings);
insertTitleAfterCheckbox.addEventListener("change", saveSettings);
mobileUsageCheckbox.addEventListener("change", () => {
  updateMobileUsageVisibility();
  saveSettings();
});
mobileUsageTextInput.addEventListener("change", saveSettings);
confidentialityCheckbox.addEventListener("change", saveSettings);
autoInsertRepliesCheckbox.addEventListener("change", saveSettings);
autoInsertForwardsCheckbox.addEventListener("change", saveSettings);
autoInsertMeetingsCheckbox.addEventListener("change", saveSettings);
skipInternalOnlyCheckbox.addEventListener("change", () => {
  updateInternalInsertionVisibility();
  saveSettings();
});
skipInternalNewMailCheckbox.addEventListener("change", saveSettings);

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) initializeSettings();
  else setSettingsStatus("Diese Seite muss als Outlook-Add-In geöffnet werden.");
});

})();

(function compactRoute(){
  if (!(new URLSearchParams(window.location.search).get("view") === "settings" && new URLSearchParams(window.location.search).get("mobile") === "1")) return;
/* global Office, msal, SignaturePreferences */

const CONFIG = ATTENSAM_CONFIG;
const AUTO_RENDER_DATA_KEY = "attensam.signature.render-data.v1";
const MOBILE_SETTINGS_SIGNATURE_ID = String(new URLSearchParams(window.location.search).get("signature") || "standard").replace(/[^a-zA-Z0-9_-]/g, "") || "standard";

const phoneModeSelect = document.getElementById("phone-mode");
const cityChangeField = document.getElementById("city-change-field");
const cityChangeSelect = document.getElementById("city-change");
const edvHotlineOption = document.getElementById("edv-hotline-option");
const combinedPhoneWarning = document.getElementById("combined-phone-warning");
const mobilePhoneWarning = document.getElementById("mobile-phone-warning");
const landlinePhoneWarning = document.getElementById("landline-phone-warning");
const greetingModeSelect = document.getElementById("greeting-mode");
const customGreetingField = document.getElementById("custom-greeting-field");
const customGreetingInput = document.getElementById("custom-greeting");
const greetingLinesField = document.getElementById("greeting-lines-field");
const greetingLinesSelect = document.getElementById("greeting-lines");
const titleBeforeField = document.getElementById("title-before-field");
const insertTitleBeforeCheckbox = document.getElementById("insert-title-before");
const titleAfterField = document.getElementById("title-after-field");
const insertTitleAfterCheckbox = document.getElementById("insert-title-after");
const mobileUsageCheckbox = document.getElementById("mobile-usage");
const mobileUsageTextField = document.getElementById("mobile-usage-text-field");
const mobileUsageTextInput = document.getElementById("mobile-usage-text");
const confidentialityCheckbox = document.getElementById("confidentiality");
const autoInsertRepliesCheckbox = document.getElementById("auto-insert-replies");
const autoInsertForwardsCheckbox = document.getElementById("auto-insert-forwards");
const autoInsertMeetingsCheckbox = document.getElementById("auto-insert-meetings");
const skipInternalOnlyCheckbox = document.getElementById("skip-internal-only");
const skipInternalNewMailField = document.getElementById("skip-internal-new-mail-field");
const skipInternalNewMailCheckbox = document.getElementById("skip-internal-new-mail");
const settingsStatus = document.getElementById("settings-status");
const closeButton = document.getElementById("close-button");
const settingsHeading = document.getElementById("settings-heading");
const settingsMain = document.getElementById("settings-main");
const settingsAccessDenied = document.getElementById("settings-access-denied");

let currentSettings;
let currentProfile;
let signatureTemplate = "";
let msalInstance;
let settingsProfile = null;
let cityChangeAuthorized = SignaturePreferences.getCityChangeAuthorized();

function setSettingsStatus(message) {
  settingsStatus.textContent = message;
}

function updateInternalInsertionVisibility() {
  skipInternalNewMailField.hidden = !skipInternalOnlyCheckbox.checked;
  skipInternalNewMailCheckbox.disabled = skipInternalOnlyCheckbox.disabled
    || !skipInternalOnlyCheckbox.checked;
}

function setControlsDisabled(disabled) {
  phoneModeSelect.disabled = disabled;
  cityChangeSelect.disabled = disabled || !cityChangeAuthorized;
  greetingModeSelect.disabled = disabled;
  customGreetingInput.disabled = disabled || greetingModeSelect.value !== "MfGCustom";
  greetingLinesSelect.disabled = disabled || greetingModeSelect.value === "MfG0";
  insertTitleBeforeCheckbox.disabled = disabled;
  insertTitleAfterCheckbox.disabled = disabled;
  mobileUsageCheckbox.disabled = disabled;
  mobileUsageTextInput.disabled = disabled || !mobileUsageCheckbox.checked;
  confidentialityCheckbox.disabled = disabled;
  autoInsertRepliesCheckbox.disabled = disabled;
  autoInsertForwardsCheckbox.disabled = disabled;
  autoInsertMeetingsCheckbox.disabled = disabled;
  skipInternalOnlyCheckbox.disabled = disabled;
  skipInternalNewMailCheckbox.disabled = disabled || !skipInternalOnlyCheckbox.checked;
}

function updateMobileUsageVisibility() {
  mobileUsageTextField.hidden = !mobileUsageCheckbox.checked;
  mobileUsageTextInput.disabled = mobileUsageCheckbox.disabled || !mobileUsageCheckbox.checked;
}

function updatePhoneWarnings() {
  if (!settingsProfile) {
    combinedPhoneWarning.hidden = true;
    mobilePhoneWarning.hidden = true;
    landlinePhoneWarning.hidden = true;
    return;
  }
  const mode = phoneModeSelect.value;
  const mobileMissing = !String(settingsProfile.mobile || "").trim();
  const phoneMissing = !String(settingsProfile.phone || "").trim();
  const bothNeeded = mode === "Alles";
  combinedPhoneWarning.hidden = !(bothNeeded && mobileMissing && phoneMissing);
  mobilePhoneWarning.hidden = !mobileMissing
    || (mode !== "Handy" && mode !== "Alles" && mode !== "EDVHotline")
    || (bothNeeded && phoneMissing);
  landlinePhoneWarning.hidden = !phoneMissing
    || (mode !== "Festnetz" && mode !== "Alles")
    || (bothNeeded && mobileMissing);
}

function updateGreetingVisibility() {
  const isCustom = greetingModeSelect.value === "MfGCustom";
  const hasGreeting = greetingModeSelect.value !== "MfG0";
  customGreetingField.hidden = !isCustom;
  greetingLinesField.hidden = !hasGreeting;
  customGreetingInput.disabled = greetingModeSelect.disabled || !isCustom;
  greetingLinesSelect.disabled = greetingModeSelect.disabled || !hasGreeting;
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

function showSettings(settings, department, titleAttributes) {
  if (titleAttributes && (Object.hasOwn(titleAttributes, "mobile") || Object.hasOwn(titleAttributes, "phone"))) {
    settingsProfile = titleAttributes;
  }
  const canUseEdvHotline = updateDepartmentOption(department);
  cityChangeAuthorized = SignaturePreferences.getCityChangeAuthorized();
  cityChangeField.hidden = !cityChangeAuthorized;
  cityChangeSelect.value = settings.CityOverride;
  phoneModeSelect.value = settings.Nummer === "EDVHotline" && !canUseEdvHotline
    ? "Alles"
    : settings.Nummer;
  updatePhoneWarnings();
  greetingModeSelect.value = settings.MfG;
  customGreetingInput.value = settings.CustomGreeting;
  greetingLinesSelect.value = String(settings.GreetingLines);
  updateGreetingVisibility();
  titleBeforeField.hidden = !titleAttributes.customAttribute10;
  titleAfterField.hidden = !titleAttributes.customAttribute11;
  insertTitleBeforeCheckbox.checked = settings.InsertTitleBefore;
  insertTitleAfterCheckbox.checked = settings.InsertTitleAfter;
  mobileUsageCheckbox.checked = settings.MobileUsage;
  mobileUsageTextInput.value = settings.MobileUsageText;
  updateMobileUsageVisibility();
  confidentialityCheckbox.checked = settings.Confidentiality;
  autoInsertRepliesCheckbox.checked = settings.AutoInsertReplies;
  autoInsertForwardsCheckbox.checked = settings.AutoInsertForwards;
  autoInsertMeetingsCheckbox.checked = settings.AutoInsertMeetings;
  skipInternalOnlyCheckbox.checked = settings.SkipInternalOnly;
  skipInternalNewMailCheckbox.checked = settings.SkipInternalOnNewMail;
  updateInternalInsertionVisibility();
}

async function acquireGraphToken() {
  if (!hasConfiguredEntraApp()) {
    throw new Error("Bitte Client-ID und Tenant-ID einmal im ATTENSAM_CONFIG-Block oben in taskpane.js eintragen.");
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
    const result = await msalInstance.acquireTokenSilent(request);
    rememberMobileCityChangeRole(result);
    return result.accessToken;
  } catch (error) {
    if (!(error instanceof msal.InteractionRequiredAuthError)) throw error;
    const result = await msalInstance.acquireTokenPopup(request);
    rememberMobileCityChangeRole(result);
    return result.accessToken;
  }
}

function rememberMobileCityChangeRole(result) {
  let tokenClaims = {};
  try {
    const encoded = String(result?.idToken || "").split(".")[1];
    if (encoded) {
      const normalized = encoded.replace(/-/g, "+").replace(/_/g, "/");
      tokenClaims = JSON.parse(atob(normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=")));
    }
  } catch {
    tokenClaims = {};
  }
  const roles = [
    ...(Array.isArray(result?.idTokenClaims?.roles) ? result.idTokenClaims.roles : []),
    ...(Array.isArray(result?.account?.idTokenClaims?.roles) ? result.account.idTokenClaims.roles : []),
    ...(Array.isArray(tokenClaims?.roles) ? tokenClaims.roles : []),
  ].map((role) => String(role).trim());
  cityChangeAuthorized = roles.includes("CityChange") || roles.includes("ATS.Signature.CityChange");
  SignaturePreferences.setAccessAuthorized(roles.includes("ATS.Signature"));
  SignaturePreferences.setCityChangeAuthorized(cityChangeAuthorized);
}

async function loadProfile() {
  const token = await acquireGraphToken();
  const select = [
    "id", "givenName", "surname", "displayName", "mail", "userPrincipalName",
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
    id: user.id || "",
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
  const standardSettings = await SignaturePreferences.getSettings();
  roamingSettings.set(AUTO_RENDER_DATA_KEY, {
    profile: { ...currentProfile },
    mailboxEmail: Office.context.mailbox.userProfile?.emailAddress || currentProfile.email,
    template: signatureTemplate,
    officeNumber: CONFIG.officeNumber,
    settings: { ...standardSettings },
    accessAuthorized: SignaturePreferences.getAccessAuthorized(),
    cityChangeAuthorized,
    settingsUpdatedAt: now,
    graphAuth: {
      clientId: CONFIG.clientId,
      tenantId: CONFIG.tenantId,
    },
    profileUpdatedAt: now,
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
    try {
      await acquireGraphToken();
    } catch (roleError) {
      console.warn("Die aktuelle Add-In-Rolle konnte mobil nicht überprüft werden.", roleError);
    }
    const accessAuthorized = SignaturePreferences.getAccessAuthorized();
    settingsMain.hidden = !accessAuthorized;
    settingsAccessDenied.hidden = accessAuthorized;
    if (!accessAuthorized) {
      setSettingsStatus("Sie haben kein Zugriff auf dieses Add-In, bitte EDV kontaktieren!");
      return;
    }
    const isVipUser = SignaturePreferences.getVipAuthorized();
    if (!isVipUser && MOBILE_SETTINGS_SIGNATURE_ID !== "standard") {
      throw new Error("Diese Signatur-Einstellungen sind nur für VIP-Benutzer verfügbar.");
    }
    const customRecord = await SignaturePreferences.getCustomSignatures();
    const selectedItem = customRecord.items.find((item) => item.id === MOBILE_SETTINGS_SIGNATURE_ID);
    if (MOBILE_SETTINGS_SIGNATURE_ID !== "standard" && !selectedItem) throw new Error("Die ausgewählte Signatur wurde nicht gefunden.");
    settingsHeading.textContent = !isVipUser
      ? "Einstellungen"
      : MOBILE_SETTINGS_SIGNATURE_ID === "standard"
      ? "Einstellungen: Standard"
      : `Einstellungen: ${selectedItem.title}`;
    currentSettings = await SignaturePreferences.getSettingsForSignature(MOBILE_SETTINGS_SIGNATURE_ID);
    settingsProfile = SignaturePreferences.getValidRenderData()?.profile || null;
    showSettings(
      currentSettings,
      SignaturePreferences.getDepartment(),
      SignaturePreferences.getTitleAttributes(),
    );
    setControlsDisabled(false);
  } catch (error) {
    setSettingsStatus(`Einstellungen konnten nicht geladen werden: ${readableError(error)}`);
    return;
  }

  setSettingsStatus("Einstellungen sind verfügbar. Signaturdaten werden geladen …");
  try {
    [signatureTemplate, currentProfile] = await Promise.all([
      fetch("template.html", { cache: "no-store" }).then((response) => {
        if (!response.ok) throw new Error("template.html konnte nicht geladen werden.");
        return response.text();
      }),
      loadProfile(),
    ]);
    SignaturePreferences.setDepartment(currentProfile.department);
    SignaturePreferences.setTitleAttributes(
      currentProfile.customAttribute10,
      currentProfile.customAttribute11,
    );
    showSettings(currentSettings, currentProfile.department, currentProfile);
    await saveAutomaticRenderData();
    setControlsDisabled(false);
    setSettingsStatus("Profil und Einstellungen sind bereit.");
  } catch (error) {
    console.error("Microsoft-365-Profil konnte mobil nicht aktualisiert werden.", error);
    setControlsDisabled(false);
    setSettingsStatus(`Einstellungen sind verfügbar; Profil konnte nicht aktualisiert werden: ${readableError(error)}`);
  }
}

async function saveSettings() {
  setControlsDisabled(true);
  setSettingsStatus("Einstellungen werden gespeichert …");
  try {
    currentSettings = await SignaturePreferences.saveSettingsForSignature(MOBILE_SETTINGS_SIGNATURE_ID, {
      Nummer: phoneModeSelect.value,
      MfG: greetingModeSelect.value,
      CustomGreeting: customGreetingInput.value,
      GreetingLines: Number(greetingLinesSelect.value),
      CityOverride: cityChangeAuthorized ? cityChangeSelect.value : currentSettings.CityOverride,
      AutoInsert: true,
      AutoInsertReplies: autoInsertRepliesCheckbox.checked,
      AutoInsertForwards: autoInsertForwardsCheckbox.checked,
      AutoInsertMeetings: autoInsertMeetingsCheckbox.checked,
      SkipInternalOnly: skipInternalOnlyCheckbox.checked,
      SkipInternalOnNewMail: skipInternalNewMailCheckbox.checked,
      InsertTitleBefore: insertTitleBeforeCheckbox.checked,
      InsertTitleAfter: insertTitleAfterCheckbox.checked,
      MobileUsage: mobileUsageCheckbox.checked,
      MobileUsageText: mobileUsageTextInput.value,
      Confidentiality: confidentialityCheckbox.checked,
    });
    setSettingsStatus("Einstellungen gespeichert.");
  } catch (error) {
    setSettingsStatus(error.message || "Einstellungen konnten nicht gespeichert werden.");
  } finally {
    setControlsDisabled(false);
  }
}

phoneModeSelect.addEventListener("change", () => {
  updatePhoneWarnings();
  saveSettings();
});
cityChangeSelect.addEventListener("change", saveSettings);
greetingModeSelect.addEventListener("change", () => {
  updateGreetingVisibility();
  saveSettings();
});
customGreetingInput.addEventListener("change", saveSettings);
greetingLinesSelect.addEventListener("change", saveSettings);
insertTitleBeforeCheckbox.addEventListener("change", saveSettings);
insertTitleAfterCheckbox.addEventListener("change", saveSettings);
mobileUsageCheckbox.addEventListener("change", () => {
  updateMobileUsageVisibility();
  saveSettings();
});
mobileUsageTextInput.addEventListener("change", saveSettings);
confidentialityCheckbox.addEventListener("change", saveSettings);
autoInsertRepliesCheckbox.addEventListener("change", saveSettings);
autoInsertForwardsCheckbox.addEventListener("change", saveSettings);
autoInsertMeetingsCheckbox.addEventListener("change", saveSettings);
skipInternalOnlyCheckbox.addEventListener("change", () => {
  updateInternalInsertionVisibility();
  saveSettings();
});
skipInternalNewMailCheckbox.addEventListener("change", saveSettings);
closeButton.addEventListener("click", () => {
  if (Office.context.ui?.closeContainer) Office.context.ui.closeContainer();
  else window.history.back();
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) initializeSettings();
  else setSettingsStatus("Diese Seite muss als Outlook-Add-In geöffnet werden.");
});

})();
