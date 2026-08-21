/* Attensam compact UI bundle. Generated from the tested UI modules. */
const ATTENSAM_CONFIG = Object.freeze({
  clientId: "89659501-37e7-4916-abeb-4dc5178e3034",
  tenantId: "https://login.microsoftonline.com/1333c2c2-fdf6-4fdc-8559-3dc12559d264",
  officeNumber: "05 7999 100",
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
/* global Office */

(function exposeSignaturePreferences(global) {
  const ROAMING_KEY = "attensam.signature.settings.v2";
  const RENDER_DATA_KEY = "attensam.signature.render-data.v1";
  const CACHE_PREFIX = "attensam.signature.settings.v2";
  const DEPARTMENT_CACHE_PREFIX = "attensam.signature.department.v1";
  const TITLE_ATTRIBUTES_CACHE_PREFIX = "attensam.signature.title-attributes.v1";
  const LEGACY_PHONE_PREFIX = "attensam.signature.phone-mode";
  const DEFAULT_SETTINGS = Object.freeze({
    Nummer: "Alles",
    MfG: "MfG1",
    CustomGreeting: "",
    GreetingLines: 1,
    AutoInsert: false,
    AutoInsertMode: "NewMail",
    InsertTitleBefore: false,
    InsertTitleAfter: false,
    MobileUsage: false,
    Confidentiality: false,
  });
  const ALLOWED_NUMBERS = new Set(["Alles", "Handy", "Festnetz", "Office", "EDVHotline"]);
  const ALLOWED_GREETINGS = new Set(["MfG0", "MfG1", "MfG2", "MfG3", "MfGCustom"]);
  const ALLOWED_GREETING_LINES = new Set([1, 2, 3]);
  const ALLOWED_AUTO_MODES = new Set(["NewMail", "AllMail"]);
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

  function storageKey() {
    return `${CACHE_PREFIX}:${currentUserKey()}`;
  }

  function departmentStorageKey() {
    return `${DEPARTMENT_CACHE_PREFIX}:${currentUserKey()}`;
  }

  function titleAttributesStorageKey() {
    return `${TITLE_ATTRIBUTES_CACHE_PREFIX}:${currentUserKey()}`;
  }

  function normalizeCustomGreeting(value) {
    return String(value || "").replace(/\s+/g, " ").trim().slice(0, 200);
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
    return {
      Nummer: ALLOWED_NUMBERS.has(value?.Nummer) ? value.Nummer : DEFAULT_SETTINGS.Nummer,
      MfG: ALLOWED_GREETINGS.has(value?.MfG) ? value.MfG : DEFAULT_SETTINGS.MfG,
      CustomGreeting: normalizeCustomGreeting(value?.CustomGreeting),
      GreetingLines: ALLOWED_GREETING_LINES.has(Number(value?.GreetingLines))
        ? Number(value.GreetingLines)
        : DEFAULT_SETTINGS.GreetingLines,
      AutoInsert: value.AutoInsert === true,
      AutoInsertMode: ALLOWED_AUTO_MODES.has(value?.AutoInsertMode)
        ? value.AutoInsertMode
        : DEFAULT_SETTINGS.AutoInsertMode,
      InsertTitleBefore: value.InsertTitleBefore === true,
      InsertTitleAfter: value.InsertTitleAfter === true,
      MobileUsage: value.MobileUsage === true,
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
      AutoInsert: record.AutoInsert,
      AutoInsertMode: record.AutoInsertMode,
      InsertTitleBefore: record.InsertTitleBefore,
      InsertTitleAfter: record.InsertTitleAfter,
      MobileUsage: record.MobileUsage,
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
      AutoInsert: DEFAULT_SETTINGS.AutoInsert,
      AutoInsertMode: DEFAULT_SETTINGS.AutoInsertMode,
      InsertTitleBefore: DEFAULT_SETTINGS.InsertTitleBefore,
      InsertTitleAfter: DEFAULT_SETTINGS.InsertTitleAfter,
      MobileUsage: DEFAULT_SETTINGS.MobileUsage,
      Confidentiality: DEFAULT_SETTINGS.Confidentiality,
    };
  }

  async function saveSettings(settings) {
    if (
      !ALLOWED_NUMBERS.has(settings?.Nummer)
      || !ALLOWED_GREETINGS.has(settings?.MfG)
      || typeof settings?.CustomGreeting !== "string"
      || !ALLOWED_GREETING_LINES.has(Number(settings?.GreetingLines))
      || typeof settings?.AutoInsert !== "boolean"
      || !ALLOWED_AUTO_MODES.has(settings?.AutoInsertMode)
      || typeof settings?.InsertTitleBefore !== "boolean"
      || typeof settings?.InsertTitleAfter !== "boolean"
      || typeof settings?.MobileUsage !== "boolean"
      || typeof settings?.Confidentiality !== "boolean"
    ) {
      throw new Error("Ungültige Einstellung.");
    }
    const record = {
      Nummer: settings.Nummer,
      MfG: settings.MfG,
      CustomGreeting: normalizeCustomGreeting(settings.CustomGreeting),
      GreetingLines: Number(settings.GreetingLines),
      AutoInsert: settings.AutoInsert,
      AutoInsertMode: settings.AutoInsertMode,
      InsertTitleBefore: settings.InsertTitleBefore,
      InsertTitleAfter: settings.InsertTitleAfter,
      MobileUsage: settings.MobileUsage,
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

  global.SignaturePreferences = Object.freeze({
    getSettings,
    saveSettings,
    getDepartment,
    setDepartment,
    getTitleAttributes,
    setTitleAttributes,
  });
})(window);

(function compactRoute(){
  if (!(new URLSearchParams(window.location.search).get("view") !== "settings")) return;
/* global Office, msal, SignaturePreferences */

const CONFIG = ATTENSAM_CONFIG;
const AUTO_RENDER_DATA_KEY = "attensam.signature.render-data.v1";
const SIGNATURE_MARKER_ID = "attensam-signature-root";
const SIGNATURE_MARKER_TEXT = "ATTENSAM-SIGNATURE-V2";

const profile = {
  id: "", firstName: "", lastName: "", jobTitle: "", company: "",
  email: "", phone: "", mobile: "", street: "",
  postalCode: "", city: "", department: "", customAttribute10: "",
  customAttribute11: "",
};
let signatureTemplate = "";
let msalInstance;
let profileLoaded = false;
let currentDelegation = null;
let signatureSettings = {
  Nummer: "Alles",
  MfG: "MfG1",
  CustomGreeting: "",
  GreetingLines: 1,
  InsertTitleBefore: false,
  InsertTitleAfter: false,
  MobileUsage: false,
  Confidentiality: false,
};

const statusElement = document.getElementById("status");
const previewElement = document.getElementById("signature-preview");
const signatureButton = document.getElementById("signature-button");
const profileWarningsElement = document.getElementById("profile-warnings");

function setStatus(message) {
  statusElement.textContent = message;
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;").replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;").replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function phoneLine(profileValue = profile) {
  const phone = escapeHtml(String(profileValue.phone || "").trim());
  const mobile = escapeHtml(String(profileValue.mobile || "").trim());
  const officeNumber = CONFIG.officeNumber.includes("YOUR_")
    ? ""
    : escapeHtml(CONFIG.officeNumber.trim());

  switch (signatureSettings.Nummer) {
    case "Handy":
      return mobile ? `Mobil ${mobile}` : "";
    case "Festnetz":
      return phone ? `Tel.: ${phone}` : "";
    case "Office":
      return officeNumber ? `Office: ${officeNumber}` : "";
    case "EDVHotline":
      if (profile.department.trim().toLocaleUpperCase("de-AT") === "IT") {
        return mobile ? `Tel. 05 7999 9999 Mobil ${mobile}` : "Tel. 05 7999 9999";
      }
      // If the department changed, fall back to the standard phone line.
    default:
      if (phone && mobile) return `Tel.: ${phone}&nbsp;&nbsp;Mobil: ${mobile}`;
      if (mobile) return `Mobil ${mobile}`;
      if (phone) return `Tel.: ${phone}`;
      return "";
  }
}

function greetingHtml() {
  if (signatureSettings.MfG === "MfG0") return "";
  let greeting = "";
  if (signatureSettings.MfG === "MfG1") {
    greeting = "Mit freundlichen Grüßen";
  } else if (signatureSettings.MfG === "MfG2") {
    greeting = "Freundliche Grüße";
  } else if (signatureSettings.MfG === "MfG3") {
    greeting = "LG";
  } else if (signatureSettings.MfG === "MfGCustom") {
    greeting = String(signatureSettings.CustomGreeting || "").trim();
  }
  if (!greeting) return "";
  const configuredLines = Number(signatureSettings.GreetingLines);
  const blankLines = [1, 2, 3].includes(configuredLines) ? configuredLines : 1;
  return `<p style="margin: 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">${escapeHtml(greeting)}${"<br>".repeat(blankLines + 1)}</p>`;
}

function isOutlookMobile() {
  const platform = Office.context?.platform;
  const platformTypes = Office.PlatformType || {};
  return platform === platformTypes.Android || platform === platformTypes.iOS;
}

function noticesHtml() {
  let html = "";
  if (signatureSettings.MobileUsage && isOutlookMobile()) {
    html += '<p style="margin: 12px 0 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Diese E-Mail wurde über Outlook Mobile versendet.</p>';
  }
  if (signatureSettings.Confidentiality) {
    html += '<p style="margin: 6px 0 0; font-family: Aptos, Arial, sans-serif; font-size: 9pt; color: rgb(0, 0, 0);">Diese E-Mail ist vertraulich.</p>';
  }
  return html;
}

function missingProfileFields(profileValue) {
  const requiredFields = [
    ["firstName", "Vorname"],
    ["lastName", "Nachname"],
    ["jobTitle", "Funktion"],
    ["email", "E-Mail-Adresse"],
    ["city", "Ort"],
    ["postalCode", "Postleitzahl"],
    ["street", "Straße"],
  ];
  return requiredFields
    .filter(([key]) => !String(profileValue?.[key] || "").trim())
    .map(([, label]) => label);
}

function profileWarningMessages(profileValue, phoneMode = null) {
  const messages = missingProfileFields(profileValue)
    .map((field) => `Information über ${field} fehlt, bitte EDV kontaktieren!`);
  const mobileMissing = !String(profileValue?.mobile || "").trim();
  const phoneMissing = !String(profileValue?.phone || "").trim();
  const needsMobile = phoneMode === null || phoneMode === "Handy" || phoneMode === "Alles";
  const needsPhone = phoneMode === null || phoneMode === "Festnetz" || phoneMode === "Alles";
  if (mobileMissing && phoneMissing && needsMobile && needsPhone) {
    messages.push("Informationen über Festnetznummer und Mobilnummer fehlen, bitte EDV kontaktieren!");
  } else {
    if (mobileMissing && needsMobile) {
      messages.push("Information über Mobilnummer fehlt, bitte EDV kontaktieren!");
    }
    if (phoneMissing && needsPhone) {
      messages.push("Information über Festnetznummer fehlt, bitte EDV kontaktieren!");
    }
  }
  return messages;
}

function insertedProfileWarningsHtml(profileValue) {
  return profileWarningMessages(profileValue, signatureSettings.Nummer)
    .map((message) => `<p style="margin: 0 0 6px; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: #c00000; font-weight: bold;"><b>${escapeHtml(message)}</b></p>`)
    .join("");
}

function showProfileWarnings(profileValue) {
  if (!profileLoaded) {
    profileWarningsElement.replaceChildren();
    profileWarningsElement.hidden = true;
    return;
  }
  const messages = profileWarningMessages(profileValue, signatureSettings.Nummer);
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

function delegatedName(profileValue) {
  const titleBefore = signatureSettings.InsertTitleBefore
    ? String(profileValue?.customAttribute10 || "").trim()
    : "";
  const titleAfter = signatureSettings.InsertTitleAfter
    ? String(profileValue?.customAttribute11 || "").trim()
    : "";
  return [titleBefore, personalName(profileValue), titleAfter].filter(Boolean).join(" ");
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

function scaleSignaturePreview() {
  // Reset before measuring so the previous scale doesn't affect the result.
  previewElement.style.transform = "none";

  const styles = getComputedStyle(signatureButton);
  const availableWidth = signatureButton.clientWidth
    - parseFloat(styles.paddingLeft)
    - parseFloat(styles.paddingRight);
  const availableHeight = signatureButton.clientHeight
    - parseFloat(styles.paddingTop)
    - parseFloat(styles.paddingBottom);
  const naturalWidth = previewElement.scrollWidth;
  const naturalHeight = previewElement.scrollHeight;

  if (!naturalWidth || !naturalHeight) return;

  const scale = Math.min(
    1,
    availableWidth / naturalWidth,
    availableHeight / naturalHeight,
  );
  previewElement.style.transform = `scale(${scale})`;
}

function renderSignature() {
  const signatureProfile = currentDelegation || profile;
  const titleBefore = !currentDelegation
    && signatureSettings.InsertTitleBefore && String(signatureProfile.customAttribute10 || "").trim()
    ? `${String(signatureProfile.customAttribute10).trim()} `
    : "";
  const titleAfter = !currentDelegation
    && signatureSettings.InsertTitleAfter && String(signatureProfile.customAttribute11 || "").trim()
    ? ` ${String(signatureProfile.customAttribute11).trim()}`
    : "";
  const senderName = personalName(profile);
  const fromName = delegatedName(currentDelegation);
  const values = {
    FirstName: currentDelegation ? senderName : signatureProfile.firstName,
    LastName: currentDelegation ? `(i.A. ${fromName})` : signatureProfile.lastName,
    Company: signatureProfile.company, City: signatureProfile.city, Street: signatureProfile.street,
    PostalCode: signatureProfile.postalCode, JobTitle: signatureProfile.jobTitle,
    "E-mail": signatureProfile.email, Mobile: signatureProfile.mobile, Phone: signatureProfile.phone,
    CustomAttribute10: titleBefore,
    CustomAttribute11: titleAfter,
  };
  const signatureBody = signatureTemplate.replace(/\{([^{}]+)\}/g, (match, key) => {
    if (key === "Phone Mobile Office Number") return phoneLine(signatureProfile);
    if (key === "Banner") return bannerForCity(signatureProfile);
    return Object.hasOwn(values, key) ? escapeHtml(values[key]) : match;
  });
  const signatureContent = greetingHtml() + signatureBody + noticesHtml();
  const marker = `<span style="display:none!important;mso-hide:all;max-height:0;overflow:hidden;font-size:0;line-height:0;color:transparent;">${SIGNATURE_MARKER_TEXT}</span>`;
  const previewHtml = `<div id="${SIGNATURE_MARKER_ID}" data-attensam-signature="v2">${marker}${signatureContent}</div>`;
  const html = `<div id="${SIGNATURE_MARKER_ID}" data-attensam-signature="v2">${marker}${insertedProfileWarningsHtml(signatureProfile)}${signatureContent}</div>`;
  previewElement.innerHTML = previewHtml;
  showProfileWarnings(signatureProfile);
  previewElement.querySelectorAll("img").forEach((image) => {
    if (!image.complete) image.addEventListener("load", scaleSignaturePreview, { once: true });
  });
  requestAnimationFrame(scaleSignaturePreview);
  const ready = Boolean(html && profileLoaded);
  signatureButton.setAttribute("aria-disabled", String(!ready));
  signatureButton.tabIndex = ready ? 0 : -1;
  signatureButton.classList.toggle("ready", ready);
  return html;
}

async function saveAutoRenderData() {
  const roamingSettings = Office.context.roamingSettings;
  if (!roamingSettings || !signatureTemplate || !profileLoaded) return;
  roamingSettings.set(AUTO_RENDER_DATA_KEY, {
    profile: { ...profile },
    template: signatureTemplate,
    officeNumber: CONFIG.officeNumber,
    settings: { ...signatureSettings },
    settingsUpdatedAt: new Date().toISOString(),
    graphAuth: {
      clientId: CONFIG.clientId,
      tenantId: CONFIG.tenantId,
    },
    updatedAt: new Date().toISOString(),
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
}

function applyMailboxBasics() {
  const mailbox = Office.context.mailbox.userProfile;
  const names = (mailbox.displayName || "").trim().split(/\s+/);
  profile.firstName = names[0] || "";
  profile.lastName = names.length > 1 ? names[names.length - 1] : "";
  profile.email = mailbox.emailAddress || "";
  showProfile();
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
    return (await msalInstance.acquireTokenSilent(request)).accessToken;
  } catch (error) {
    if (!(error instanceof msal.InteractionRequiredAuthError)) throw error;
    return (await msalInstance.acquireTokenPopup(request)).accessToken;
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
    } else {
      const address = String(fromDetails.emailAddress || "").replaceAll("'", "''");
      const filter = `mail eq '${address}' or userPrincipalName eq '${address}'`;
      response = await fetch(
        `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(filter)}&$select=${encodeURIComponent(select)}&$top=1`,
        { headers: { Authorization: `Bearer ${token}` } },
      );
      if (!response.ok) return fallback;
      const result = await response.json();
      user = result.value?.[0];
      if (!user) return fallback;
    }
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
    return;
  }
  const fromDomain = emailDomain(fromEmail);
  const ownDomains = new Set([
    emailDomain(profile.email),
    emailDomain(Office.context.mailbox.userProfile.emailAddress),
  ].filter(Boolean));
  if (!fromDomain || !ownDomains.has(fromDomain)) {
    currentDelegation = null;
    return;
  }
  currentDelegation = await loadDelegatedUser(fromDetails);
  if (currentDelegation.id && profile.id && currentDelegation.id === profile.id) {
    currentDelegation = null;
  }
}

async function loadProfile() {
  setStatus("Microsoft-365-Profil wird automatisch geladen …");
  try {
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
    SignaturePreferences.setDepartment(profile.department);
    SignaturePreferences.setTitleAttributes(profile.customAttribute10, profile.customAttribute11);
    await refreshDelegationForCurrentFrom();
    profileLoaded = true;
    showProfile();
    try {
      await saveAutoRenderData();
    } catch (cacheError) {
      setStatus(`Profil geladen, aber automatische Signaturdaten konnten nicht gespeichert werden: ${cacheError.message}`);
      return;
    }
    setStatus("Microsoft-365-Profil wurde automatisch geladen.");
  } catch (error) {
    profileLoaded = false;
    signatureButton.setAttribute("aria-disabled", "true");
    signatureButton.tabIndex = -1;
    signatureButton.classList.remove("ready");
    setStatus(error.message || "Profildaten konnten nicht geladen werden.");
  }
}

async function insertSignature() {
  if (!profileLoaded || signatureButton.getAttribute("aria-disabled") === "true") return;
  await refreshDelegationForCurrentFrom();
  const html = renderSignature();
  const body = Office.context.mailbox.item?.body;
  if (!body) {
    setStatus("Bitte eine neue Nachricht öffnen.");
    return;
  }
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
    signatureTemplate = await fetch("template.html", { cache: "no-store" }).then((response) => {
      if (!response.ok) throw new Error("template.html konnte nicht geladen werden.");
      return response.text();
    });
    signatureSettings = await SignaturePreferences.getSettings();
    applyMailboxBasics();
    await loadProfile();
  } catch (error) {
    setStatus(error.message);
  }
}

signatureButton.addEventListener("click", insertSignature);
signatureButton.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    insertSignature();
  }
});
window.addEventListener("resize", scaleSignaturePreview);
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) initialize();
  else setStatus("Diese Seite muss als Outlook-Add-In geöffnet werden.");
});

})();

(function compactRoute(){
  if (!(new URLSearchParams(window.location.search).get("view") === "settings" && new URLSearchParams(window.location.search).get("mobile") !== "1")) return;
/* global Office, SignaturePreferences, AttensamSignatureRuntime, DOMParser */

const AUTO_RENDER_DATA_KEY = "attensam.signature.render-data.v1";
const SETTINGS_SIGNATURE_MARKER_ID = "attensam-signature-root";
const SETTINGS_SIGNATURE_MARKER_TEXT = "ATTENSAM-SIGNATURE-V2";

const phoneModeSelect = document.getElementById("phone-mode");
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
const confidentialityCheckbox = document.getElementById("confidentiality");
const autoInsertCheckbox = document.getElementById("auto-insert");
const autoInsertModeField = document.getElementById("auto-insert-mode-field");
const autoInsertModeSelect = document.getElementById("auto-insert-mode");
const settingsStatus = document.getElementById("settings-status");
let currentSettings;
let settingsProfile = null;

function setSettingsStatus(message) {
  settingsStatus.textContent = message;
}

function updateAutoInsertVisibility() {
  autoInsertModeField.hidden = !autoInsertCheckbox.checked;
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
  greetingModeSelect.disabled = disabled;
  customGreetingInput.disabled = disabled || greetingModeSelect.value !== "MfGCustom";
  greetingLinesSelect.disabled = disabled || greetingModeSelect.value === "MfG0";
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
    settingsProfile = Office.context.roamingSettings?.get(AUTO_RENDER_DATA_KEY)?.profile || null;
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
      GreetingLines: Number(greetingLinesSelect.value),
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

phoneModeSelect.addEventListener("change", () => {
  updatePhoneWarnings();
  saveSettings();
});
greetingModeSelect.addEventListener("change", () => {
  updateGreetingVisibility();
  saveSettings();
});
customGreetingInput.addEventListener("change", saveSettings);
greetingLinesSelect.addEventListener("change", saveSettings);
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

})();

(function compactRoute(){
  if (!(new URLSearchParams(window.location.search).get("view") === "settings" && new URLSearchParams(window.location.search).get("mobile") === "1")) return;
/* global Office, msal, SignaturePreferences */

const CONFIG = ATTENSAM_CONFIG;
const AUTO_RENDER_DATA_KEY = "attensam.signature.render-data.v1";

const phoneModeSelect = document.getElementById("phone-mode");
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
const confidentialityCheckbox = document.getElementById("confidentiality");
const autoInsertCheckbox = document.getElementById("auto-insert");
const autoInsertModeField = document.getElementById("auto-insert-mode-field");
const autoInsertModeSelect = document.getElementById("auto-insert-mode");
const settingsStatus = document.getElementById("settings-status");
const closeButton = document.getElementById("close-button");

let currentSettings;
let currentProfile;
let signatureTemplate = "";
let msalInstance;
let settingsProfile = null;

function setSettingsStatus(message) {
  settingsStatus.textContent = message;
}

function setControlsDisabled(disabled) {
  phoneModeSelect.disabled = disabled;
  greetingModeSelect.disabled = disabled;
  customGreetingInput.disabled = disabled || greetingModeSelect.value !== "MfGCustom";
  greetingLinesSelect.disabled = disabled || greetingModeSelect.value === "MfG0";
  insertTitleBeforeCheckbox.disabled = disabled;
  insertTitleAfterCheckbox.disabled = disabled;
  mobileUsageCheckbox.disabled = disabled;
  confidentialityCheckbox.disabled = disabled;
  autoInsertCheckbox.disabled = disabled;
  autoInsertModeSelect.disabled = disabled;
}

function updateAutoInsertVisibility() {
  autoInsertModeField.hidden = !autoInsertCheckbox.checked;
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
  confidentialityCheckbox.checked = settings.Confidentiality;
  autoInsertCheckbox.checked = settings.AutoInsert;
  autoInsertModeSelect.value = settings.AutoInsertMode;
  updateAutoInsertVisibility();
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
    return (await msalInstance.acquireTokenSilent(request)).accessToken;
  } catch (error) {
    if (!(error instanceof msal.InteractionRequiredAuthError)) throw error;
    return (await msalInstance.acquireTokenPopup(request)).accessToken;
  }
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
  roamingSettings.set(AUTO_RENDER_DATA_KEY, {
    profile: { ...currentProfile },
    template: signatureTemplate,
    officeNumber: CONFIG.officeNumber,
    settings: { ...currentSettings },
    settingsUpdatedAt: now,
    graphAuth: {
      clientId: CONFIG.clientId,
      tenantId: CONFIG.tenantId,
    },
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
    settingsProfile = Office.context.roamingSettings?.get(AUTO_RENDER_DATA_KEY)?.profile || null;
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

  setSettingsStatus("Einstellungen sind verfügbar. Microsoft-365-Profil wird geladen …");
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
    currentSettings = await SignaturePreferences.saveSettings({
      Nummer: phoneModeSelect.value,
      MfG: greetingModeSelect.value,
      CustomGreeting: customGreetingInput.value,
      GreetingLines: Number(greetingLinesSelect.value),
      AutoInsert: autoInsertCheckbox.checked,
      AutoInsertMode: autoInsertModeSelect.value,
      InsertTitleBefore: insertTitleBeforeCheckbox.checked,
      InsertTitleAfter: insertTitleAfterCheckbox.checked,
      MobileUsage: mobileUsageCheckbox.checked,
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
greetingModeSelect.addEventListener("change", () => {
  updateGreetingVisibility();
  saveSettings();
});
customGreetingInput.addEventListener("change", saveSettings);
greetingLinesSelect.addEventListener("change", saveSettings);
insertTitleBeforeCheckbox.addEventListener("change", saveSettings);
insertTitleAfterCheckbox.addEventListener("change", saveSettings);
mobileUsageCheckbox.addEventListener("change", saveSettings);
confidentialityCheckbox.addEventListener("change", saveSettings);
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

})();
