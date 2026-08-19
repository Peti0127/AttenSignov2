/* global Office */

(function exposeSignaturePreferences(global) {
  const ROAMING_KEY = "attensam.signature.settings.v2";
  const RENDER_DATA_KEY = "attensam.signature.render-data.v1";
  const CACHE_PREFIX = "attensam.signature.settings.v2";
  const DEPARTMENT_CACHE_PREFIX = "attensam.signature.department.v1";
  const LEGACY_PHONE_PREFIX = "attensam.signature.phone-mode";
  const DEFAULT_SETTINGS = Object.freeze({
    Nummer: "Alles",
    MfG: "MfG1",
    AutoInsert: false,
    AutoInsertMode: "NewMail",
  });
  const ALLOWED_NUMBERS = new Set(["Alles", "Handy", "Festnetz", "Office", "EDVHotline"]);
  const ALLOWED_GREETINGS = new Set(["MfG0", "MfG1", "MfG2", "MfG3"]);
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

  function normalizeRecord(value) {
    if (!value || typeof value !== "object") return null;
    return {
      Nummer: ALLOWED_NUMBERS.has(value?.Nummer) ? value.Nummer : DEFAULT_SETTINGS.Nummer,
      MfG: ALLOWED_GREETINGS.has(value?.MfG) ? value.MfG : DEFAULT_SETTINGS.MfG,
      AutoInsert: value.AutoInsert === true,
      AutoInsertMode: ALLOWED_AUTO_MODES.has(value?.AutoInsertMode)
        ? value.AutoInsertMode
        : DEFAULT_SETTINGS.AutoInsertMode,
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
      AutoInsert: record.AutoInsert,
      AutoInsertMode: record.AutoInsertMode,
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
      AutoInsert: DEFAULT_SETTINGS.AutoInsert,
      AutoInsertMode: DEFAULT_SETTINGS.AutoInsertMode,
    };
  }

  async function saveSettings(settings) {
    if (
      !ALLOWED_NUMBERS.has(settings?.Nummer)
      || !ALLOWED_GREETINGS.has(settings?.MfG)
      || typeof settings?.AutoInsert !== "boolean"
      || !ALLOWED_AUTO_MODES.has(settings?.AutoInsertMode)
    ) {
      throw new Error("Ungültige Einstellung.");
    }
    const record = {
      Nummer: settings.Nummer,
      MfG: settings.MfG,
      AutoInsert: settings.AutoInsert,
      AutoInsertMode: settings.AutoInsertMode,
      updatedAt: new Date().toISOString(),
    };
    const roamingSettings = Office.context.roamingSettings;
    if (!roamingSettings) throw new Error("Outlook RoamingSettings ist nicht verfügbar.");

    roamingSettings.set(ROAMING_KEY, record);
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
  });
})(window);
