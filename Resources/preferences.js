/* global Office */

(function exposeSignaturePreferences(global) {
  const STORAGE_PREFIX = "attensam.signature.settings.v2";
  const LEGACY_PHONE_PREFIX = "attensam.signature.phone-mode";
  const DEFAULT_SETTINGS = Object.freeze({ Nummer: "Alles", MfG: "MfG1" });
  const ALLOWED_NUMBERS = new Set(["Alles", "Handy", "Festnetz", "Office"]);
  const ALLOWED_GREETINGS = new Set(["MfG1", "MfG2", "MfG3"]);
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
    return `${STORAGE_PREFIX}:${currentUserKey()}`;
  }

  function normalizeSettings(value) {
    return {
      Nummer: ALLOWED_NUMBERS.has(value?.Nummer) ? value.Nummer : DEFAULT_SETTINGS.Nummer,
      MfG: ALLOWED_GREETINGS.has(value?.MfG) ? value.MfG : DEFAULT_SETTINGS.MfG,
    };
  }

  async function getSettings() {
    // Temporary adapter. Replace this block with GET /api/user-settings when the
    // authenticated SQL API is available. The API should return { Nummer, MfG }.
    const stored = localStorage.getItem(storageKey());
    if (stored) {
      try {
        return normalizeSettings(JSON.parse(stored));
      } catch {
        localStorage.removeItem(storageKey());
      }
    }

    // Preserve an existing phone choice from the previous add-in version.
    const legacy = localStorage.getItem(`${LEGACY_PHONE_PREFIX}:${currentUserKey()}`);
    return normalizeSettings({ Nummer: LEGACY_NUMBER_MAP[legacy], MfG: DEFAULT_SETTINGS.MfG });
  }

  async function saveSettings(settings) {
    if (!ALLOWED_NUMBERS.has(settings?.Nummer) || !ALLOWED_GREETINGS.has(settings?.MfG)) {
      throw new Error("Ungültige Einstellung.");
    }
    // Temporary adapter. Replace this line with PUT /api/user-settings and send
    // the same { Nummer, MfG } object when the authenticated SQL API is available.
    const normalized = normalizeSettings(settings);
    localStorage.setItem(storageKey(), JSON.stringify(normalized));
    return normalized;
  }

  global.SignaturePreferences = Object.freeze({ getSettings, saveSettings });
})(window);
