/* global Office, msal, SignaturePreferences */

const CONFIG = {
  clientId: "89659501-37e7-4916-abeb-4dc5178e3034",
  tenantId: "https://login.microsoftonline.com/1333c2c2-fdf6-4fdc-8559-3dc12559d264",
  officeNumber: "05 7999 100",
};
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
  InsertTitleBefore: false,
  InsertTitleAfter: false,
  MobileUsage: false,
  Confidentiality: false,
};

const statusElement = document.getElementById("status");
const previewElement = document.getElementById("signature-preview");
const signatureButton = document.getElementById("signature-button");

function setStatus(message) {
  statusElement.textContent = message;
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;").replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;").replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function phoneLine() {
  const phone = escapeHtml(profile.phone.trim());
  const mobile = escapeHtml(profile.mobile.trim());
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
  return `<p style="margin: 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">${escapeHtml(greeting)}<br><br></p>`;
}

function noticesHtml() {
  let html = "";
  if (signatureSettings.MobileUsage) {
    html += '<p style="margin: 12px 0 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Diese E-Mail wurde über Outlook Mobile versendet.</p>';
  }
  if (signatureSettings.Confidentiality) {
    html += '<p style="margin: 6px 0 0; font-family: Aptos, Arial, sans-serif; font-size: 9pt; color: rgb(0, 0, 0);">Diese E-Mail ist vertraulich.</p>';
  }
  return html;
}

function normalizeEmail(value) {
  return String(value || "").trim().toLocaleLowerCase("de-AT");
}

function delegationName(delegation) {
  if (!delegation) return "";
  const titleBefore = String(delegation.customAttribute10 || "").trim();
  const titleAfter = String(delegation.customAttribute11 || "").trim();
  const personalName = [delegation.firstName, delegation.lastName]
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .join(" ") || String(delegation.displayName || "").trim();
  return [titleBefore, personalName, titleAfter].filter(Boolean).join(" ");
}

function delegationHtml() {
  const name = delegationName(currentDelegation);
  if (!name) return "";
  return `<p style="margin: 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Im Auftrag von <b>${escapeHtml(name)}</b><br><br></p>`;
}

function bannerForCity() {
  const city = profile.city;
  if (city === "Wien") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_w" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_w.png" border="0" alt="Banner Wien"></a></p>';
  }
  if (city === "st. pölten-radlberg" || city === "Krems an der Donau") {
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
  const titleBefore = signatureSettings.InsertTitleBefore && profile.customAttribute10.trim()
    ? `${profile.customAttribute10.trim()} `
    : "";
  const titleAfter = signatureSettings.InsertTitleAfter && profile.customAttribute11.trim()
    ? ` ${profile.customAttribute11.trim()}`
    : "";
  const values = {
    FirstName: profile.firstName, LastName: profile.lastName,
    Company: profile.company, City: profile.city, Street: profile.street,
    PostalCode: profile.postalCode, JobTitle: profile.jobTitle,
    "E-mail": profile.email, Mobile: profile.mobile, Phone: profile.phone,
    CustomAttribute10: titleBefore,
    CustomAttribute11: titleAfter,
  };
  const signatureBody = signatureTemplate.replace(/\{([^{}]+)\}/g, (match, key) => {
    if (key === "Phone Mobile Office Number") return phoneLine();
    if (key === "Banner") return bannerForCity();
    return Object.hasOwn(values, key) ? escapeHtml(values[key]) : match;
  });
  const signatureContent = greetingHtml() + delegationHtml() + signatureBody + noticesHtml();
  const html = `<div id="${SIGNATURE_MARKER_ID}" data-attensam-signature="v2"><span style="display:none!important;mso-hide:all;max-height:0;overflow:hidden;font-size:0;line-height:0;color:transparent;">${SIGNATURE_MARKER_TEXT}</span>${signatureContent}</div>`;
  previewElement.innerHTML = html;
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
  if (CONFIG.clientId.includes("YOUR_")) {
    throw new Error("Bitte die Entra Client-ID oben in taskpane.js eintragen.");
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
    customAttribute10: "",
    customAttribute11: "",
  };
  try {
    const token = await acquireGraphToken(["User.Read.All"]);
    const select = "id,displayName,givenName,surname,mail,userPrincipalName,onPremisesExtensionAttributes";
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
