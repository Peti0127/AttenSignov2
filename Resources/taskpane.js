/* global Office, msal, SignaturePreferences */

const CONFIG = {
  clientId: "89659501-37e7-4916-abeb-4dc5178e3034",
  tenantId: "https://login.microsoftonline.com/1333c2c2-fdf6-4fdc-8559-3dc12559d264",
  officeNumber: "05 7999 100",
};
const AUTO_RENDER_DATA_KEY = "attensam.signature.render-data.v1";

const profile = {
  firstName: "", lastName: "", jobTitle: "", company: "",
  email: "", phone: "", mobile: "", street: "",
  postalCode: "", city: "", customAttribute10: "",
  customAttribute11: "",
};
let signatureTemplate = "";
let msalInstance;
let profileLoaded = false;
let signatureSettings = { Nummer: "Alles", MfG: "MfG1" };

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
  }
  if (!greeting) return "";
  return `<p style="margin: 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">${escapeHtml(greeting)}<br><br></p>`;
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
  const values = {
    FirstName: profile.firstName, LastName: profile.lastName,
    Company: profile.company, City: profile.city, Street: profile.street,
    PostalCode: profile.postalCode, JobTitle: profile.jobTitle,
    "E-mail": profile.email, Mobile: profile.mobile, Phone: profile.phone,
    CustomAttribute10: profile.customAttribute10,
    CustomAttribute11: profile.customAttribute11,
  };
  const signatureBody = signatureTemplate.replace(/\{([^{}]+)\}/g, (match, key) => {
    if (key === "Phone Mobile Office Number") return phoneLine();
    if (key === "Banner") return bannerForCity();
    return Object.hasOwn(values, key) ? escapeHtml(values[key]) : match;
  });
  const html = greetingHtml() + signatureBody;
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

async function acquireGraphToken() {
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
  const request = { scopes: ["User.Read"] };
  try {
    return (await msalInstance.acquireTokenSilent(request)).accessToken;
  } catch (error) {
    if (!(error instanceof msal.InteractionRequiredAuthError)) throw error;
    return (await msalInstance.acquireTokenPopup(request)).accessToken;
  }
}

async function loadProfile() {
  setStatus("Microsoft-365-Profil wird automatisch geladen …");
  try {
    const token = await acquireGraphToken();
    const select = [
      "givenName", "surname", "displayName", "mail", "userPrincipalName",
      "companyName", "city", "streetAddress", "postalCode", "jobTitle",
      "mobilePhone", "businessPhones", "onPremisesExtensionAttributes",
    ].join(",");
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me?$select=${encodeURIComponent(select)}`,
      { headers: { Authorization: `Bearer ${token}` } },
    );
    if (!response.ok) throw new Error(`Microsoft Graph: ${response.status}`);
    const user = await response.json();
    Object.assign(profile, {
      firstName: user.givenName || "",
      lastName: user.surname || "",
      email: user.mail || user.userPrincipalName || profile.email,
      company: user.companyName || "",
      city: user.city || "",
      street: user.streetAddress || "",
      postalCode: user.postalCode || "",
      jobTitle: user.jobTitle || "",
      mobile: user.mobilePhone || "",
      phone: user.businessPhones?.[0] || "",
      customAttribute10: user.onPremisesExtensionAttributes?.extensionAttribute10 || "",
      customAttribute11: user.onPremisesExtensionAttributes?.extensionAttribute11 || "",
    });
    profileLoaded = true;
    showProfile();
    await saveAutoRenderData().catch(() => {});
    setStatus("Microsoft-365-Profil wurde automatisch geladen.");
  } catch (error) {
    profileLoaded = false;
    signatureButton.setAttribute("aria-disabled", "true");
    signatureButton.tabIndex = -1;
    signatureButton.classList.remove("ready");
    setStatus(error.message || "Profildaten konnten nicht geladen werden.");
  }
}

function insertSignature() {
  if (!profileLoaded || signatureButton.getAttribute("aria-disabled") === "true") return;
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
