/* global Office, msal */

const CONFIG = {
  clientId: "89659501-37e7-4916-abeb-4dc5178e3034",
  tenantId: "https://login.microsoftonline.com/1333c2c2-fdf6-4fdc-8559-3dc12559d264",
};

const fields = [
  ["firstName", "Vorname"], ["lastName", "Nachname"],
  ["jobTitle", "Position"], ["company", "Unternehmen"],
  ["email", "E-Mail"], ["phone", "Telefon"], ["mobile", "Mobil"],
  ["street", "Straße"], ["postalCode", "PLZ"], ["city", "Ort"],
  ["customAttribute10", "CustomAttribute10"],
  ["customAttribute11", "CustomAttribute11"],
];

const profile = Object.fromEntries(fields.map(([key]) => [key, ""]));
let signatureTemplate = "";
let msalInstance;

const statusElement = document.getElementById("status");
const previewElement = document.getElementById("signature-preview");
const insertButton = document.getElementById("insert-button");
const refreshButton = document.getElementById("refresh-button");

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
  if (phone && mobile) return `Tel.: ${phone}&nbsp;&nbsp;Mobil: ${mobile}`;
  if (mobile) return `Mobil ${mobile}`;
  if (phone) return `Tel.: ${phone}`;
  return "";
}

function bannerForCity() {
  const city = profile.city.trim().toLocaleLowerCase("de-AT");
  if (city === "wien") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://ASD.com/banner_w" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://ASD.com/images/banner_w.png" border="0" alt="Banner Wien"></a></p>';
  }
  if (city === "st. pölten-radlberg" || city === "krems an der donau") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://ASD.com/banner_noe_nord" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://ASD.com/images/banner_noe_nord.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  return "";
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
  const html = signatureTemplate.replace(/\{([^{}]+)\}/g, (match, key) => {
    if (key === "Phone Mobile Office Number") return phoneLine();
    if (key === "Banner") return bannerForCity();
    if (key === "MfG") return match;
    return Object.hasOwn(values, key) ? escapeHtml(values[key]) : match;
  });
  previewElement.innerHTML = html;
  insertButton.disabled = !html;
  return html;
}

function buildFields() {
  const container = document.getElementById("fields");
  fields.forEach(([key, label]) => {
    const wrapper = document.createElement("label");
    wrapper.innerHTML = `<span>${label}</span>`;
    const input = document.createElement("input");
    input.id = key;
    input.autocomplete = "off";
    input.addEventListener("input", () => {
      profile[key] = input.value;
      renderSignature();
    });
    wrapper.appendChild(input);
    container.appendChild(wrapper);
  });
}

function showProfile() {
  fields.forEach(([key]) => {
    document.getElementById(key).value = profile[key] || "";
  });
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
    msalInstance = await msal.createNestablePublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
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
  refreshButton.disabled = true;
  setStatus("Microsoft-365-Profil wird geladen …");
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
    showProfile();
    setStatus("Profildaten wurden geladen.");
  } catch (error) {
    setStatus(error.message || "Profildaten konnten nicht geladen werden.");
  } finally {
    refreshButton.disabled = false;
  }
}

function insertSignature() {
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
  buildFields();
  try {
    signatureTemplate = await fetch("template.html", { cache: "no-store" }).then((response) => {
      if (!response.ok) throw new Error("template.html konnte nicht geladen werden.");
      return response.text();
    });
    applyMailboxBasics();
    setStatus("Bereit. Laden Sie das vollständige Microsoft-365-Profil.");
  } catch (error) {
    setStatus(error.message);
  }
}

refreshButton.addEventListener("click", loadProfile);
insertButton.addEventListener("click", insertSignature);
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) initialize();
  else setStatus("Diese Seite muss als Outlook-Add-In geöffnet werden.");
});
