/* global Office, msal */

const SETTINGS_KEY = "attensam.signature.settings.v2";
const RENDER_DATA_KEY = "attensam.signature.render-data.v1";
const SIGNATURE_MARKER_ID = "attensam-signature-root";
const SIGNATURE_MARKER_TEXT = "ATTENSAM-SIGNATURE-V2";
let eventMsalInstance;

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function phoneLine(profile, settings, officeNumberValue) {
  const phone = escapeHtml(String(profile.phone || "").trim());
  const mobile = escapeHtml(String(profile.mobile || "").trim());
  let officeNumber = "";
  if (!String(officeNumberValue || "").includes("YOUR_")) {
    officeNumber = escapeHtml(String(officeNumberValue || "").trim());
  }

  if (settings.Nummer === "Handy") {
    if (mobile) return `Mobil ${mobile}`;
    return "";
  }
  if (settings.Nummer === "Festnetz") {
    if (phone) return `Tel.: ${phone}`;
    return "";
  }
  if (settings.Nummer === "Office") {
    if (officeNumber) return `Office: ${officeNumber}`;
    return "";
  }
  // The settings page already restricts this value to users in department IT.
  // Don't re-check an older cached profile here, because it may predate the
  // addition of the department field.
  if (settings.Nummer === "EDVHotline") {
    if (mobile) return `Tel. 05 7999 9999 Mobil ${mobile}`;
    return "Tel. 05 7999 9999";
  }
  if (phone && mobile) return `Tel.: ${phone}&nbsp;&nbsp;Mobil: ${mobile}`;
  if (mobile) return `Mobil ${mobile}`;
  if (phone) return `Tel.: ${phone}`;
  return "";
}

function greetingHtml(settings) {
  let greeting = "";
  if (settings.MfG === "MfG1") greeting = "Mit freundlichen Grüßen";
  else if (settings.MfG === "MfG2") greeting = "Freundliche Grüße";
  else if (settings.MfG === "MfG3") greeting = "LG";
  else if (settings.MfG === "MfGCustom") greeting = String(settings.CustomGreeting || "").trim();
  if (!greeting) return "";
  return `<p style="margin: 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">${escapeHtml(greeting)}<br><br></p>`;
}

function noticesHtml(settings) {
  let html = "";
  if (settings.MobileUsage === true) {
    html += '<p style="margin: 12px 0 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Diese E-Mail wurde über Outlook Mobile versendet.</p>';
  }
  if (settings.Confidentiality === true) {
    html += '<p style="margin: 6px 0 0; font-family: Aptos, Arial, sans-serif; font-size: 9pt; color: rgb(0, 0, 0);">Diese E-Mail ist vertraulich.</p>';
  }
  return html;
}

function normalizeEmail(value) {
  return String(value || "").trim().toLocaleLowerCase("de-AT");
}

function delegationName(delegation) {
  let personalName = "";
  const parts = [];
  if (!delegation) return "";
  if (delegation.firstName) personalName = String(delegation.firstName).trim();
  if (delegation.lastName) {
    if (personalName) personalName += " ";
    personalName += String(delegation.lastName).trim();
  }
  if (!personalName) personalName = String(delegation.displayName || "").trim();
  if (String(delegation.customAttribute10 || "").trim()) {
    parts.push(String(delegation.customAttribute10).trim());
  }
  if (personalName) parts.push(personalName);
  if (String(delegation.customAttribute11 || "").trim()) {
    parts.push(String(delegation.customAttribute11).trim());
  }
  return parts.join(" ");
}

function delegationHtml(delegation) {
  const name = delegationName(delegation);
  if (!name) return "";
  return `<p style="margin: 0; font-family: Aptos, Arial, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">Im Auftrag von <b>${escapeHtml(name)}</b><br><br></p>`;
}

function bannerForCity(cityValue) {
  const city = String(cityValue || "").trim();
  if (city === "Wien") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_w" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_w.png" border="0" alt="Banner Wien"></a></p>';
  }
  if (city === "St. Pölten-Radlberg" || city === "Krems an der Donau") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_noe_nord" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_noe_nord.png" border="0" alt="Banner Niederösterreich Nord"></a></p>';
  }
  if (city === "Wr. Neustadt") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_noe_sued" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_noe_sued.png" border="0" alt="Banner Niederösterreich Süd"></a></p>';
  }
  if (city === "Neusiedl am See" || city === "Oberwart") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_bgld" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_bgld.png" border="0" alt="Banner Burgenland"></a></p>';
  }
  if (city === "Klagenfurt") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_ktn" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_ktn.png" border="0" alt="Banner Kärnten"></a></p>';
  }
  if (city === "Kalsdorf" || city === "Graz" || city === "Leoben") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_stmk" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_stmk.png" border="0" alt="Banner Steiermark"></a></p>';
  }
  if (city === "Linz" || city === "Regau") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_ooe" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_ooe.png" border="0" alt="Banner Oberösterreich"></a></p>';
  }
  if (city === "Salzburg" || city === "Bruck an der Großglocknerstraße") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_sbg" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_sbg.png" border="0" alt="Banner Salzburg"></a></p>';
  }
  if (city === "Innsbruck") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_t" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_t.png" border="0" alt="Banner Tirol"></a></p>';
  }
  if (city === "Rankweil") {
    return '<p style="font-size: 12pt; font-family: Aptos, Arial, sans-serif; color: rgb(0, 0, 0);"><a href="https://www.attensam.at/banner_vbg" title="" style="font-family: Arial; font-size: 10pt;"><img src="https://storage.googleapis.com/signaturen-attensam-at/images/banner_vbg.png" border="0" alt="Banner Vorarlberg"></a></p>';
  }
  return "";
}

function renderSignature(renderData, settings, delegation) {
  const profile = renderData.profile;
  let titleBefore = "";
  let titleAfter = "";
  if (settings.InsertTitleBefore === true && String(profile.customAttribute10 || "").trim()) {
    titleBefore = `${String(profile.customAttribute10).trim()} `;
  }
  if (settings.InsertTitleAfter === true && String(profile.customAttribute11 || "").trim()) {
    titleAfter = ` ${String(profile.customAttribute11).trim()}`;
  }
  const values = {
    FirstName: profile.firstName,
    LastName: profile.lastName,
    Company: profile.company,
    City: profile.city,
    Street: profile.street,
    PostalCode: profile.postalCode,
    JobTitle: profile.jobTitle,
    "E-mail": profile.email,
    Mobile: profile.mobile,
    Phone: profile.phone,
    CustomAttribute10: titleBefore,
    CustomAttribute11: titleAfter,
  };
  const signatureBody = renderData.template.replace(/\{([^{}]+)\}/g, (match, key) => {
    if (key === "Phone Mobile Office Number") {
      return phoneLine(profile, settings, renderData.officeNumber);
    }
    if (key === "Banner") return bannerForCity(profile.city);
    if (Object.prototype.hasOwnProperty.call(values, key)) return escapeHtml(values[key]);
    return match;
  });
  const signatureContent = greetingHtml(settings) + delegationHtml(delegation) + signatureBody + noticesHtml(settings);
  return `<div id="${SIGNATURE_MARKER_ID}" data-attensam-signature="v2"><span style="display:none!important;mso-hide:all;max-height:0;overflow:hidden;font-size:0;line-height:0;color:transparent;">${SIGNATURE_MARKER_TEXT}</span>${signatureContent}</div>`;
}

function completeEvent(event) {
  event.completed();
}

function settingsTime(value) {
  const parsed = Date.parse(value || "");
  if (Number.isFinite(parsed)) return parsed;
  return 0;
}

function synchronizedSettings(savedSettings, renderData) {
  let result = savedSettings;
  if (renderData && renderData.settings) {
    if (
      !result
      || settingsTime(renderData.settingsUpdatedAt) >= settingsTime(result.updatedAt)
    ) {
      result = renderData.settings;
    }
  }
  return result;
}

function acquireEventGraphToken(renderData) {
  const graphAuth = renderData && renderData.graphAuth;
  if (
    typeof msal === "undefined"
    || !graphAuth
    || !graphAuth.clientId
    || String(graphAuth.clientId).includes("YOUR_")
  ) {
    return Promise.reject(new Error("Graph-Konfiguration für den Ereignisruntime fehlt."));
  }
  let authority = String(graphAuth.tenantId || "");
  if (!authority.startsWith("https://")) {
    authority = `https://login.microsoftonline.com/${authority}`;
  }
  let initialize;
  if (eventMsalInstance) {
    initialize = Promise.resolve(eventMsalInstance);
  } else {
    initialize = msal.createNestablePublicClientApplication({
      auth: { clientId: graphAuth.clientId, authority },
      cache: { cacheLocation: "localStorage" },
    }).then(function rememberInstance(instance) {
      eventMsalInstance = instance;
      return instance;
    });
  }
  return initialize.then(function acquireToken(instance) {
    return instance.acquireTokenSilent({ scopes: ["User.Read.All"] });
  }).then(function tokenValue(result) {
    return result.accessToken;
  });
}

function fetchDelegatedUser(renderData, fromDetails) {
  return acquireEventGraphToken(renderData).then(function requestDelegatedUser(token) {
    const select = "id,displayName,givenName,surname,mail,userPrincipalName,onPremisesExtensionAttributes";
    const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromDetails.emailAddress)}?$select=${encodeURIComponent(select)}`;
    return fetch(url, { headers: { Authorization: `Bearer ${token}` } }).then(function directResult(response) {
      if (response.ok) return response.json();
      const address = String(fromDetails.emailAddress || "").replace(/'/g, "''");
      const filter = `mail eq '${address}' or userPrincipalName eq '${address}'`;
      const fallbackUrl = `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(filter)}&$select=${encodeURIComponent(select)}&$top=1`;
      return fetch(fallbackUrl, { headers: { Authorization: `Bearer ${token}` } }).then(function fallbackResult(fallbackResponse) {
        if (!fallbackResponse.ok) throw new Error(`Microsoft Graph: ${fallbackResponse.status}`);
        return fallbackResponse.json().then(function firstUser(payload) {
          if (!payload.value || !payload.value[0]) throw new Error("Microsoft Graph: Benutzer nicht gefunden");
          return payload.value[0];
        });
      });
    });
  }).then(function mapDelegatedUser(user) {
    const attributes = user.onPremisesExtensionAttributes || {};
    return {
      displayName: user.displayName || fromDetails.displayName || fromDetails.emailAddress || "",
      id: user.id || "",
      firstName: user.givenName || "",
      lastName: user.surname || "",
      email: user.mail || user.userPrincipalName || fromDetails.emailAddress || "",
      customAttribute10: attributes.extensionAttribute10 || "",
      customAttribute11: attributes.extensionAttribute11 || "",
    };
  });
}

function resolveDelegation(renderData, callback) {
  const from = Office.context.mailbox.item.from;
  if (!from || !from.getAsync) {
    callback(null);
    return;
  }
  from.getAsync(function onFromRead(result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
      callback(null);
      return;
    }
    const fromDetails = result.value;
    const fromEmail = normalizeEmail(fromDetails.emailAddress);
    const ownEmail = normalizeEmail(renderData && renderData.profile && renderData.profile.email);
    if (!fromEmail || fromEmail === ownEmail) {
      callback(null);
      return;
    }
    const fallback = {
      displayName: fromDetails.displayName || fromDetails.emailAddress || "",
      id: "",
      firstName: "",
      lastName: "",
      email: fromDetails.emailAddress || "",
      customAttribute10: "",
      customAttribute11: "",
    };
    fetchDelegatedUser(renderData, fromDetails).then(function resolved(user) {
      if (
        user.id
        && renderData.profile.id
        && user.id === renderData.profile.id
      ) {
        callback(null);
        return;
      }
      callback(user);
    }).catch(function fallbackToFrom(error) {
      console.error("Titel der abweichenden Absenderadresse konnten nicht geladen werden.", error);
      callback(fallback);
    });
  });
}

function insertCachedSignature(event, settings, renderData, delegation) {
  let html;
  try {
    if (!renderData || !renderData.profile || typeof renderData.template !== "string") {
      completeEvent(event);
      return;
    }
    html = renderSignature(renderData, settings, delegation);
  } catch (error) {
    console.error("Automatische Signatur konnte nicht erstellt werden.", error);
    completeEvent(event);
    return;
  }

  try {
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function onSignatureSet(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Automatische Signatur konnte nicht eingefügt werden.", result.error);
        }
        completeEvent(event);
      }
    );
  } catch (error) {
    console.error("Automatische Signatur konnte nicht eingefügt werden.", error);
    completeEvent(event);
  }
}

function resolveAndInsertSignature(event, settings, renderData) {
  resolveDelegation(renderData, function onDelegationResolved(delegation) {
    insertCachedSignature(event, settings, renderData, delegation);
  });
}

function autoInsertSignature(event) {
  let settings;
  let renderData;
  try {
    renderData = Office.context.roamingSettings.get(RENDER_DATA_KEY);
    settings = synchronizedSettings(
      Office.context.roamingSettings.get(SETTINGS_KEY),
      renderData
    );
    if (!settings || settings.AutoInsert !== true) {
      completeEvent(event);
      return;
    }

    if (settings.AutoInsertMode === "AllMail") {
      resolveAndInsertSignature(event, settings, renderData);
      return;
    }

    Office.context.mailbox.item.getComposeTypeAsync(function onComposeType(result) {
      if (
        result.status !== Office.AsyncResultStatus.Succeeded
        || !result.value
        || result.value.composeType !== "newMail"
      ) {
        completeEvent(event);
        return;
      }
      resolveAndInsertSignature(event, settings, renderData);
    });
  } catch (error) {
    console.error("Automatische Signatur konnte nicht eingefügt werden.", error);
    completeEvent(event);
  }
}

function updateSignatureForFrom(event) {
  let settings;
  let renderData;
  try {
    renderData = Office.context.roamingSettings.get(RENDER_DATA_KEY);
    settings = synchronizedSettings(
      Office.context.roamingSettings.get(SETTINGS_KEY),
      renderData
    );
    if (!settings || settings.AutoInsert !== true) {
      completeEvent(event);
      return;
    }
    if (settings.AutoInsertMode === "AllMail") {
      resolveAndInsertSignature(event, settings, renderData);
      return;
    }
    Office.context.mailbox.item.getComposeTypeAsync(function onComposeType(result) {
      if (
        result.status !== Office.AsyncResultStatus.Succeeded
        || !result.value
        || result.value.composeType !== "newMail"
      ) {
        completeEvent(event);
        return;
      }
      resolveAndInsertSignature(event, settings, renderData);
    });
  } catch (error) {
    console.error("Signatur konnte nach dem Absenderwechsel nicht aktualisiert werden.", error);
    completeEvent(event);
  }
}

if (typeof window !== "undefined") {
  window.AttensamSignatureRuntime = Object.freeze({
    renderSignature,
    resolveDelegation,
  });
}

Office.onReady();
Office.actions.associate("autoInsertSignature", autoInsertSignature);
Office.actions.associate("updateSignatureForFrom", updateSignatureForFrom);
