/* global Office */

const SETTINGS_KEY = "attensam.signature.settings.v2";
const RENDER_DATA_KEY = "attensam.signature.render-data.v1";

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

function renderSignature(renderData, settings) {
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
  return greetingHtml(settings) + signatureBody + noticesHtml(settings);
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

function insertCachedSignature(event, settings, renderData) {
  let html;
  try {
    if (!renderData || !renderData.profile || typeof renderData.template !== "string") {
      completeEvent(event);
      return;
    }
    html = renderSignature(renderData, settings);
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
      insertCachedSignature(event, settings, renderData);
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
      insertCachedSignature(event, settings, renderData);
    });
  } catch (error) {
    console.error("Automatische Signatur konnte nicht eingefügt werden.", error);
    completeEvent(event);
  }
}

Office.onReady();
Office.actions.associate("autoInsertSignature", autoInsertSignature);
