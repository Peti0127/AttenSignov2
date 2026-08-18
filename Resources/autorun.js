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
  const officeNumber = String(officeNumberValue || "").includes("YOUR_")
    ? ""
    : escapeHtml(String(officeNumberValue || "").trim());

  if (settings.Nummer === "Handy") return mobile ? `Mobil ${mobile}` : "";
  if (settings.Nummer === "Festnetz") return phone ? `Tel.: ${phone}` : "";
  if (settings.Nummer === "Office") return officeNumber ? `Office: ${officeNumber}` : "";
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
    CustomAttribute10: profile.customAttribute10,
    CustomAttribute11: profile.customAttribute11,
  };
  const signatureBody = renderData.template.replace(/\{([^{}]+)\}/g, (match, key) => {
    if (key === "Phone Mobile Office Number") {
      return phoneLine(profile, settings, renderData.officeNumber);
    }
    if (key === "Banner") return bannerForCity(profile.city);
    return Object.prototype.hasOwnProperty.call(values, key) ? escapeHtml(values[key]) : match;
  });
  return greetingHtml(settings) + signatureBody;
}

function getComposeType() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getComposeTypeAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value.composeType);
      else reject(result.error);
    });
  });
}

function setSignature(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(result.error);
      },
    );
  });
}

async function autoInsertSignature(event) {
  try {
    const settings = Office.context.roamingSettings.get(SETTINGS_KEY);
    if (!settings || settings.AutoInsert !== true) return;

    if (settings.AutoInsertMode !== "AllMail") {
      const composeType = await getComposeType();
      if (composeType !== "newMail") return;
    }

    const renderData = Office.context.roamingSettings.get(RENDER_DATA_KEY);
    if (!renderData || !renderData.profile || typeof renderData.template !== "string") return;
    await setSignature(renderSignature(renderData, settings));
  } catch (error) {
    console.error("Automatische Signatur konnte nicht eingefügt werden.", error);
  } finally {
    event.completed();
  }
}

Office.onReady();
Office.actions.associate("autoInsertSignature", autoInsertSignature);
