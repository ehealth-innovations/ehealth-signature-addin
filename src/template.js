/*
 * Ehealth Innovations e-mail signature template.
 * Table-based HTML with inline styles only (no classes), web-safe fonts.
 * Column grid: 100 | 22 | 1 | 22 | 230 = 375 px (same as the original design).
 *
 * buildSignature(person, assetBase, variant)
 *   person:   { name, title, phone, email }  – empty fields are left out entirely
 *   assetBase: absolute https URL of the hosted add-in, without trailing slash
 *   variant:  "full" (logo + details) or "compact" (text only, for replies)
 */

const INK = "#2B1A27";
const PLUM = "#8E0E5F";
const MAGENTA = "#C8007A";
const SLATE = "#5E5260";
const MIST = "#E8DDE5";
const LINKEDIN = "https://www.linkedin.com/company/ehealth-innovations/";

function esc(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function clean(value) {
  return String(value == null ? "" : value).replace(/\s+/g, " ").trim();
}

/** "+31 6 12 34 56 78" / "06-12345678" / "0031 6..." -> "+31612345678" */
export function telHref(phone) {
  let nr = clean(phone).replace(/[^\d+]/g, "");
  if (nr.indexOf("00") === 0) nr = "+" + nr.slice(2);
  else if (nr.charAt(0) === "0") nr = "+31" + nr.slice(1);
  return nr.length > 4 ? "tel:" + nr : "";
}

const TD_TEXT =
  "padding:0;margin:0;font-family:Arial,Helvetica,sans-serif;font-size:10pt;line-height:18px;mso-line-height-rule:exactly;";

function spacerRow(assetBase, height, colspan, width) {
  return (
    `<tr><td colspan="${colspan}" height="${height}" style="height:${height}px;padding:0;margin:0;font-size:1px;line-height:1px;">` +
    `<img src="${assetBase}/assets/spacer.gif" width="${width}" height="${height}" alt="" style="display:block;width:${width}px;height:${height}px;border:0;"></td></tr>`
  );
}

function labelRow(label, valueHtml) {
  return (
    `<tr>` +
    `<td width="22" height="18" align="left" valign="top" style="width:22px;height:18px;${TD_TEXT}color:${MAGENTA};font-weight:bold;">${label}</td>` +
    `<td width="208" height="18" align="left" valign="top" style="width:208px;height:18px;${TD_TEXT}color:${INK};">${valueHtml}</td>` +
    `</tr>`
  );
}

function detailRows(p, assetBase) {
  const rows = [];
  rows.push(
    `<tr><td colspan="2" height="20" align="left" valign="top" style="padding:0;margin:0;height:20px;font-family:Arial,Helvetica,sans-serif;font-size:11pt;line-height:20px;mso-line-height-rule:exactly;color:${PLUM};font-weight:bold;">${esc(p.name)}</td></tr>`
  );
  if (p.title) {
    rows.push(
      `<tr><td colspan="2" height="18" align="left" valign="top" style="height:18px;${TD_TEXT}color:${SLATE};font-weight:normal;">${esc(p.title)}</td></tr>`
    );
  }
  if (p.phone || p.email) rows.push(spacerRow(assetBase, 14, 2, 1));
  if (p.phone) {
    const href = telHref(p.phone);
    const text = esc(p.phone);
    rows.push(
      labelRow("T", href ? `<a href="${esc(href)}" style="color:${INK};text-decoration:none;">${text}</a>` : text)
    );
  }
  if (p.email) {
    rows.push(
      labelRow(
        "E",
        `<a href="mailto:${esc(p.email)}" style="color:${INK};text-decoration:none;">${esc(p.email)}</a>`
      )
    );
  }
  rows.push(spacerRow(assetBase, 12, 2, 1));
  rows.push(
    `<tr><td colspan="2" height="18" align="left" valign="top" style="height:18px;${TD_TEXT}"><a href="${LINKEDIN}" style="color:${MAGENTA};text-decoration:underline;">Volg ons op LinkedIn</a></td></tr>`
  );
  return rows.join("");
}

function fullSignature(p, a) {
  return (
    `<table cellpadding="0" cellspacing="0" border="0" width="375" style="width:375px;border:0;border-collapse:collapse;border-spacing:0;mso-table-lspace:0pt;mso-table-rspace:0pt;font-family:Arial,Helvetica,sans-serif;font-size:10pt;line-height:18px;color:${INK};">` +
    `<tr>` +
    `<td width="100" valign="middle" style="width:100px;padding:0;margin:0;font-size:1px;line-height:1px;"><img src="${a}/assets/logo.png" width="100" height="145" alt="Ehealth Innovations" style="display:block;width:100px;height:145px;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic;"></td>` +
    `<td width="22" style="width:22px;padding:0;margin:0;font-size:1px;line-height:1px;">&nbsp;</td>` +
    `<td width="1" bgcolor="${MIST}" style="width:1px;padding:0;margin:0;font-size:1px;line-height:1px;background-color:${MIST};"><img src="${a}/assets/spacer.gif" width="1" height="145" alt="" style="display:block;width:1px;height:145px;border:0;"></td>` +
    `<td width="22" style="width:22px;padding:0;margin:0;font-size:1px;line-height:1px;">&nbsp;</td>` +
    `<td width="230" valign="middle" style="width:230px;padding:0;margin:0;">` +
    `<table cellpadding="0" cellspacing="0" border="0" width="230" style="width:230px;border:0;border-collapse:collapse;border-spacing:0;mso-table-lspace:0pt;mso-table-rspace:0pt;">` +
    detailRows(p, a) +
    `</table></td></tr>` +
    spacerRow(a, 16, 5, 375) +
    `<tr><td colspan="5" height="4" style="height:4px;padding:0;margin:0;font-size:1px;line-height:1px;"><img src="${a}/assets/band.png" width="375" height="4" alt="" style="display:block;width:375px;height:4px;border:0;"></td></tr>` +
    `</table>`
  );
}

function compactSignature(p, a) {
  return (
    `<table cellpadding="0" cellspacing="0" border="0" width="230" style="width:230px;border:0;border-collapse:collapse;border-spacing:0;mso-table-lspace:0pt;mso-table-rspace:0pt;font-family:Arial,Helvetica,sans-serif;font-size:10pt;line-height:18px;color:${INK};">` +
    detailRows(p, a) +
    `</table>`
  );
}

export function buildSignature(person, assetBase, variant) {
  const p = {
    name: clean(person.name) || clean(person.email),
    title: clean(person.title),
    phone: clean(person.phone),
    email: clean(person.email),
  };
  const a = String(assetBase).replace(/\/+$/, "");
  return variant === "compact" ? compactSignature(p, a) : fullSignature(p, a);
}
