/*
 * Ehealth Innovations – centrally managed Outlook signature.
 *
 * Runs in every Outlook client:
 *  - web, new Outlook for Windows, Mac, iOS, Android: loaded by commands.html
 *  - classic Outlook for Windows: this file is loaded directly (JS-only runtime,
 *    no DOM, no window.localStorage) – keep it DOM-free.
 *
 * Flow per compose event:
 *   profile (roamingSettings cache → Graph /me via NAA → Office userProfile fallback)
 *   → pick template variant by compose type → disable the user's own signature
 *   → setSignatureAsync → event.completed()
 */

/* global Office, fetch, console, __CONFIG__, __ASSETS__ */

import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { buildSignature } from "./template";

const CONFIG = __CONFIG__; // injected at build time from config.json
const ASSETS = __ASSETS__; // { logo, band } base64 PNGs, injected at build time from assets/
const SESSION_KEY = "ehi_sig_attachments";
const CACHE_KEY = "ehi_sig_profile_v1";
const GRAPH_URL =
  "https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,department,businessPhones,mobilePhone,mail";
const GRAPH_TIMEOUT_MS = 10000;

let pcaPromise = null;

Office.onReady(() => {});

function log(msg) {
  try {
    console.log("[ehi-signature] " + msg);
  } catch (e) {
    /* no console in some runtimes */
  }
}

function withTimeout(promise, ms, label) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error(label + " timed out")), ms);
    promise.then(
      (v) => {
        clearTimeout(t);
        resolve(v);
      },
      (e) => {
        clearTimeout(t);
        reject(e);
      }
    );
  });
}

/* ---------- Office async helpers (promise wrappers, feature-detected) ---------- */

function officeCall(fn) {
  return new Promise((resolve, reject) => {
    fn((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(result.error || new Error("Office call failed"));
    });
  });
}

function getComposeType(item) {
  if (typeof item.getComposeTypeAsync !== "function") return Promise.resolve("newMail");
  return officeCall((cb) => item.getComposeTypeAsync(cb))
    .then((v) => (v && v.composeType) || "newMail")
    .catch(() => "newMail");
}

function getFromAddress(item) {
  if (!item.from || typeof item.from.getAsync !== "function") return Promise.resolve(null);
  return officeCall((cb) => item.from.getAsync(cb))
    .then((v) => (v && v.emailAddress) || null)
    .catch(() => null);
}

function disableClientSignature(item) {
  if (typeof item.disableClientSignatureAsync !== "function") return Promise.resolve();
  return officeCall((cb) => item.disableClientSignatureAsync(cb)).catch((e) =>
    log("disableClientSignatureAsync failed: " + (e && e.message))
  );
}

function setSignature(item, html) {
  return officeCall((cb) =>
    item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, cb)
  );
}

/* ---------- Profile: cache → Graph → fallback ---------- */

function readCache() {
  try {
    const raw = Office.context.roamingSettings && Office.context.roamingSettings.get(CACHE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function writeCache(profile) {
  try {
    const rs = Office.context.roamingSettings;
    if (!rs) return Promise.resolve();
    rs.set(CACHE_KEY, JSON.stringify({ ts: Date.now(), profile: profile }));
    return new Promise((resolve) => rs.saveAsync(() => resolve()));
  } catch (e) {
    return Promise.resolve();
  }
}

function naaSupported() {
  try {
    return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
  } catch (e) {
    return false;
  }
}

async function getGraphToken() {
  if (!pcaPromise) {
    pcaPromise = createNestablePublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: "https://login.microsoftonline.com/" + CONFIG.tenantId,
      },
    });
  }
  const pca = await pcaPromise;
  const result = await pca.acquireTokenSilent({ scopes: ["User.Read"] });
  return result.accessToken;
}

async function fetchGraphProfile() {
  const token = await withTimeout(getGraphToken(), GRAPH_TIMEOUT_MS, "token");
  const res = await withTimeout(
    fetch(GRAPH_URL, { headers: { Authorization: "Bearer " + token } }),
    GRAPH_TIMEOUT_MS,
    "graph"
  );
  if (!res.ok) throw new Error("Graph HTTP " + res.status);
  const me = await res.json();
  const phone = me.mobilePhone || (me.businessPhones && me.businessPhones[0]) || "";
  return {
    name: me.displayName || "",
    title: me.jobTitle || "",
    department: me.department || "",
    phone: phone,
    email: me.mail || "",
  };
}

function fallbackProfile() {
  const up = (Office.context.mailbox && Office.context.mailbox.userProfile) || {};
  return { name: up.displayName || "", title: "", department: "", phone: "", email: up.emailAddress || "" };
}

async function getProfile() {
  const cached = readCache();
  const maxAge = (CONFIG.cacheHours || 24) * 3600 * 1000;
  if (cached && cached.profile && Date.now() - cached.ts < maxAge) return cached.profile;

  if (naaSupported()) {
    try {
      const profile = await fetchGraphProfile();
      await writeCache(profile);
      return profile;
    } catch (e) {
      log("Graph lookup failed: " + (e && e.message));
    }
  } else {
    log("NestedAppAuth 1.1 not supported in this client");
  }
  if (cached && cached.profile) return cached.profile; // stale beats nothing
  return fallbackProfile();
}

/* ---------- Embedded (inline) images ---------- */

/** config.inlineImages: "all", "none", or a list of mailbox addresses (for testing). */
function inlineImagesEnabled() {
  const v = CONFIG.inlineImages;
  if (v === "all") return true;
  if (Array.isArray(v)) {
    const me = String((Office.context.mailbox.userProfile || {}).emailAddress || "").toLowerCase();
    return v.some((x) => String(x).toLowerCase() === me);
  }
  return false;
}

function sessionGet(item) {
  if (!item.sessionData || typeof item.sessionData.getAsync !== "function") return Promise.resolve(null);
  return officeCall((cb) => item.sessionData.getAsync(SESSION_KEY, cb)).catch(() => null);
}

function sessionSet(item, value) {
  if (!item.sessionData || typeof item.sessionData.setAsync !== "function") return Promise.resolve();
  return officeCall((cb) => item.sessionData.setAsync(SESSION_KEY, value, cb)).catch(() => {});
}

function removeAttachment(item, id) {
  if (typeof item.removeAttachmentAsync !== "function") return Promise.resolve();
  return officeCall((cb) => item.removeAttachmentAsync(id, cb)).catch(() => {});
}

/** Removes the images this add-in attached earlier to the same message (From change / button). */
async function removePreviousImages(item) {
  const raw = await sessionGet(item);
  if (!raw) return;
  let ids = [];
  try {
    ids = JSON.parse(raw) || [];
  } catch (e) {
    ids = [];
  }
  for (const id of ids) await removeAttachment(item, id);
  await sessionSet(item, "[]");
}

/**
 * Attaches logo + band as inline images and returns their cid: sources.
 * Unique names per message, so they never clash with images in quoted replies.
 * Returns null (→ hosted images) when the client can't attach inline images.
 */
async function attachInlineImages(item) {
  if (!inlineImagesEnabled() || typeof item.addFileAttachmentFromBase64Async !== "function") return null;
  await removePreviousImages(item);
  const tag = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const names = { logo: "ehi-logo-" + tag + ".png", band: "ehi-band-" + tag + ".png" };
  const ids = [];
  try {
    for (const key of ["logo", "band"]) {
      const id = await officeCall((cb) =>
        item.addFileAttachmentFromBase64Async(ASSETS[key], names[key], { isInline: true }, cb)
      );
      ids.push(id);
    }
  } catch (e) {
    log("Inline image failed, using hosted images: " + (e && e.message));
    for (const id of ids) await removeAttachment(item, id);
    return null;
  }
  await sessionSet(item, JSON.stringify(ids));
  return { logo: "cid:" + names.logo, band: "cid:" + names.band };
}

/* ---------- Main ---------- */

async function applySignature(event) {
  const item = Office.context.mailbox.item;
  try {
    const [profile, composeType, fromAddress] = await Promise.all([
      getProfile(),
      getComposeType(item),
      getFromAddress(item),
    ]);

    // Shared mailbox / send-as: keep the person's details, show the address it is sent from.
    const person = {
      name: profile.name,
      title: profile.title,
      phone: profile.phone,
      email: fromAddress || profile.email,
    };

    const isReply = composeType === "reply" || composeType === "forward";
    const variant = isReply ? CONFIG.replyVariant || "full" : "full";
    await disableClientSignature(item);
    const images = variant === "compact" ? null : await attachInlineImages(item);
    const html = buildSignature(person, CONFIG.host, variant, images);
    await setSignature(item, html);
  } catch (e) {
    log("Signature not set: " + (e && e.message));
  } finally {
    event.completed();
  }
}

function onNewMessageComposeHandler(event) {
  applySignature(event);
}

function onMessageFromChangedHandler(event) {
  applySignature(event);
}

/** Ribbon button "Handtekening invoegen" – re-inserts the signature (and refreshes the cache). */
function insertSignatureButton(event) {
  try {
    if (Office.context.roamingSettings) Office.context.roamingSettings.remove(CACHE_KEY);
  } catch (e) {
    /* ignore */
  }
  applySignature(event);
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
Office.actions.associate("insertSignatureButton", insertSignatureButton);
