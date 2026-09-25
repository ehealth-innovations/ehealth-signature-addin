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

/* global Office, fetch, console, __CONFIG__ */

import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { buildSignature } from "./template";

const CONFIG = __CONFIG__; // injected at build time from config.json
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
    const html = buildSignature(person, CONFIG.host, variant);

    await disableClientSignature(item);
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
