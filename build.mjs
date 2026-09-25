// Build: bundles the event handler (incl. MSAL) into one DOM-free file and
// generates manifest.xml, the .well-known allow-list and a local preview page.
import { build } from "esbuild";
import { readFileSync, writeFileSync, mkdirSync, cpSync, rmSync } from "node:fs";

const cfg = JSON.parse(readFileSync("config.json", "utf8"));
const host = cfg.host.replace(/\/+$/, "");
const placeholder = /REPLACE-ME|^0{8}-/;
if (placeholder.test(host) || placeholder.test(cfg.clientId) || placeholder.test(cfg.tenantId)) {
  console.warn("⚠  config.json still has placeholder values (host / clientId / tenantId).");
}

rmSync("dist", { recursive: true, force: true });
mkdirSync("dist/.well-known", { recursive: true });

const runtimeConfig = {
  host,
  clientId: cfg.clientId,
  tenantId: cfg.tenantId,
  cacheHours: cfg.cacheHours,
  replyVariant: cfg.replyVariant,
};

await build({
  entryPoints: ["src/launchevent.js"],
  bundle: true,
  minify: true,
  format: "iife",
  // Classic Outlook's JS-only runtime is conservative: lower syntax (incl. async/await).
  target: ["es2016"],
  define: { __CONFIG__: JSON.stringify(runtimeConfig) },
  outfile: "dist/launchevent.js",
  legalComments: "none",
});

cpSync("src/commands.html", "dist/commands.html");
cpSync("assets", "dist/assets", { recursive: true });
cpSync("staticwebapp.config.json", "dist/staticwebapp.config.json");

writeFileSync(
  "dist/manifest.xml",
  readFileSync("manifest.template.xml", "utf8")
    .replace(/{{HOST}}/g, host)
    .replace(/{{ADDIN_ID}}/g, cfg.addinId)
    .replace(/{{VERSION}}/g, cfg.version)
);

// Required by classic Outlook for Windows so the JS runtime may use SSO/NAA.
writeFileSync(
  "dist/.well-known/microsoft-officeaddins-allowed.json",
  JSON.stringify({ allowed: [host + "/launchevent.js"] }, null, 2)
);

// Local preview of the template with sample data (open dist/preview.html in a browser).
const { buildSignature } = await import("./src/template.js");
const samples = [
  ["Volledig (alle velden)", { name: "Jelle van der Coelen", title: "Directeur", phone: "+31 6 12 34 56 78", email: "jelle@ehealthinnovations.nl" }, "full"],
  ["Zonder functie en telefoon", { name: "Jelle van der Coelen", title: "", phone: "", email: "jelle@ehealthinnovations.nl" }, "full"],
  ["Compact (antwoord)", { name: "Jelle van der Coelen", title: "Directeur", phone: "06-12345678", email: "jelle@ehealthinnovations.nl" }, "compact"],
];
const previewBase = "./"; // assets relative to dist/
writeFileSync(
  "dist/preview.html",
  `<!DOCTYPE html><html lang="nl"><meta charset="utf-8"><title>Handtekening preview</title>
<body style="font-family:Arial,sans-serif;background:#f4f4f4;padding:24px;">` +
    samples
      .map(
        ([label, p, v]) =>
          `<h3 style="font-size:13px;color:#555;">${label}</h3><div style="background:#fff;padding:24px;margin-bottom:24px;">` +
          buildSignature(p, previewBase.replace(/\/$/, "") || ".", v) +
          `</div>`
      )
      .join("") +
    `</body></html>`
);

console.log("✓ dist/ built for " + host);
