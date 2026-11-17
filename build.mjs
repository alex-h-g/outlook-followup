// Bundles src/ into dist/ (what GitHub Pages serves). Usage: npm run build
import { build } from "esbuild";
import { cpSync, mkdirSync, rmSync } from "node:fs";

rmSync("dist", { recursive: true, force: true });
mkdirSync("dist/assets", { recursive: true });

await build({
  entryPoints: { taskpane: "src/taskpane.js", compose: "src/compose.js" },
  bundle: true,
  format: "iife",
  target: ["safari15"], // Outlook for Mac uses WKWebView
  minify: true,
  sourcemap: false,
  outdir: "dist",
  legalComments: "none",
});

for (const f of ["taskpane.html", "compose.html", "commands.html", "styles.css"]) cpSync(`src/${f}`, `dist/${f}`);
cpSync("assets", "dist/assets", { recursive: true });
cpSync("manifest.xml", "dist/manifest.xml");
console.log("Built dist/");
