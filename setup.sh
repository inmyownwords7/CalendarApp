#!/usr/bin/env bash
set -euo pipefail

# -----------------------------
# Init project
# -----------------------------
if [ ! -f package.json ]; then
  npm init -y
fi

# -----------------------------
# Install deps
# -----------------------------
npm i -D \
  typescript \
  esbuild \
  tsx \
  @types/node \
  @types/google-apps-script \
  @google/clasp

npm i \
  @slack/web-api \
  @notionhq/client

# -----------------------------
# Create folders
# -----------------------------
mkdir -p src/gas src/node src/shared dist/gas dist/node scripts

# -----------------------------
# Write tsconfig.node.json
# -----------------------------
cat > tsconfig.node.json <<'JSON'
{
  "compilerOptions": {
    "target": "ES2022",
    "lib": ["ES2022"],
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "types": ["node"],
    "strict": true,
    "skipLibCheck": true,

    "rootDir": "src",
    "outDir": "dist/node",

    "esModuleInterop": true,
    "resolveJsonModule": true
  },
  "include": ["src/node/**/*.ts", "src/shared/**/*.ts"],
  "exclude": ["node_modules", "dist", "src/gas/**"]
}
JSON

# -----------------------------
# Write tsconfig.gas.json
# -----------------------------
cat > tsconfig.gas.json <<'JSON'
{
  "compilerOptions": {
    "target": "ES2019",
    "lib": ["ES2019"],
    "module": "None",
    "types": ["google-apps-script"],
    "strict": true,
    "skipLibCheck": true,

    "rootDir": "src",
    "outDir": "dist/gas",

    "esModuleInterop": true,
    "resolveJsonModule": true
  },
  "include": ["src/gas/**/*.ts", "src/shared/**/*.ts"],
  "exclude": ["node_modules", "dist", "src/node/**"]
}
JSON

# -----------------------------
# GAS single-file bundler (esbuild)
# -----------------------------
cat > scripts/build-gas.mjs <<'MJS'
import { GAS_BUILD_OUTFILE, buildGas } from "./gas-build.mjs";

await buildGas();

console.log(`Built ${GAS_BUILD_OUTFILE}`);
MJS

cat > scripts/gas-build.mjs <<'MJS'
import { build } from "esbuild";

export const GAS_BUILD_OUTFILE = "dist/Code.gs";

export const GAS_EXPOSED_FUNCTIONS = Object.freeze([
  "hello"
]);

function buildGasFooter() {
  return GAS_EXPOSED_FUNCTIONS.map(
    (name) =>
      `function ${name}() { return globalThis._${name}.apply(globalThis, arguments); }`
  ).join("\n");
}

export async function buildGas() {
  await build({
    entryPoints: ["src/gas/entrypoints/index.ts"],
    bundle: true,
    platform: "browser",
    target: "es2019",
    format: "iife",
    outfile: GAS_BUILD_OUTFILE,
    sourcemap: false,
    legalComments: "none",
    footer: {
      js: buildGasFooter()
    }
  });
}
MJS

# -----------------------------
# Node bundler (optional) - outputs dist/node/index.js
# If you prefer pure tsx for dev only, you can ignore this.
# -----------------------------
cat > scripts/build-node.mjs <<'MJS'
import { build } from "esbuild";

await build({
  entryPoints: ["src/node/index.ts"],
  bundle: true,
  platform: "node",
  target: "node20",
  format: "esm",
  outfile: "dist/node/index.js",
  sourcemap: true,
  legalComments: "none"
});

console.log("Built dist/node/index.js");
MJS

# -----------------------------
# appsscript.json scaffold (optional)
# -----------------------------
if [ ! -f appsscript.json ]; then
  cat > appsscript.json <<'JSON'
{
  "timeZone": "America/Montreal",
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
JSON
fi

# -----------------------------
# Update package.json (type=module + scripts)
# -----------------------------
node - <<'NODE'
const fs = require("fs");
const pkg = JSON.parse(fs.readFileSync("package.json", "utf8"));

pkg.type = "module";

pkg.scripts = {
  ...(pkg.scripts || {}),
  "typecheck:node": "tsc -p tsconfig.node.json --noEmit",
  "typecheck:gas": "tsc -p tsconfig.gas.json --noEmit",

  "build:gas": "node scripts/build-gas.mjs",
  "build:node": "node scripts/build-node.mjs",

  "dev:node": "tsx watch src/node/index.ts",
  "start:node": "node dist/node/index.js",

  "clasp:login": "clasp login",
  "clasp:push": "clasp push",
  "clasp:pull": "clasp pull",
  "clasp:status": "clasp status"
};

fs.writeFileSync("package.json", JSON.stringify(pkg, null, 2));
console.log("Updated package.json (type=module + scripts).");
NODE

# -----------------------------
# Minimal entrypoints (optional helpers)
# -----------------------------
if [ ! -f src/node/index.ts ]; then
  cat > src/node/index.ts <<'TS'
import { WebClient } from "@slack/web-api";
import { Client as NotionClient } from "@notionhq/client";

console.log("Node dev entrypoint running.");

const slack = new WebClient(process.env.SLACK_BOT_TOKEN);
const notion = new NotionClient({ auth: process.env.NOTION_TOKEN });

// Put your Node-side experimentation here.
TS
fi

mkdir -p src/gas/entrypoints

if [ ! -f src/gas/entrypoints/index.ts ]; then
  cat > src/gas/entrypoints/index.ts <<'TS'
/**
 * GAS bundle entrypoint.
 * Assign handlers to globalThis._name and let the esbuild footer emit
 * the real top-level Apps Script functions.
 */
(
  globalThis as Record<string, unknown>
)._hello = () => {
  Logger.log("Hello from GAS bundle!");
};
TS
fi

echo
echo "✅ Setup complete."
echo
echo "Next steps:"
echo "  1) GAS: npm run build:gas  (outputs dist/Code.gs)"
echo "  2) Node dev: npm run dev:node"
echo "  3) Clasp: npm run clasp:login && npm run clasp:push"
