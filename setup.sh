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
import { build } from "esbuild";

await build({
  entryPoints: ["src/gas/index.ts"],
  bundle: true,
  platform: "browser",
  target: "es2019",
  format: "iife",
  outfile: "dist/gas/Code.gs",
  sourcemap: false,
  legalComments: "none"
});

console.log("Built dist/gas/Code.gs");
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

if [ ! -f src/gas/index.ts ]; then
  cat > src/gas/index.ts <<'TS'
/**
 * GAS entrypoint.
 * Export global functions by assigning to globalThis if you are bundling as IIFE.
 */
(globalThis as any).hello = () => {
  Logger.log("Hello from GAS bundle!");
};
TS
fi

echo
echo "✅ Setup complete."
echo
echo "Next steps:"
echo "  1) GAS: npm run build:gas  (outputs dist/gas/Code.gs)"
echo "  2) Node dev: npm run dev:node"
echo "  3) Clasp: npm run clasp:login && npm run clasp:push"