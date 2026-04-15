/// <reference types="node" />

import { build } from "esbuild";

await build({
  entryPoints: ["src/gas/index.ts"],
  bundle: true,
  platform: "browser",
  target: "es2019",
  format: "iife",
  outfile: "dist/gas/Code.gs",
  sourcemap: false,
  legalComments: "none",
});

console.log("Built dist/gas/Code.gs");
