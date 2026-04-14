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
