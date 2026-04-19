/// <reference types="node" />

import { GAS_BUILD_OUTFILE, buildGas } from "./scripts/gas-build.mjs";

await buildGas();

console.log(`Built ${GAS_BUILD_OUTFILE}`);
