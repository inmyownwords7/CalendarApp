import { GAS_BUILD_OUTFILE, buildGas } from "./gas-build.mjs";

await buildGas();

console.log(`Built ${GAS_BUILD_OUTFILE}`);
