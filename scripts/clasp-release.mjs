import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { spawn } from "node:child_process";

function loadDotEnv(filePath = ".env") {
  const absolutePath = resolve(process.cwd(), filePath);
  if (!existsSync(absolutePath)) {
    return;
  }

  const content = readFileSync(absolutePath, "utf8");
  for (const rawLine of content.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      continue;
    }

    const separatorIndex = line.indexOf("=");
    if (separatorIndex === -1) {
      continue;
    }

    const key = line.slice(0, separatorIndex).trim();
    if (!key || process.env[key] !== undefined) {
      continue;
    }

    let value = line.slice(separatorIndex + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }

    process.env[key] = value;
  }
}

function run(command, args, options = {}) {
  const { captureOutput = false } = options;

  return new Promise((resolvePromise, rejectPromise) => {
    const child = spawn(command, args, {
      stdio: captureOutput ? ["inherit", "pipe", "pipe"] : "inherit",
      shell: false
    });

    let output = "";
    if (captureOutput && child.stdout) {
      child.stdout.on("data", (chunk) => {
        const text = chunk.toString();
        output += text;
        process.stdout.write(text);
      });
    }
    if (captureOutput && child.stderr) {
      child.stderr.on("data", (chunk) => {
        const text = chunk.toString();
        output += text;
        process.stderr.write(text);
      });
    }

    child.on("exit", (code) => {
      if (code === 0) {
        resolvePromise(output);
        return;
      }

      rejectPromise(
        new Error(`${command} ${args.join(" ")} exited with code ${code ?? 1}`)
      );
    });

    child.on("error", rejectPromise);
  });
}

loadDotEnv();

const [, , target = "current", ...descriptionParts] = process.argv;
if (target !== "current" && target !== "dev") {
  console.error('Usage: node scripts/clasp-release.mjs [current|dev] [description...]');
  process.exit(1);
}

const deploymentEnvKey =
  target === "dev" ? "CLASP_DEV_DEPLOYMENT_ID" : "CLASP_DEPLOYMENT_ID";
const deploymentId = process.env[deploymentEnvKey]?.trim();

if (!deploymentId) {
  console.error(`Missing ${deploymentEnvKey} in .env`);
  process.exit(1);
}

const defaultDescription = `release ${new Date().toISOString()}`;
const description = descriptionParts.join(" ").trim() || defaultDescription;

await run("npm", ["run", "build:gas"]);
await run("npx", ["clasp", "push"]);

const versionOutput = await run(
  "npx",
  ["clasp", "version", description],
  { captureOutput: true }
);

const versionMatch = versionOutput.match(/Created version (\d+)/i);
if (!versionMatch) {
  throw new Error("Failed to parse version number from clasp version output.");
}

const versionNumber = versionMatch[1];
const redeployOutput = await run("npx", [
  "clasp",
  "redeploy",
  deploymentId,
  "-V",
  versionNumber,
  "-d",
  description
], { captureOutput: true });

if (/read-only deployments may not be modified/i.test(redeployOutput)) {
  throw new Error(`Deployment ${deploymentId} is read-only and cannot be updated.`);
}

console.log(
  `Updated ${target} deployment ${deploymentId} to version ${versionNumber}.`
);
