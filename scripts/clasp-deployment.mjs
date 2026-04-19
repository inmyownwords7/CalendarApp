import { spawn } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";

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

function runClasp(args) {
  return new Promise((resolvePromise, rejectPromise) => {
    const child = spawn("npx", ["clasp", ...args], {
      stdio: "inherit",
      shell: false
    });

    child.on("exit", (code) => {
      if (code === 0) {
        resolvePromise();
        return;
      }

      rejectPromise(new Error(`clasp ${args.join(" ")} exited with code ${code ?? 1}`));
    });

    child.on("error", rejectPromise);
  });
}

loadDotEnv();

const [, , action, maybeTarget, ...remainingArgs] = process.argv;
const hasExplicitTarget = maybeTarget === "current" || maybeTarget === "dev";
const target = hasExplicitTarget ? maybeTarget : "current";
const restArgs = hasExplicitTarget
  ? remainingArgs
  : maybeTarget
    ? [maybeTarget, ...remainingArgs]
    : remainingArgs;
const deploymentEnvKey =
  target === "dev" ? "CLASP_DEV_DEPLOYMENT_ID" : "CLASP_DEPLOYMENT_ID";
const deploymentId = process.env[deploymentEnvKey]?.trim();

if (!action) {
  console.error(
    "Usage: node scripts/clasp-deployment.mjs <open|update|delete> [current|dev] [args...]"
  );
  process.exit(1);
}

if (!deploymentId) {
  console.error(`Missing ${deploymentEnvKey} in .env`);
  process.exit(1);
}

const actionToArgs = {
  open: ["open-web-app", deploymentId, ...restArgs],
  update: ["redeploy", deploymentId, ...restArgs],
  delete: ["undeploy", deploymentId, ...restArgs]
};

const claspArgs = actionToArgs[action];

if (!claspArgs) {
  console.error(`Unsupported action "${action}". Expected one of: open, update, delete.`);
  process.exit(1);
}

await runClasp(claspArgs);
