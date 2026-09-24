import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { createRequire } from "node:module";
import { spawnSync } from "node:child_process";

const require = createRequire(import.meta.url);
const candidates = [process.env.CODEX_PRIMARY_NODE_MODULES, process.env.NODE_PATH, join(homedir(), ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules")].filter(Boolean);
let modulePath = "";
try { modulePath = require.resolve("playwright"); } catch {}
for (const candidate of candidates) {
  if (modulePath || !existsSync(candidate)) continue;
  try { modulePath = require.resolve("playwright", { paths: [candidate] }); } catch {}
}
if (!modulePath) throw new Error("Playwright is required for runtime privacy verification. Install it locally or set CODEX_PRIMARY_NODE_MODULES; this script never downloads dependencies.");
const run = spawnSync(process.execPath, ["--test", "--test-reporter=tap", "tests/browser.browser.mjs"], { encoding: "utf8", timeout: 10000, killSignal: "SIGKILL", env: { ...process.env, BHB_PLAYWRIGHT_MODULE: modulePath } });
const output = `${run.stdout || ""}${run.stderr || ""}`;
if ((run.status || run.error) && /MachPortRendezvousServer[\s\S]*Permission denied/u.test(output)) {
  console.log("SKIP browser runtime: the managed macOS sandbox denied Chromium's Mach rendezvous service. Run `npm run test:browser` outside the sandbox; no weaker fallback is treated as browser evidence.");
  process.exit(0);
}
process.stdout.write(output);
process.exit(run.status ?? 1);
