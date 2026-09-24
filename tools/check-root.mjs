// Asserts that every file in deploy/ is byte-identical to the same path at the repository root,
// which is what GitHub Pages serves. Run after `npm run build` so deploy/ is fresh.
import { createHash } from "node:crypto";
import { readdir, readFile } from "node:fs/promises";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const deploy = resolve(root, "deploy");

async function walk(dir) {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) out.push(...(await walk(full)));
    else out.push(full);
  }
  return out;
}

const sha = (buf) => createHash("sha256").update(buf).digest("hex");
const files = (await walk(deploy)).map((f) => relative(deploy, f)).sort();
if (files.length === 0) throw new Error("deploy/ is empty; run npm run build first");

const failures = [];
for (const file of files) {
  const built = await readFile(join(deploy, file));
  let served;
  try {
    served = await readFile(join(root, file));
  } catch {
    failures.push(`${file}: missing at repository root`);
    continue;
  }
  if (!built.equals(served)) failures.push(`${file}: root ${sha(served).slice(0, 12)} != deploy ${sha(built).slice(0, 12)}`);
}

if (failures.length) {
  console.error(`served root differs from deploy/ (${failures.length} of ${files.length} files):`);
  for (const line of failures) console.error(`  ${line}`);
  process.exit(1);
}
console.log(`served root matches deploy/: ${files.length} files byte-identical`);
