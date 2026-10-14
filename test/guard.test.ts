import { expect, test } from "bun:test";
import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";
import * as extension from "../index";

const root = join(import.meta.dir, "..");

async function sources(): Promise<{ file: string; source: string }[]> {
  const files = (await readdir(root)).filter((name) => name.endsWith(".ts"));
  return Promise.all(
    files.map(async (file) => ({ file, source: await readFile(join(root, file), "utf8") })),
  );
}

test("the manifest has no launcher, and one Microsoft connection with the write scopes", async () => {
  const manifest = JSON.parse(await readFile(join(root, "prifly-extension.json"), "utf8"));
  expect(manifest.launch).toBeUndefined();
  expect(manifest.menu).toBeUndefined();
  expect(manifest.panels).toBeUndefined();
  expect(manifest.connections).toEqual([
    expect.objectContaining({
      provider: "microsoft",
      scopes: ["Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite"],
    }),
  ]);
  expect(manifest.description).not.toMatch(/never sends/i);
  expect(manifest.settings.map((field: { id: string }) => field.id)).toEqual(["account"]);
});

test("the module exports only activate: no launcher entry points", () => {
  expect(Object.keys(extension)).toEqual(["activate"]);
});

test("no source shows a chip, lights a flag or offers a board", async () => {
  const all = await sources();
  expect(all.map((each) => each.file)).not.toContain("board.ts");
  expect(all.map((each) => each.file)).not.toContain("launch.ts");
  for (const { file, source } of all) {
    expect(source, file).not.toMatch(/\bapi\??\.(show|flag|snoozeOptions)\b/);
    expect(source, file).not.toMatch(/\.(show|flag)\??\(/);
    expect(source, file).not.toMatch(/LaunchBoard|LaunchChoice|LauncherFlag|Decoration/);
  }
});

test("every file that calls a Graph write also asks the reader first", async () => {
  const writers = ["compose.ts", "changes.ts", "api-tool.ts", "tools.ts"];
  for (const { file, source } of await sources()) {
    const writes =
      /\b(sendMail|createDraft|sendDraft|addAttachment|moveMessage|patchMessage|updateDraft)\(|method(,|:\s*["'](POST|PATCH|PUT|DELETE))/.test(
        source,
      );
    if (file === "graph.ts" || !writes) continue;
    expect(writers, file).toContain(file);
    expect(source, file).toContain("await ask(api,");
  }
});
