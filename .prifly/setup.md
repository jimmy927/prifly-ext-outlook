---
setup:
  - bun install --frozen-lockfile
---

A plain Bun TypeScript extension: no workspaces, no build step. The install brings biome,
TypeScript, `@types/bun` and the extension API's types; then `bun test`, `bun run typecheck` and
`bun run check` work (see `package.json`).

What it relies on outside the worktree:
- `bun` on PATH.
- `@prifly/extension-api` is a git dependency
  (`git+ssh://git@github.com/jimmy927/prifly-extension-api.git#v0.1.0`), so the install needs ssh
  access to GitHub with a key that can read that private repository.
- `bun.lock` committed beside `package.json`. The template's first commit has neither the lock nor
  the git dependency (it used `link:@prifly/extension-api`), so a worktree of that commit fails this
  step until the first writer's change, which adds both, has landed.
