# OfficeIMO repository instructions

## Documentation audiences

- The root README and package READMEs are user-facing. They explain what the current package does, how to install it, and how to use its public API.
- `MIGRATION.md` is the user-facing upgrade contract. Keep actionable old-to-new package, API, configuration, and behavior changes there; GitHub Releases owns release history.
- `Docs/README.md` is a navigation page for readers and contributors. Keep it focused on finding the right guide, contract, evidence, or roadmap entry.
- `Docs/ROADMAP.md` is the single product backlog. It contains open work, not completed milestones or implementation journals.
- `AGENTS.md` owns repository-maintenance instructions for coding agents. Do not put agent workflow, cleanup policy, or documentation-governance rules into user-facing READMEs.
- Website comparison pages may help users choose between libraries. Repository product documentation should explain OfficeIMO through its contracts, standards, fixtures, and measured workloads rather than competitor parity narratives.

## Documentation maintenance

- Document the current source and package contract. Do not add release-wait, preview-state, pull-request progress, or "after publication" language to current product docs.
- Keep installation and public API examples in the README of the package that owns the API.
- Keep release summaries out of the repository. Link GitHub Releases for history and preserve only required upgrade actions in `MIGRATION.md`.
- Keep exact coverage and known limitations in the relevant support or capability matrix.
- Move actionable product gaps to `Docs/ROADMAP.md`; do not create another backlog, readiness review, gap plan, or package roadmap.
- Keep dated reports only when the date is part of reproducible evidence, such as a benchmark run.
- Update generated documents through their source catalogs, manifests, tests, or generators.
- Before removing a stale review or planning document, move current behavior to its public owner and open work to `Docs/ROADMAP.md`.
- Keep user-facing install commands on the latest version actually published to NuGet. Validate install examples through a representative restore/build consumer or release tool rather than unit-test substring assertions over documentation prose.

## Testing evidence

- Do not use product unit tests to pin human-authored README, roadmap, migration, compatibility, planning, or agent prose.
- Test Markdown when it is a supported input/output format. Compare generated documents with their executable generator or machine-readable source of truth.
- Exercise documentation examples through real compile, restore, or run consumers where practical.
- When low-value prose-only, implementation-only, duplicate, or obsolete tests are encountered in a touched area, remove them instead of preserving or replacing them for test-count optics.

## Benchmark boundaries

- Third-party libraries used for comparisons stay isolated in benchmark projects and opt-in verification runners; do not add them to OfficeIMO runtime projects without an explicit product decision.
- Keep opt-in benchmark runners outside the normal solution when their dependency or license profile should not affect normal restore and build.
- Add a comparison lane only when every implementation performs equivalent work and the output is validated. Keep OfficeIMO-specific workflows out of parity runners that cannot measure the same contract without artificial adapters.
- Keep benchmark execution and evidence-publication policy here or in benchmark tooling. Benchmark READMEs should explain how to run a suite, what it measures, how it validates output, and how to interpret its results.

## Independent HTML engine development

- Keep the architecture in `Docs/officeimo.html-engine-design.md` and open milestones in `Docs/ROADMAP.md`; do not introduce a second program backlog. Update generated support contracts only when behavior is implemented and qualified.
- Prioritize usable components with replaceable dependency boundaries. Breaking API cleanup is accepted: prove the owned contract, migrate affected consumers and remove superseded paths instead of retaining compatibility wrappers by default. Keep effective parser, shaping, encoding and scripting providers until replacement is justified and qualified. The first implementation PR should be a usable foundation milestone; dependency independence is not its publication gate.
- Develop the engine on dedicated branches/worktrees under the configured Evotec repository root. Use a long-lived semantic integration branch such as `feature/html-engine` and short milestone branches when they improve isolation. Commit coherent changes and push checkpoints; keep exact owner/consumer commits in validation records.
- For this program, do not open a PR automatically at each checkpoint. Prepare a normal ready-for-review PR when a selected milestone is complete: implementation, consumer migration, focused tests, representative rendered evidence, relevant package/platform validation, current documentation and risk-appropriate local review. Static-engine, parser-independence and interactive-runtime milestones can become ready separately; readiness does not require finishing every long-term browser API.
- Treat pushes as branch backups, not release qualification. Inspect existing workflow triggers before relying on branch CI; use current build/test tooling for explicit branch validation where needed. Do not add release/publishing triggers as part of checkpoint setup.
- Keep upstream integration deliberate and regular. Inspect all worktrees before branch operations, use one coordinator for integration history, preserve unrelated work and avoid history rewrites on shared branches. Do not leave a giant unvalidated merge until the end of the program.
- Keep test references and optional provider comparisons outside default runtime dependencies. Store bulky corpus outputs in a named task-owned location, retain compact manifests and decisive evidence, and remove superseded output before repeated runs.
