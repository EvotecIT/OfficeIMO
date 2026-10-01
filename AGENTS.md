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

- Keep host-dependent timing, allocation, retained/peak-memory budgets and large scale measurements in explicit opt-in evidence runs, outside ordinary pull-request correctness gates. Retain deterministic format, interoperability, cancellation, and API-enforced resource-limit checks in correctness validation, including those hosted in benchmark projects.
- Tag measurement-only xUnit cases with `Category=Performance` and exclude that category from ordinary CI routes. Report relocation separately from test deletion; keep optional measurement failures visible.
- Third-party libraries used for comparisons stay isolated in benchmark projects and opt-in verification runners; do not add them to OfficeIMO runtime projects without an explicit product decision.
- Keep opt-in benchmark runners outside the normal solution when their dependency or license profile should not affect normal restore and build.
- Add a comparison lane only when every implementation performs equivalent work and the output is validated. Keep OfficeIMO-specific workflows out of parity runners that cannot measure the same contract without artificial adapters.
- Keep benchmark execution and evidence-publication policy here or in benchmark tooling. Benchmark READMEs should explain how to run a suite, what it measures, how it validates output, and how to interpret its results.

## Project implementation discipline

- Follow the [open Microsoft Project milestones](Docs/ROADMAP.md#microsoft-project-document-library) as the single product scope. Use the [current model/XML contract and native feasibility evidence](OfficeIMO.Project/SUPPORT.md) as the baseline. Follow the remaining milestone order unless a concrete dependency or user direction justifies a recorded change.
- At each implementation task start, identify the active milestone, remaining acceptance criteria, existing evidence, and the next bounded deliverable. Inspect current source and prior task history before repeating discovery. Keep one milestone or an explicitly agreed delivery bundle in active focus apart from necessary shared-owner fixes.
- Preserve the adopted baseline: one typed model, thin fluent wrappers, no new external runtime dependencies, explicit scheduling, separate native-generation codecs, and operation-level preservation/conversion reports. Treat the executable public examples as the current API contract; do not propagate a later breaking redesign without identifying its effects.
- Reproduce and fix defects within the active milestone, including sibling paths and the shared owner, without asking for routine implementation permission. A defect is required work, not an excuse to abandon the milestone for a broader redesign.
- Before pursuing a newly discovered feature or changing an adopted boundary, record the trigger, affected milestone/API/dependencies, evidence, alternatives, and recommended change in the task. Ask the user only for material scope changes, new external dependencies, reduced acceptance criteria, or unresolved product choices. Continue independent authorized work while that decision is pending.
- Do not silently replace native writing with XML export, template-free creation with seed-file patching, legacy lifecycle support with reading only, or independent-producer proof with self-round-trip tests. An unavailable oracle, unresolved writer, or missing fixture is an explicit gap; elapsed time and green unrelated tests do not close it.
- Implement meaningful vertical slices and use focused validation before broad gates. Read-only local review, PR/CI follow-through, and merge/publication authorization retain their existing repository rules. Planning or implementation authorization does not automatically authorize a PR, merge, package release, or deployment.
- Before reporting a milestone complete, account for every acceptance criterion and validated in-scope defect. During active work use checklists in the task; once delivered, remove completed roadmap work after its current contract and evidence are in the proper package/matrix owners. Preserve stable milestone references in remaining dependencies, linking to those owners when a completed subsection is removed.
- End each implementation task with a compact handoff: milestone and satisfied/remaining criteria, branch/worktree and head, relevant evidence commands/results, unresolved decisions, PR/release state when applicable, exact next deliverable, and retained artifact paths. Do not create a second roadmap or a permanent implementation journal.

## Independent HTML engine development

- Keep the architecture in `Docs/officeimo.html-engine-design.md` and open milestones in `Docs/ROADMAP.md`; do not introduce a second program backlog. Update generated support contracts only when behavior is implemented and qualified.
- Prioritize usable components with replaceable dependency boundaries. Breaking API cleanup is accepted: prove the owned contract, migrate affected consumers and remove superseded paths instead of retaining compatibility wrappers by default. Keep effective parser, shaping, encoding and scripting providers until replacement is justified and qualified. The first implementation PR should be a usable foundation milestone; dependency independence is not its publication gate.
- Develop the engine on dedicated branches/worktrees under the configured Evotec repository root. Use a long-lived semantic integration branch such as `feature/html-engine` and short milestone branches when they improve isolation. Commit coherent changes and push checkpoints; keep exact owner/consumer commits in validation records.
- For this program, do not open a PR automatically at each checkpoint. Prepare a normal ready-for-review PR when a selected milestone is complete: implementation, consumer migration, focused tests, representative rendered evidence, relevant package/platform validation, current documentation and risk-appropriate local review. Static-engine, parser-independence and interactive-runtime milestones can become ready separately; readiness does not require finishing every long-term browser API.
- Treat pushes as branch backups, not release qualification. Inspect existing workflow triggers before relying on branch CI; use current build/test tooling for explicit branch validation where needed. Do not add release/publishing triggers as part of checkpoint setup.
- Keep upstream integration deliberate and regular. Inspect all worktrees before branch operations, use one coordinator for integration history, preserve unrelated work and avoid history rewrites on shared branches. Do not leave a giant unvalidated merge until the end of the program.
- Keep test references and optional provider comparisons outside default runtime dependencies. Store bulky corpus outputs in a named task-owned location, retain compact manifests and decisive evidence, and remove superseded output before repeated runs.

## Agent plugin and MCP Registry maintenance


`plugin.json` and `mcp.json` own metadata and server configuration. PowerForge generates the compatibility files used by older Codex clients and Claude. Use a source build of [PowerForge](https://github.com/EvotecIT/PSPublishModule) containing the `agent-plugin` command. Set `POWERFORGE_SOURCE` to that checkout, build with its pinned .NET SDK, and run this build from the PowerForge checkout:

```sh
dotnet build PowerForge.Cli/PowerForge.Cli.csproj -c Release -f net10.0
```

From the OfficeIMO checkout, invoke the built CLI:

```sh
dotnet "$POWERFORGE_SOURCE/PowerForge.Cli/bin/Release/net10.0/PowerForge.Cli.dll" agent-plugin sync --source .agents/plugins/officeimo-document-tools
dotnet "$POWERFORGE_SOURCE/PowerForge.Cli/bin/Release/net10.0/PowerForge.Cli.dll" agent-plugin validate --source .agents/plugins/officeimo-document-tools --project OfficeIMO.Tool/OfficeIMO.Tool.csproj
dotnet "$POWERFORGE_SOURCE/PowerForge.Cli/bin/Release/net10.0/PowerForge.Cli.dll" agent-plugin pack --source .agents/plugins/officeimo-document-tools --project OfficeIMO.Tool/OfficeIMO.Tool.csproj --out Artefacts/AgentPlugins
```

In PowerShell, use `$env:POWERFORGE_SOURCE` in place of `$POWERFORGE_SOURCE`. The resulting CLI runs on Windows, macOS, and Linux.

The packer produces a versioned ZIP and SHA-256 sidecar. Run the Agent Skills validator and real client/server checks in addition to package validation. OfficeIMO's release version bindings update the pinned tool version in all MCP configurations and the manual launcher in `.agents/plugins/officeimo-document-tools/README.md`; regenerate compatibility files after other metadata changes. Contributor skills live separately in `.agents/skills` and are not part of this user plugin.

The **Agent Plugin Package** workflow validates the package and MCP Registry metadata on plugin and Registry package-input changes and manual runs. It builds a pinned PowerForge source revision with its own SDK and uploads the ZIP and checksum as a workflow artifact. Normal `OfficeIMO-vYYYYMMDDHHMMSS` release events attach those files to the [GitHub release](https://github.com/EvotecIT/OfficeIMO/releases). Existing assets are preserved; attaching a duplicate filename fails. The plugin version follows the OfficeIMO.Tool NuGet release version. The canonical plugin version binding regenerates client manifests in the same PowerForge release transaction; generated files have no separate release bindings. Every published bundle change, including skills-only changes, requires a new product patch release. Keep schema/protocol and shared PowerForge versions independent.

PowerForge's project release bindings update both version fields in `server.json`. Publish the signed NuGet package first and check that its embedded README contains the matching `mcp-name` ownership marker. The registry validates the published package, so source metadata alone is insufficient. Then, from the repository root, authenticate an authorized EvotecIT publisher and submit the metadata:

```text
mcp-publisher login github
mcp-publisher publish OfficeIMO.Tool/server.json
```

For CI publication, the official publisher supports `mcp-publisher login github-oidc` with `id-token: write` on an authorized GitHub workflow. Follow the [MCP Registry publishing instructions](https://modelcontextprotocol.io/registry/quickstart) and verify the returned name and version in the Registry API. Registry publication provides discovery; ChatGPT and Claude public directories have their own submission and approval requirements.

`Build/validate_mcp_registry.py` validates against a fixed official schema URI and checks the source ownership marker. The OfficeIMO.Tool package smoke gate checks the same marker in the README declared by the actual NuGet artifact. Keep both checks when changing Registry metadata or package inputs.
