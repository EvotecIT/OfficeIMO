# OfficeIMO roadmap

This is the repository's single product backlog. It contains open work only. Implemented behavior is documented in package READMEs, support matrices, generated inventories, and current-state guides linked from [the documentation index](README.md).

An item belongs here when it has a clear product outcome and an owning package. Implementation checkpoints, completed task lists, release-wait notes, architectural rules, and competitor parity tables do not belong here.

Deliberately bounded compatibility contracts are not backlog by themselves. A preserved or rejected profile with a documented diagnostic becomes roadmap work only when the repository adopts a concrete supported shape and evidence plan for it.

## Release-wide quality

- [ ] Extend the generated Office compatibility catalog beyond the current Word, Excel, and PowerPoint legacy-format families into a package-neutral operation model for create, read, edit, preserve, inspect, convert, export, reject, and unsupported behavior.
- [ ] Generate compatible package README sections, website capability pages, MCP discovery, and support matrices from that model wherever one source can truthfully own the claim.
- [ ] Expand cross-producer fixture corpora with producer/version provenance and stable package or semantic diff policies.
- [ ] Add reproducible correctness, file-size, elapsed-time, peak-memory, allocation, cancellation, and deterministic-output evidence for representative workloads on every supported operating system.
- [ ] Add shared conversion reports and strict no-loss policies wherever an adapter can simplify, omit, rasterize, or preserve unsupported content.

## Security and protected content

- [ ] Add interoperable ODF encryption/decryption only after an external producer corpus, explicit password and key policy, and fail-safe preservation evidence are available.

## Completion rule

Remove an item when its public API, compatibility boundary, tests, generated evidence, and user documentation agree. GitHub Releases records delivered history, while `MIGRATION.md` retains only upgrade actions; this file does not retain completed milestones.
