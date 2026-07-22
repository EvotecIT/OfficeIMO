---
title: "Validation and Support Evidence"
description: "Understand what source, automated tests, generated API reference, package artifacts, and deployment smoke tests each prove."
layout: docs
---

OfficeIMO support is reported at the level that was actually validated. A project file proves that a component exists; it does not prove every workflow on every runtime. A passing unit test proves its scenario; it does not prove a published package or a NativeAOT binary.

## Evidence ladder

| Evidence | What it proves | What it does not prove |
|---|---|---|
| Project metadata | Ownership, dependencies, target frameworks, package intent | Runtime behavior or output fidelity |
| Generated API reference | Public types and members in a built assembly | That a particular workflow was exercised |
| Unit and integration tests | The asserted scenario on the tested framework and host | Every document variant or deployment mode |
| Generated artifact readback | The file can be reopened and key content survives | Pixel-perfect parity with every Office client |
| Packaged-artifact smoke test | The consumer path works outside project references | Public feed availability unless downloaded from that feed |
| Platform smoke test | The scenario executed on that OS and runtime identifier | Other architectures or unexercised code paths |
| NativeAOT publish and execution | The tested dependency graph compiled and ran natively | Blanket AOT support for unrelated packages and workflows |

## How the documentation uses evidence

Conceptual pages describe the intended workflow and link into generated reference for exact APIs. The [component catalog](/docs/capabilities/packages/) is generated from all production project files. PSWriteOffice counts come from its module manifest and fail validation if a cmdlet belongs to no documentation family or more than one total is reported.

The AOT page is maintained against checked-in executable smoke scenarios and records the runtime identifiers that were published and run. A repository script reproduces both passing paths and known compiler blockers. The matrix uses “not tested” instead of guessing from dependency size, annotations, or whether a project looks like a good candidate.

## Validate your own document corpus

Document formats are broad and real files combine features in ways synthetic samples do not. Before production:

- select representative source files, including damaged and edge-case inputs;
- assert semantic outcomes such as text, table data, comments, fields, formulas, pages, or signatures;
- capture structured diagnostics and decide which codes block delivery;
- reopen generated files with OfficeIMO and, where relevant, the target desktop or web client;
- preserve artifacts from failures so a regression can be reproduced.

The repository tests and examples are a starting point. Your accepted fidelity policy remains part of the application contract.
