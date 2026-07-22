---
title: "Reliable Automation Patterns"
description: "Structure document jobs for repeatability, diagnostics, safe mutation, and CI artifacts."
layout: docs
---

Document automation becomes easier to operate when scripts separate source data, document composition, validation, and delivery.

## Keep paths and data explicit

Resolve input and output roots once. Pass plain objects into document builders instead of reading global state from nested DSL blocks. Write validation output beside the artifact or into a CI report folder.

## Prefer canonical commands in reusable scripts

PSWriteOffice publishes aliases for a compact DSL, but canonical commands make long-lived scripts easier to search, document, and compare with generated help. Use aliases when they materially improve a local composition block; keep entry points and operational steps canonical.

## Inspect before mutation

For existing files, collect preflight, capability, signature, compliance, accessibility, or format-detection evidence before changing the document. Write to a new path during development and reopen the result through the same engine.

## Treat conversion as a result

Keep the source, destination, diagnostics, warnings, and any fidelity policy together. A successful file write does not by itself prove that complex layout or unsupported features were preserved.

## Make jobs idempotent

Create a fresh output folder or use stable file names with deliberate overwrite behavior. Avoid appending the same section, sheet, or slide on every rerun unless the script first detects whether it already exists.

## Bound batch work

For Reader and conversion batches, limit concurrency based on document size and native-memory pressure. Capture failures per file so one malformed input does not erase evidence for the rest of the batch.

## Use the generated surface

The [command reference](/api/powershell/) is the authoritative parameter contract. The [examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples) demonstrate complete job shapes. Pin the module version in production automation and review release notes before changing that pin.
