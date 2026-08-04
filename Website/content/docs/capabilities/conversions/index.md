---
title: "Conversion and Rendering Map"
description: "Choose the owning model, focused adapter, diagnostic policy, and deployment path for OfficeIMO conversions. Includes examples and API links."
layout: docs
---

OfficeIMO keeps document models and conversion adapters separate. That lets an application reference the source engine it needs and add only the destination routes it actually ships.

## Choose by source and destination

This table is rendered from the same public `OfficeConversionCapabilityCatalog` used by MCP discovery and the browser converter. “Managed only” means the package route exists but is intentionally absent from the WebAssembly app.

{{< conversion-routes >}}

## Preserve the source model

Use the source package for loading, editing, and source-specific validation. The adapter should own projection into the destination. This matters because a table, field, comment, animation, formula, or drawing can have no exact equivalent in the target format.

Do not treat “a file was written” as proof that the conversion preserved everything important. For production routes:

1. load or build the source with its native model;
2. set an explicit resource policy for fonts, images, links, or external content;
3. run the focused converter and capture its structured warnings;
4. inspect or reopen the destination artifact;
5. keep representative fixtures in automated tests.

## Browser-local routes

The [browser converter](/convert/) exposes only routes that can execute safely inside the WebAssembly application. That is intentionally smaller than the managed server-side conversion surface. A missing browser route does not mean that no .NET adapter exists; consult the [complete component index](/docs/capabilities/packages/) and the adapter API.

## Loss policy

Converters should be selected by the content that must survive, not only by file extensions. Decide whether unsupported content should fail the operation, produce a warning, be approximated, or be omitted. Keep the original when legal or audit requirements make the conversion evidence important.

## Next steps

- Use the [Word conversion guide](/docs/word/conversion/) for DOCX-specific resource and review concerns.
- Use the [PDF conversion guide](/docs/pdf/conversion/) for fixed-layout delivery and diagnostic review.
- Use [Reader](/docs/reader/) when the goal is normalized extraction rather than a destination document.
