---
title: "Conversion and Rendering Map"
description: "Find the OfficeIMO package, public API, fidelity model, runtime, and diagnostic result type for each supported document conversion route."
layout: docs
---

Choose a source and destination below, then add the focused package to your .NET application. Each route identifies what the conversion preserves, where it can run, and which result type carries warnings or loss information.

## Choose by source and destination

Routes marked **Browser or .NET** are available in the [browser converter](/convert/) and through the listed package. Routes marked **.NET application** run through the package in an application or service you control.

{{< conversion-routes >}}

## Preserve the source model

Load and edit the document with its source package, then call the focused adapter shown in the table. This matters because a table, field, comment, animation, formula, or drawing may have no exact equivalent in the destination format.

Do not treat “a file was written” as proof that the conversion preserved everything important. For production routes:

1. load or build the source with its native model;
2. set an explicit resource policy for fonts, images, links, or external content;
3. run the focused converter and capture its structured warnings;
4. inspect or reopen the destination artifact;
5. keep representative fixtures in automated tests.

## Browser-local routes

Browser routes run locally in the current tab; OfficeIMO does not upload the source file to a conversion service. Routes for a **.NET application** run wherever you host them, so your application controls authentication, storage, logging, and retention. If a route is absent from the browser converter, check the [complete component index](/docs/capabilities/packages/) for a focused .NET adapter.

## Loss policy

Converters should be selected by the content that must survive, not only by file extensions. Decide whether unsupported content should fail the operation, produce a warning, be approximated, or be omitted. Keep the original when legal or audit requirements make the conversion evidence important.

## Next steps

- Use the [Word conversion guide](/docs/word/conversion/) for DOCX-specific resource and review concerns.
- Use the [PDF conversion guide](/docs/pdf/conversion/) for fixed-layout delivery and diagnostic review.
- Use [Reader](/docs/reader/) when the goal is normalized extraction rather than a destination document.
