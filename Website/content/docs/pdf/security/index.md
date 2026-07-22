---
title: "PDF Security and Digital Signatures"
description: "Encrypt PDFs, inspect signature structure, verify CMS signatures, preserve signed revisions, and enrich long-term validation evidence."
layout: docs
---

PDF passwords, permissions, digital signatures, certification policies, and long-term validation evidence are different controls. Choose the control that matches the threat and compliance requirement.

## Encryption and permissions

`PdfStandardEncryptionOptions` configures user and owner passwords plus allowed operations. Encryption protects access to the file; permission flags depend on conforming readers and are not a substitute for authorization around the stored artifact.

## Signature validation

`PdfSignatureValidator` inspects byte ranges, signed revisions, field structure, subfilters, and security findings. Supply `PdfCmsSignatureCryptographyProvider` when mathematical CMS verification and certificate information are required.

Separate these questions in policy:

- Is the signature structure well formed?
- Does the signed-content digest match?
- Is the CMS signature mathematically valid?
- Is the signer certificate trusted for this purpose and time?
- Were later revisions added, and are they permitted by the certification level?

## Signing workflows

The external-signing API prepares the exact byte range and a bounded signature placeholder, delegates signing through `IPdfExternalSigner`, and completes the PDF with returned signature bytes. This supports certificate stores, hardware devices, remote signing services, and custom trust infrastructure without moving private keys into the document engine.

Visible signature appearance is optional and does not itself prove cryptographic validity. Conversely, a valid invisible signature can be cryptographically meaningful.

## Preserve existing signatures

Mutating a signed PDF can invalidate or supersede earlier signatures. Signature-aware operations produce mutation reports that compare coverage before and after the change. Follow the certification permission level and keep an immutable original when policy does not clearly allow the mutation.

Long-term validation enrichment can attach revocation and validation evidence only after the target signature has been cryptographically verified. Trust evaluation remains an application policy, not a blanket library claim.
