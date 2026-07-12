# OfficeIMO.Email MSG parity review - 2026-07-11

This review checked whether `OfficeIMO.Email` can replace the MsgKit, MsgReader, and OpenMcdf dependency chain used by Mailozaurr without turning OfficeIMO into a transport or cryptography library. The current product contract lives in the [support matrix](../officeimo.email-support-matrix.md); this file records the comparison and the defects found during implementation.

## Decision

Mailozaurr can move MSG and Outlook-item artifact handling to `OfficeIMO.Email`, keep MimeKit and MailKit for EML, transport, and security, and remove the separate `Mailozaurr.Msg` implementation after an `OfficeIMO.Email` package is published.

The replacement is capability-based rather than API-for-API cloning:

- typed models cover the fields applications normally use;
- every decoded MAPI value remains accessible for custom Outlook properties;
- OfficeIMO owns the compound storage needed by MSG, but does not advertise a general CFB transaction library;
- protected MSG content is extracted for MimeKit instead of validating certificates inside OfficeIMO.

## Defects found and closed

- String8 decoding now resolves Outlook code-page properties and LCID fallbacks, including DBCS input, and retains untouched source bytes for safe rewrites.
- Sender, representing sender, received-by addresses, Exchange address types, SMTP fallbacks, original addresses, room, resource, and Reply-To roles now have explicit mappings.
- Recipient One-Off Entry IDs and search keys are deterministic. The writer no longer invents top-level or attachment store identity blobs where no valid provider identity exists.
- The compound writer did not persist the MSG root storage CLSID. It now writes `00020D0B-0000-0000-C000-000000000046` for MSG without changing other Office compound formats.
- The MS-OXMSG named-property entry stream had its GUID/kind and property-index fields reversed. Self-round-trips hid the problem. The reader and writer now use the specification layout, and the external corpus no longer produces `EMAIL_MSG_NAMEID_GUID_INVALID` warnings.
- Named-property output omitted the required property-name-to-ID hash streams. It now emits numerical and string lookup streams, groups hash collisions, uses the Outlook-compatible CRC-32 calculation, and normalizes Internet-header names for their hash calculation.
- Newly-authored items now carry the Outlook-compatible `PidNameAcceptLanguage` mapping, a complete NameID storage, the Unicode store-support mask, and a deterministic creation-time fallback when the model has no date.
- Trailing zero alignment bytes in a real property stream are accepted; nonzero partial entries remain diagnostic.
- Contact Email1 used the original-display-name property instead of the email-address property. It now uses `0x8083` and retains the complete three-slot contact address model.
- Task Complete and Owner used the wrong named-property IDs. They now use `0x811C` and `0x811F` and interoperate with MsgReader.
- Appointments, contacts, tasks, journals, and sticky notes now expose cohesive typed read/write models while preserving recurrence, time-zone, and custom binary payloads.
- Attachment handling now includes inline/hidden/contact-photo flags, rendering position, linked paths, dates, embedded MSG, OLE/custom storage, and nested TNEF.
- Opaque and clear-signed S/MIME MSG classes expose their payload for host-side MimeKit processing without claiming verification.
- Common metadata now covers subject/conversation state, importance/priority, draft/read/receipt state, categories, reaction payloads, modification metadata, sensitivity, locale, and editor format.

## Oracle evidence

The test project uses MsgKit, MsgReader, OpenMcdf, and MimeKit only as test oracles. Product assemblies do not reference them.

- MsgKit 3.0.5 generates EML-to-MSG and named-contact fixtures consumed by OfficeIMO.
- MsgReader 6.0.12 consumes OfficeIMO-authored messages and typed Outlook items.
- Fifteen MsgReader repository MSG fixtures cover text, HTML, RTF, special-character subjects, Exchange sender fallbacks, reactions, attachments, and embedded messages. All matched stable subject and collection contracts without MSG parse errors or structural warnings.
- OpenMcdf 3.1.4 opens OfficeIMO compound output, including a large file that requires DIFAT sectors.
- MimeKit's TNEF reader accepts OfficeIMO output as compliant.
- Microsoft Outlook for Mac opens OfficeIMO-authored message, appointment, contact, task, journal, and note files under their expected subjects. The message view displayed sender, recipient, body, and attachment. Mac Outlook presents non-mail MSG classes through its generic item viewer, so Windows-specific appointment/contact/task/journal/note editors remain a separate UI validation step.

Body comparisons normalize line endings because MsgReader rewrites CR/LF sequences in some projections. OfficeIMO retains the decoded source form rather than reproducing that presentation normalization.

## Boundaries that remain intentional

- Signature verification and decryption require a configured MimeKit cryptography context and are part of Mailozaurr, not OfficeIMO.
- PST/OST stores require a mailbox-store product, not the MSG artifact parser.
- Exchange directory lookup is not available offline; native EX addresses and SMTP fallbacks are retained so the host can resolve them.
- Arbitrary CFB transaction APIs remain outside `OfficeIMO.Email`. The internal compound engine is complete for MSG and structured attachments.
- Unknown or vendor-specific named properties remain typed MAPI entries until a stable cross-consumer convenience model is justified.

This boundary is sufficient for the existing Mailozaurr MSG/EML conversion and import commands and leaves room for richer Outlook-item output without recreating the removed dependency forest.
