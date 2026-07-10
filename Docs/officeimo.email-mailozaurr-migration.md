# Moving Mailozaurr MSG support to OfficeIMO.Email

Mailozaurr can keep MailKit and MimeKit for transport, server access, and message security while moving persisted-file handling to `OfficeIMO.Email`. This removes the reason for the separate `Mailozaurr.Msg` implementation and its MsgKit, MsgReader, OpenMcdf, RtfPipe, and Microsoft.Maui.Graphics dependency chain.

## Ownership after the change

| Capability | Owner |
| --- | --- |
| Read and write EML/MIME files | `OfficeIMO.Email` |
| Read and write Outlook MSG/MAPI files | `OfficeIMO.Email` |
| Outlook appointments, contacts, tasks, journals, and notes | `OfficeIMO.Email` |
| Read and write TNEF/`winmail.dat` | `OfficeIMO.Email` |
| Read and write mbox archives | `OfficeIMO.Email` |
| Compressed RTF body handling | `OfficeIMO.Email` bridge over `OfficeIMO.Rtf` |
| CFB storage needed by MSG | Internal shared OfficeIMO source; not a public general-purpose CFB API |
| SMTP, IMAP, POP3, and network authentication | Mailozaurr over MailKit |
| MIME security, PGP, S/MIME, DKIM, and ARC | Mailozaurr over MimeKit and its security policy |

This split avoids two competing message models. `OfficeIMO.Email.EmailDocument` is the persisted-artifact model. `MimeKit.MimeMessage` remains Mailozaurr's transport and cryptography model where those features are required.

## Mailozaurr package shape

The public types currently built into `Mailozaurr.Msg` already use the `Mailozaurr` namespace. Move these compatibility DTOs and facades into the main `Mailozaurr` assembly:

- `MailFileMessage`, `MailFileAddress`, `MailFileRecipient`, and `MailFileAttachment`
- `MailFileReader` and `MailFileReaderOptions`
- `EmailMessage` conversion helpers and their result types

Their implementation should be a thin mapping over `EmailDocumentReader` and `EmailDocumentWriter`. Do not copy MIME, MAPI, CFB, TNEF, mbox, or RTF logic into Mailozaurr.

Once the main package contains those types, the PowerShell build can stop loading `Mailozaurr.Msg.dll`, `MsgKit.dll`, `MsgReader.dll`, `OpenMcdf.dll`, and `RtfPipe.dll`. Microsoft.Maui.Graphics then disappears with the MsgReader dependency rather than requiring a separate exclusion.

Existing NuGet consumers of `Mailozaurr.Msg` need a release note because the assembly identity changes even though the namespace and public type names can remain stable. If a compatibility release is required, publish one final `Mailozaurr.Msg` version containing type forwarders to the main assembly, then deprecate it. The normal PowerShell distribution does not need to keep loading that compatibility assembly.

## Reader adapter

The current `MailFileReaderOptions` map directly to immutable OfficeIMO options:

```csharp
var officeOptions = new EmailReaderOptions(
    includeAttachmentContent: options.IncludeAttachments && options.IncludeAttachmentContent);

EmailReadResult read = new EmailDocumentReader(officeOptions).Read(fileInfo.FullName);
EmailDocument document = read.Document;
```

The Mailozaurr mapper then projects `EmailDocument` into its compatibility DTO:

- `From`, `Recipients`, `Date`, `ReceivedDate`, `Subject`, `MessageId`, `Body.Text`, and `Body.Html` map directly.
- Filter `Recipients` by `EmailRecipientKind` to populate `To`, `Cc`, and `Bcc`.
- Map attachment metadata regardless of `IncludeAttachmentContent`; expose `Content` only when requested.
- Merge `EmailHeader` values into the existing case-insensitive header dictionary only when `IncludeHeaders` is true. OfficeIMO keeps duplicate headers in source order; the compatibility dictionary remains intentionally lossy.
- Treat an `EmailDiagnosticSeverity.Error` diagnostic as a failed `MailFileReader` operation. Log or return warning diagnostics according to Mailozaurr policy.

`SignatureIsValid`, `SignedBy`, and `SignedOn` must not be inferred by OfficeIMO. For EML workflows that request signature verification, Mailozaurr should run its existing MimeKit verification path against the original source. Leave those compatibility properties null when no verification was requested or completed.

## Conversion adapter

The existing EML-to-MSG and MSG-to-EML commands can keep their public signatures. Only the conversion body changes:

```csharp
EmailDocument document = new EmailDocumentReader().Read(inputFile.FullName).Document;
new EmailDocumentWriter().Write(document, temporaryOutputPath, targetFormat);
```

Keep Mailozaurr's current temporary-file and force/replace behavior around that call. It is command behavior, not an email-format concern.

## Release sequence

1. Publish a three-part `OfficeIMO.Email` package version containing the required EML, MSG, TNEF, mbox, Outlook-item, and RTF behavior.
2. Add that released package to the main Mailozaurr project and move the compatibility surface from `Mailozaurr.Msg` into the main assembly.
3. Replace the MsgReader/MsgKit adapters with thin `OfficeIMO.Email` mappings and keep MimeKit verification as a separate opt-in step.
4. Remove `Mailozaurr.Msg` from the PowerShell library list and remove MsgKit, MsgReader, OpenMcdf, and RtfPipe references.
5. Run the existing conversion and `MailFileReader` contracts against EML and MSG fixtures, then add TNEF, mbox, embedded-message, and typed Outlook-item coverage.
6. Deprecate or unlist the old `Mailozaurr.Msg` package after any chosen type-forwarder transition.

Do not commit a Mailozaurr project reference to an OfficeIMO worktree or pin a private four-part build. Until `OfficeIMO.Email` is published, validate the downstream adapter with a temporary local package feed or an uncommitted local project reference.
