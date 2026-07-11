using OfficeIMO.Email;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Reader;

public static partial class DocumentReader {
    private static void BuildEmailChunks(EmailExtraction extraction, ReaderOptions opt, CancellationToken cancellationToken) {
        var context = new EmailChunkContext(extraction, opt, cancellationToken);
        for (int messageIndex = 0; messageIndex < extraction.Documents.Count; messageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            EmailMailboxEntry? mailboxEntry = messageIndex < extraction.MailboxEntries.Count
                ? extraction.MailboxEntries[messageIndex]
                : null;
            BuildEmailDocumentChunks(
                extraction.Documents[messageIndex],
                context,
                messageIndex,
                depth: 0,
                parentHeading: null,
                logicalPath: BuildEmailMessagePath(extraction.SourceName, extraction.Format, messageIndex),
                mailboxEntry);
        }

        if (extraction.Chunks.Count == 0) {
            IReadOnlyList<string> warnings = BuildEmailDiagnosticWarnings(extraction.Diagnostics);
            extraction.Chunks.Add(new ReaderChunk {
                Id = BuildStableId("email", Path.GetFileName(extraction.SourceName), 0, null),
                Kind = ReaderInputKind.Email,
                Location = new ReaderLocation {
                    Path = extraction.SourceName,
                    BlockIndex = 0,
                    SourceBlockKind = "email-diagnostic",
                    BlockAnchor = "email-diagnostic"
                },
                Text = "The email artifact did not contain a readable message.",
                Warnings = warnings.Count == 0 ? null : warnings
            });
        }
    }

    private static void BuildEmailDocumentChunks(
        EmailDocument document,
        EmailChunkContext context,
        int messageIndex,
        int depth,
        string? parentHeading,
        string logicalPath,
        EmailMailboxEntry? mailboxEntry) {
        context.CancellationToken.ThrowIfCancellationRequested();
        string subject = string.IsNullOrWhiteSpace(document.Subject)
            ? document.OutlookItemKind.ToString()
            : document.Subject!.Trim();
        string heading = string.IsNullOrWhiteSpace(parentHeading)
            ? subject
            : string.Concat(parentHeading, " > ", subject);

        List<string> metadataWarnings = new List<string>();
        if (!context.DiagnosticsAttached) {
            metadataWarnings.AddRange(BuildEmailDiagnosticWarnings(context.Extraction.Diagnostics));
            context.DiagnosticsAttached = true;
        }

        string summary = BuildEmailSummary(document, mailboxEntry, subject);
        summary = LimitEmailChunkText(summary, context.Options.MaxChars, metadataWarnings);
        int metadataBlockIndex = context.NextBlockIndex++;
        context.Extraction.Chunks.Add(new ReaderChunk {
            Id = BuildStableId("email", Path.GetFileName(context.Extraction.SourceName), metadataBlockIndex, messageIndex),
            Kind = ReaderInputKind.Email,
            Location = new ReaderLocation {
                Path = logicalPath,
                BlockIndex = metadataBlockIndex,
                SourceBlockIndex = messageIndex,
                HeadingPath = heading,
                SourceBlockKind = "email-message",
                BlockAnchor = BuildEmailAnchor("message", messageIndex, depth)
            },
            Text = summary,
            Markdown = summary,
            Warnings = metadataWarnings.Count == 0 ? null : metadataWarnings.ToArray()
        });

        AddEmailBodyChunks(document, context, messageIndex, depth, heading, logicalPath);

        for (int attachmentIndex = 0; attachmentIndex < document.Attachments.Count; attachmentIndex++) {
            context.CancellationToken.ThrowIfCancellationRequested();
            EmailAttachment attachment = document.Attachments[attachmentIndex];
            AddEmailAttachment(attachment, context, messageIndex, attachmentIndex, depth, heading, logicalPath);
        }
    }

    private static void AddEmailBodyChunks(
        EmailDocument document,
        EmailChunkContext context,
        int messageIndex,
        int depth,
        string heading,
        string logicalPath) {
        string? body = document.Body.Text;
        string bodyKind = "plain text";
        string? bodyWarning = null;
        if (string.IsNullOrEmpty(body) && !string.IsNullOrEmpty(document.Body.Html)) {
            body = document.Body.Html;
            bodyKind = "HTML";
            bodyWarning = "EMAIL_HTML_BODY_PRESERVED: No plain-text alternative was available; the HTML body is preserved without lossy tag stripping.";
        } else if (string.IsNullOrEmpty(body) && !string.IsNullOrEmpty(document.Body.Rtf)) {
            if (TryAddEmailRtfBodyChunks(document.Body.Rtf!, context, messageIndex, depth, heading, logicalPath,
                out bodyWarning)) return;
            body = document.Body.Rtf;
            bodyKind = "RTF";
            bodyWarning = bodyWarning ?? "EMAIL_RTF_BODY_PRESERVED: Register OfficeIMO.Reader.Rtf to extract the preserved RTF body semantically.";
        }
        if (string.IsNullOrEmpty(body)) {
            return;
        }

        int bodyPartIndex = 0;
        foreach (ReaderChunk sourceChunk in ChunkPlainTextFromText(
            body!, logicalPath, Path.GetFileName(context.Extraction.SourceName), context.Options,
            ReaderInputKind.Email, context.CancellationToken, treatAsMarkdown: false)) {
            int blockIndex = context.NextBlockIndex++;
            var warnings = sourceChunk.Warnings == null
                ? new List<string>()
                : new List<string>(sourceChunk.Warnings);
            if (bodyWarning != null) {
                warnings.Add(bodyWarning);
            }
            sourceChunk.Id = BuildStableId("email-body", Path.GetFileName(context.Extraction.SourceName), blockIndex, bodyPartIndex);
            sourceChunk.Kind = ReaderInputKind.Email;
            sourceChunk.Location.Path = logicalPath;
            sourceChunk.Location.BlockIndex = blockIndex;
            sourceChunk.Location.SourceBlockIndex = bodyPartIndex;
            sourceChunk.Location.HeadingPath = string.Concat(heading, " > Body");
            sourceChunk.Location.SourceBlockKind = "email-body";
            sourceChunk.Location.BlockAnchor = BuildEmailAnchor("body", messageIndex, depth) + "-" + bodyPartIndex.ToString(CultureInfo.InvariantCulture);
            sourceChunk.Warnings = warnings.Count == 0 ? null : warnings.ToArray();
            if (string.Equals(bodyKind, "plain text", StringComparison.Ordinal)) {
                sourceChunk.Markdown = sourceChunk.Text;
            }
            context.Extraction.Chunks.Add(sourceChunk);
            bodyPartIndex++;
        }
    }

    private static bool TryAddEmailRtfBodyChunks(
        string rtf,
        EmailChunkContext context,
        int messageIndex,
        int depth,
        string heading,
        string logicalPath,
        out string? warning) {
        warning = null;
        const string sourceName = "email-body.rtf";
        if (!TryResolveCustomHandlerBySourceName(sourceName, out ReaderHandlerDescriptor handler) ||
            (handler.ReadStream == null && handler.ReadDocumentStream == null)) return false;

        byte[] bytes = new byte[rtf.Length];
        for (int index = 0; index < rtf.Length; index++) {
            if (rtf[index] > byte.MaxValue) {
                warning = "EMAIL_RTF_BODY_ENCODING_INVALID: The preserved RTF contains a character that is not byte-preserving.";
                return false;
            }
            bytes[index] = unchecked((byte)rtf[index]);
        }

        try {
            ReaderOptions rtfOptions = CloneOptions(context.Options, computeHashes: false);
            ReaderChunk[] rtfChunks = Read(bytes, sourceName, rtfOptions, context.CancellationToken).ToArray();
            if (rtfChunks.Length == 0) return false;
            for (int index = 0; index < rtfChunks.Length; index++) {
                ReaderChunk chunk = rtfChunks[index];
                int blockIndex = context.NextBlockIndex++;
                chunk.Id = BuildStableId("email-body-rtf", Path.GetFileName(context.Extraction.SourceName), blockIndex, index);
                chunk.Location.Path = logicalPath;
                chunk.Location.BlockIndex = blockIndex;
                chunk.Location.SourceBlockIndex = index;
                chunk.Location.HeadingPath = string.Concat(heading, " > Body");
                chunk.Location.SourceBlockKind = "email-body-rtf";
                chunk.Location.BlockAnchor = BuildEmailAnchor("body", messageIndex, depth) + "-rtf-" +
                    index.ToString(CultureInfo.InvariantCulture);
                chunk.SourceId = null;
                chunk.SourceHash = null;
                chunk.ChunkHash = null;
                chunk.SourceLastWriteUtc = null;
                chunk.SourceLengthBytes = null;
                context.Extraction.Chunks.Add(chunk);
            }
            return true;
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception exception) {
            warning = string.Concat("EMAIL_RTF_BODY_READER_FAILED: ", exception.GetType().Name,
                " while extracting the preserved RTF body.");
            return false;
        }
    }

    private static void AddEmailAttachment(
        EmailAttachment attachment,
        EmailChunkContext context,
        int messageIndex,
        int attachmentIndex,
        int depth,
        string heading,
        string logicalPath) {
        string fileName = string.IsNullOrWhiteSpace(attachment.FileName)
            ? string.Concat("attachment-", (attachmentIndex + 1).ToString(CultureInfo.InvariantCulture))
            : attachment.FileName!.Trim();
        string attachmentPath = string.Concat(logicalPath, "!/", fileName);
        var warnings = new List<string>();
        string summary = BuildEmailAttachmentSummary(attachment, fileName);
        summary = LimitEmailChunkText(summary, context.Options.MaxChars, warnings);
        int blockIndex = context.NextBlockIndex++;
        var attachmentChunk = new ReaderChunk {
            Id = BuildStableId("email-attachment", Path.GetFileName(context.Extraction.SourceName), blockIndex, attachmentIndex),
            Kind = ReaderInputKind.Email,
            Location = new ReaderLocation {
                Path = attachmentPath,
                BlockIndex = blockIndex,
                SourceBlockIndex = attachmentIndex,
                HeadingPath = string.Concat(heading, " > Attachment: ", fileName),
                SourceBlockKind = "email-attachment",
                BlockAnchor = BuildEmailAnchor("attachment", messageIndex, depth) + "-" + attachmentIndex.ToString(CultureInfo.InvariantCulture)
            },
            Text = summary,
            Markdown = summary,
            Warnings = warnings.Count == 0 ? null : warnings.ToArray()
        };
        context.Extraction.Chunks.Add(attachmentChunk);

        AddEmailAttachmentAsset(attachment, context, attachmentIndex, attachmentPath, attachmentChunk.Location);

        if (attachment.EmbeddedDocument != null) {
            BuildEmailDocumentChunks(
                attachment.EmbeddedDocument,
                context,
                messageIndex,
                depth + 1,
                attachmentChunk.Location.HeadingPath,
                attachmentPath,
                mailboxEntry: null);
            return;
        }

        if (attachment.Content == null || attachment.Content.Length == 0 ||
            !CanReadEmailAttachmentContent(fileName)) {
            return;
        }

        try {
            ReaderOptions attachmentOptions = CloneOptions(context.Options, computeHashes: false);
            ReaderChunk[] childChunks = Read(attachment.Content, fileName, attachmentOptions, context.CancellationToken).ToArray();
            for (int childIndex = 0; childIndex < childChunks.Length; childIndex++) {
                ReaderChunk child = childChunks[childIndex];
                int childBlockIndex = context.NextBlockIndex++;
                child.Id = BuildStableId("email-attachment-content", fileName, childBlockIndex, child.Location.SourceBlockIndex);
                child.Location = CloneEmailAttachmentLocation(
                    child.Location,
                    attachmentPath,
                    attachmentChunk.Location.HeadingPath,
                    childBlockIndex);
                child.SourceId = null;
                child.SourceHash = null;
                child.ChunkHash = null;
                child.SourceLastWriteUtc = null;
                child.SourceLengthBytes = null;
                context.Extraction.Chunks.Add(child);
            }
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception exception) {
            warnings.Add(string.Concat("EMAIL_ATTACHMENT_READER_FAILED: ", exception.GetType().Name,
                " while extracting ", fileName, "."));
            attachmentChunk.Warnings = warnings.ToArray();
        }
    }

    private static void AddEmailAttachmentAsset(
        EmailAttachment attachment,
        EmailChunkContext context,
        int attachmentIndex,
        string attachmentPath,
        ReaderLocation location) {
        string fileName = GetLogicalFileName(attachmentPath);
        byte[]? payload = attachment.Content;
        string kind = attachment.EmbeddedDocument != null
            ? "embedded-message"
            : attachment.StructuredStorageStreams.Count > 0
                ? "ole-storage"
                : !string.IsNullOrWhiteSpace(attachment.ContentType) && attachment.ContentType!.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
                    ? "image"
                    : attachment.IsInline ? "inline-attachment" : "attachment";
        string? extension = null;
        try {
            extension = Path.GetExtension(fileName);
        } catch (ArgumentException) {
            // Keep the original filename even when it is not a valid local path.
        }

        int assetIndex = context.Extraction.Assets.Count;
        context.Extraction.Assets.Add(new OfficeDocumentAsset {
            Id = string.Concat("email-asset-", assetIndex.ToString("D4", CultureInfo.InvariantCulture)),
            Kind = kind,
            MediaType = attachment.ContentType,
            Extension = string.IsNullOrWhiteSpace(extension) ? null : extension,
            FileName = fileName,
            Title = fileName,
            LengthBytes = attachment.Length,
            PayloadHash = payload == null ? null : OfficeDocumentAssetHash.ComputeSha256Hex(payload),
            PayloadBytes = payload,
            SourceObjectId = attachmentPath,
            Location = location
        });
    }

    private static string GetLogicalFileName(string path) {
        int separator = Math.Max(path.LastIndexOf('/'), path.LastIndexOf('\\'));
        return separator < 0 ? path : path.Substring(separator + 1);
    }

    private static bool CanReadEmailAttachmentContent(string fileName) {
        ReaderInputKind kind = DetectKind(fileName);
        if (kind == ReaderInputKind.Email) {
            return false;
        }
        if (kind != ReaderInputKind.Unknown) {
            return true;
        }
        return TryResolveCustomHandlerBySourceName(fileName, out ReaderHandlerDescriptor handler) &&
            (handler.ReadStream != null || handler.ReadDocumentStream != null);
    }

    private static ReaderLocation CloneEmailAttachmentLocation(ReaderLocation source, string path, string? parentHeading, int blockIndex) {
        string? headingPath = string.IsNullOrWhiteSpace(source.HeadingPath)
            ? parentHeading
            : string.Concat(parentHeading, " > ", source.HeadingPath);
        return new ReaderLocation {
            Path = path,
            BlockIndex = blockIndex,
            SourceBlockIndex = source.SourceBlockIndex,
            StartLine = source.StartLine,
            EndLine = source.EndLine,
            NormalizedStartLine = source.NormalizedStartLine,
            NormalizedEndLine = source.NormalizedEndLine,
            HeadingPath = headingPath,
            HeadingSlug = source.HeadingSlug,
            SourceBlockKind = source.SourceBlockKind,
            BlockAnchor = source.BlockAnchor,
            Sheet = source.Sheet,
            A1Range = source.A1Range,
            Slide = source.Slide,
            Page = source.Page,
            TableIndex = source.TableIndex
        };
    }

    private static string BuildEmailSummary(EmailDocument document, EmailMailboxEntry? mailboxEntry, string subject) {
        var builder = new StringBuilder();
        builder.Append("# ").AppendLine(subject);
        AppendEmailSummaryValue(builder, "Format", document.Format.ToString());
        AppendEmailSummaryValue(builder, "Outlook item", document.OutlookItemKind.ToString());
        AppendEmailSummaryValue(builder, "Message class", document.MessageClass);
        AppendEmailSummaryValue(builder, "From", document.From?.ToString());
        AppendEmailSummaryValue(builder, "Sender", document.Sender?.ToString());
        AppendEmailSummaryValue(builder, "To", JoinEmailRecipients(document, EmailRecipientKind.To));
        AppendEmailSummaryValue(builder, "Cc", JoinEmailRecipients(document, EmailRecipientKind.Cc));
        AppendEmailSummaryValue(builder, "Bcc", JoinEmailRecipients(document, EmailRecipientKind.Bcc));
        AppendEmailSummaryValue(builder, "Date", FormatEmailDate(document.Date));
        AppendEmailSummaryValue(builder, "Received", FormatEmailDate(document.ReceivedDate));
        AppendEmailSummaryValue(builder, "Message-ID", document.MessageId);
        if (mailboxEntry != null) {
            AppendEmailSummaryValue(builder, "Envelope sender", mailboxEntry.EnvelopeSender);
            AppendEmailSummaryValue(builder, "Envelope date", FormatEmailDate(mailboxEntry.EnvelopeDate));
        }
        AppendEmailSummaryValue(builder, "Attachments", document.Attachments.Count.ToString(CultureInfo.InvariantCulture));
        AppendOutlookItemSummary(builder, document);
        return builder.ToString().TrimEnd();
    }

    private static string BuildEmailAttachmentSummary(EmailAttachment attachment, string fileName) {
        var builder = new StringBuilder();
        builder.Append("## Attachment: ").AppendLine(fileName);
        AppendEmailSummaryValue(builder, "Content type", attachment.ContentType);
        AppendEmailSummaryValue(builder, "Length", attachment.Length.ToString(CultureInfo.InvariantCulture));
        AppendEmailSummaryValue(builder, "Content-ID", attachment.ContentId);
        AppendEmailSummaryValue(builder, "Content location", attachment.ContentLocation);
        AppendEmailSummaryValue(builder, "Inline", attachment.IsInline ? "true" : "false");
        AppendEmailSummaryValue(builder, "Embedded item", attachment.EmbeddedDocument == null ? null : "true");
        AppendEmailSummaryValue(builder, "Storage streams", attachment.StructuredStorageStreams.Count == 0
            ? null
            : attachment.StructuredStorageStreams.Count.ToString(CultureInfo.InvariantCulture));
        return builder.ToString().TrimEnd();
    }

    private static void AppendOutlookItemSummary(StringBuilder builder, EmailDocument document) {
        if (document.Appointment != null) {
            AppendEmailSummaryValue(builder, "Start", FormatEmailDate(document.Appointment.Start));
            AppendEmailSummaryValue(builder, "End", FormatEmailDate(document.Appointment.End));
            AppendEmailSummaryValue(builder, "Location", document.Appointment.Location);
            AppendEmailSummaryValue(builder, "All day", FormatNullableBoolean(document.Appointment.IsAllDay));
            AppendEmailSummaryValue(builder, "Recurrence", document.Appointment.RecurrencePattern);
        }
        if (document.Contact != null) {
            AppendEmailSummaryValue(builder, "Given name", document.Contact.GivenName);
            AppendEmailSummaryValue(builder, "Surname", document.Contact.Surname);
            AppendEmailSummaryValue(builder, "Company", document.Contact.CompanyName);
            AppendEmailSummaryValue(builder, "Job title", document.Contact.JobTitle);
            AppendEmailSummaryValue(builder, "Business phone", document.Contact.BusinessPhone);
            AppendEmailSummaryValue(builder, "Home phone", document.Contact.HomePhone);
            AppendEmailSummaryValue(builder, "Mobile phone", document.Contact.MobilePhone);
            AppendEmailSummaryValue(builder, "Email", document.Contact.Email1Address);
        }
        if (document.Task != null) {
            AppendEmailSummaryValue(builder, "Task start", FormatEmailDate(document.Task.Start));
            AppendEmailSummaryValue(builder, "Task due", FormatEmailDate(document.Task.Due));
            AppendEmailSummaryValue(builder, "Task owner", document.Task.Owner);
            AppendEmailSummaryValue(builder, "Task complete", FormatNullableBoolean(document.Task.IsComplete));
            AppendEmailSummaryValue(builder, "Task percent", document.Task.PercentComplete?.ToString("0.####", CultureInfo.InvariantCulture));
        }
        if (document.Journal != null) {
            AppendEmailSummaryValue(builder, "Journal start", FormatEmailDate(document.Journal.Start));
            AppendEmailSummaryValue(builder, "Journal end", FormatEmailDate(document.Journal.End));
            AppendEmailSummaryValue(builder, "Journal type", document.Journal.Type);
        }
        if (document.Note != null) {
            AppendEmailSummaryValue(builder, "Note color", document.Note.Color?.ToString(CultureInfo.InvariantCulture));
            AppendEmailSummaryValue(builder, "Note width", document.Note.Width?.ToString(CultureInfo.InvariantCulture));
            AppendEmailSummaryValue(builder, "Note height", document.Note.Height?.ToString(CultureInfo.InvariantCulture));
        }
    }

    private static void AppendEmailSummaryValue(StringBuilder builder, string name, string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return;
        }
        builder.Append("- ").Append(name).Append(": ").AppendLine(CollapseEmailLine(value!));
    }

    private static string JoinEmailRecipients(EmailDocument document, EmailRecipientKind kind) {
        return string.Join(", ", document.Recipients
            .Where(recipient => recipient.Kind == kind)
            .Select(recipient => recipient.Address.ToString())
            .Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    private static string? FormatEmailDate(DateTimeOffset? value) {
        return value?.ToString("O", CultureInfo.InvariantCulture);
    }

    private static string? FormatNullableBoolean(bool? value) {
        return value.HasValue ? (value.Value ? "true" : "false") : null;
    }

    private static string CollapseEmailLine(string value) {
        return value.Replace("\r", " ").Replace("\n", " ").Trim();
    }

    private static string LimitEmailChunkText(string value, int maxChars, List<string> warnings) {
        if (value.Length <= maxChars) {
            return value;
        }
        warnings.Add(string.Concat("EMAIL_CHUNK_TRUNCATED: Metadata exceeded MaxChars (", maxChars.ToString(CultureInfo.InvariantCulture), ")."));
        return value.Substring(0, maxChars);
    }

    private static IReadOnlyList<string> BuildEmailDiagnosticWarnings(IReadOnlyList<EmailDiagnostic> diagnostics) {
        var warnings = new List<string>();
        for (int index = 0; index < diagnostics.Count; index++) {
            EmailDiagnostic diagnostic = diagnostics[index];
            if (diagnostic.Severity == EmailDiagnosticSeverity.Information) {
                continue;
            }
            warnings.Add(string.Concat(diagnostic.Code, ": ", diagnostic.Message));
        }
        return warnings;
    }

    private static string BuildEmailMessagePath(string sourceName, EmailFileFormat format, int messageIndex) {
        return format == EmailFileFormat.Mbox
            ? string.Concat(sourceName, "!/message-", (messageIndex + 1).ToString("D6", CultureInfo.InvariantCulture), ".eml")
            : sourceName;
    }

    private static string BuildEmailAnchor(string kind, int messageIndex, int depth) {
        return string.Concat("email-", kind, "-m", messageIndex.ToString(CultureInfo.InvariantCulture),
            "-d", depth.ToString(CultureInfo.InvariantCulture));
    }

    private sealed class EmailChunkContext {
        internal EmailChunkContext(EmailExtraction extraction, ReaderOptions options, CancellationToken cancellationToken) {
            Extraction = extraction;
            Options = options;
            CancellationToken = cancellationToken;
        }

        internal EmailExtraction Extraction { get; }
        internal ReaderOptions Options { get; }
        internal CancellationToken CancellationToken { get; }
        internal int NextBlockIndex { get; set; }
        internal bool DiagnosticsAttached { get; set; }
    }
}
