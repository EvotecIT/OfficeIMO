using OfficeIMO.Email;

namespace OfficeIMO.Reader.Email;

internal static class EmailReaderProjection {
    internal static IReadOnlyList<ReaderChunk> ProjectEmailDocumentToChunks(
        EmailDocument document, string logicalPath, IReadOnlyList<EmailDiagnostic> diagnostics,
        string sourceName, ReaderOptions options, EmailDocumentProjectionCursor cursor,
        CancellationToken cancellationToken) {
        var projection = new Projection(sourceName, document.Format);
        projection.Diagnostics.AddRange(diagnostics);
        AddDocument(document, null, logicalPath, projection, options, cursor, depth: 0, cancellationToken);
        return projection.Chunks;
    }

    internal static IReadOnlyList<ReaderChunk> ProjectEmailDocumentsToChunks(
        IReadOnlyList<EmailDocument> documents, IReadOnlyList<string?> logicalPaths,
        IReadOnlyList<EmailDiagnostic> diagnostics, EmailFileFormat format, string sourceName,
        ReaderOptions options, CancellationToken cancellationToken) =>
        CreateProjection(documents, null, logicalPaths, diagnostics, format, sourceName, options, cancellationToken).Chunks;

    internal static OfficeDocumentReadResult ProjectEmailDocumentsToPathResult(
        IReadOnlyList<EmailDocument> documents, IReadOnlyList<string?> logicalPaths,
        IReadOnlyList<EmailDiagnostic> diagnostics, EmailFileFormat format, string sourceName,
        string path, ReaderOptions options, CancellationToken cancellationToken,
        bool? computeSourceHash = null) {
        Projection projection = CreateProjection(documents, null, logicalPaths, diagnostics, format, sourceName, options, cancellationToken);
        var source = new OfficeDocumentSource {
            Path = path,
            SourceId = "src:" + Hash(NormalizeSourceKey(path)),
            SourceHash = (computeSourceHash ?? options.ComputeHashes) ? TryHashFile(path) : null,
            LengthBytes = TryLength(path),
            LastWriteUtc = TryLastWrite(path)
        };
        EnrichChunks(projection.Chunks, source, options.ComputeHashes);
        return CreateResult(projection, sourceName, source);
    }

    internal static OfficeDocumentReadResult ProjectEmailDocumentsToStreamResult(
        IReadOnlyList<EmailDocument> documents, IReadOnlyList<string?> logicalPaths,
        IReadOnlyList<EmailDiagnostic> diagnostics, EmailFileFormat format, string sourceName,
        Stream stream, ReaderOptions options, CancellationToken cancellationToken,
        bool? computeSourceHash = null) {
        Projection projection = CreateProjection(documents, null, logicalPaths, diagnostics, format, sourceName, options, cancellationToken);
        var source = new OfficeDocumentSource {
            Path = sourceName,
            SourceId = "src:" + Hash(sourceName),
            SourceHash = (computeSourceHash ?? options.ComputeHashes) ? TryHashStream(stream) : null,
            LengthBytes = TryLength(stream)
        };
        EnrichChunks(projection.Chunks, source, options.ComputeHashes);
        return CreateResult(projection, sourceName, source);
    }

    internal static OfficeDocumentReadResult ProjectMailboxToPathResult(
        EmailMailboxReadResult mailbox, string path, ReaderOptions options, CancellationToken cancellationToken) {
        EmailDocument[] documents = mailbox.Mailbox.Messages.Select(static entry => entry.Document).ToArray();
        EmailMailboxEntry[] entries = mailbox.Mailbox.Messages.ToArray();
        string?[] paths = Enumerable.Range(1, documents.Length).Select(index => (string?)string.Concat(path, "!/message-", index.ToString("D6", CultureInfo.InvariantCulture), ".eml")).ToArray();
        Projection projection = CreateProjection(documents, entries, paths, mailbox.Diagnostics, EmailFileFormat.Mbox, path, options, cancellationToken);
        var source = new OfficeDocumentSource { Path = path, SourceId = "src:" + Hash(NormalizeSourceKey(path)), SourceHash = options.ComputeHashes ? TryHashFile(path) : null, LengthBytes = TryLength(path), LastWriteUtc = TryLastWrite(path) };
        EnrichChunks(projection.Chunks, source, options.ComputeHashes);
        return CreateResult(projection, path, source);
    }

    internal static OfficeDocumentReadResult ProjectMailboxToStreamResult(
        EmailMailboxReadResult mailbox, string sourceName, Stream stream, ReaderOptions options, CancellationToken cancellationToken) {
        EmailDocument[] documents = mailbox.Mailbox.Messages.Select(static entry => entry.Document).ToArray();
        EmailMailboxEntry[] entries = mailbox.Mailbox.Messages.ToArray();
        string?[] paths = Enumerable.Range(1, documents.Length).Select(index => (string?)string.Concat(sourceName, "!/message-", index.ToString("D6", CultureInfo.InvariantCulture), ".eml")).ToArray();
        Projection projection = CreateProjection(documents, entries, paths, mailbox.Diagnostics, EmailFileFormat.Mbox, sourceName, options, cancellationToken);
        var source = new OfficeDocumentSource { Path = sourceName, SourceId = "src:" + Hash(sourceName), SourceHash = options.ComputeHashes ? TryHashStream(stream) : null, LengthBytes = TryLength(stream) };
        EnrichChunks(projection.Chunks, source, options.ComputeHashes);
        return CreateResult(projection, sourceName, source);
    }

    private static Projection CreateProjection(
        IReadOnlyList<EmailDocument> documents, IReadOnlyList<EmailMailboxEntry>? mailboxEntries,
        IReadOnlyList<string?> logicalPaths, IReadOnlyList<EmailDiagnostic> diagnostics,
        EmailFileFormat format, string sourceName, ReaderOptions options,
        CancellationToken cancellationToken) {
        if (documents.Count != logicalPaths.Count) throw new ArgumentException("A logical path is required for every email document.", nameof(logicalPaths));
        var projection = new Projection(sourceName, format);
        projection.Documents.AddRange(documents);
        projection.Diagnostics.AddRange(diagnostics);
        var cursor = new EmailDocumentProjectionCursor();
        for (int index = 0; index < documents.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            EmailMailboxEntry? entry = mailboxEntries != null && index < mailboxEntries.Count ? mailboxEntries[index] : null;
            projection.MailboxEntries.Add(entry);
            AddDocument(documents[index], entry, logicalPaths[index] ?? sourceName, projection, options, cursor, 0, cancellationToken);
        }
        if (projection.Chunks.Count == 0) {
            projection.Chunks.Add(new ReaderChunk { Id = "email:diagnostic:0000", Kind = ReaderInputKind.Email, Location = new ReaderLocation { Path = sourceName, BlockIndex = 0, SourceBlockKind = "email-diagnostic" }, Text = "The email artifact did not contain a readable message.", Warnings = DiagnosticsToWarnings(diagnostics) });
        }
        return projection;
    }

    private static void AddDocument(EmailDocument document, EmailMailboxEntry? mailboxEntry, string logicalPath,
        Projection projection, ReaderOptions options, EmailDocumentProjectionCursor cursor, int depth,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        int messageIndex = cursor.NextMessageIndex++;
        string subject = string.IsNullOrWhiteSpace(document.Subject) ? document.OutlookItemKind.ToString() : document.Subject!.Trim();
        int summaryIndex = cursor.NextBlockIndex++;
        var warnings = new List<string>();
        if (!cursor.DiagnosticsAttached) {
            IReadOnlyList<string>? diagnosticWarnings = DiagnosticsToWarnings(projection.Diagnostics);
            if (diagnosticWarnings != null) warnings.AddRange(diagnosticWarnings);
            cursor.DiagnosticsAttached = true;
        }
        string summary = Limit(BuildSummary(document, mailboxEntry, subject), options.MaxChars, warnings, "Email summary");
        projection.Chunks.Add(new ReaderChunk {
            Id = $"email:message:{messageIndex.ToString("D6", CultureInfo.InvariantCulture)}",
            Kind = ReaderInputKind.Email,
            Location = new ReaderLocation { Path = logicalPath, BlockIndex = summaryIndex, SourceBlockIndex = messageIndex, HeadingPath = subject, SourceBlockKind = "email-message", BlockAnchor = $"email-message-m{messageIndex}-d{depth}" },
            Text = summary, Markdown = summary, Warnings = warnings.Count == 0 ? null : warnings
        });

        string? body = document.Body.Text;
        string bodyKind = "plain";
        if (string.IsNullOrEmpty(body) && !string.IsNullOrEmpty(document.Body.Html)) { body = document.Body.Html; bodyKind = "html"; }
        if (string.IsNullOrEmpty(body) && !string.IsNullOrEmpty(document.Body.Rtf)) { body = document.Body.Rtf; bodyKind = "rtf"; }
        if (!string.IsNullOrEmpty(body) &&
            !TryAddSemanticBody(body!, bodyKind, logicalPath, subject, messageIndex,
                projection, options, cursor, cancellationToken)) {
            AddBody(body!, bodyKind, logicalPath, subject, messageIndex, projection, options, cursor);
        }

        for (int attachmentIndex = 0; attachmentIndex < document.Attachments.Count; attachmentIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            EmailAttachment attachment = document.Attachments[attachmentIndex];
            string name = string.IsNullOrWhiteSpace(attachment.FileName) ? "attachment-" + (attachmentIndex + 1).ToString(CultureInfo.InvariantCulture) : attachment.FileName!.Trim();
            string attachmentPath = logicalPath + "!/" + name;
            int blockIndex = cursor.NextBlockIndex++;
            string attachmentSummary = Limit(BuildAttachmentSummary(attachment, name), options.MaxChars, warnings = new List<string>(), "Attachment summary");
            var location = new ReaderLocation { Path = attachmentPath, BlockIndex = blockIndex, SourceBlockIndex = attachmentIndex, HeadingPath = subject + " > Attachment: " + name, SourceBlockKind = "email-attachment", BlockAnchor = $"email-attachment-m{messageIndex}-d{depth}-{attachmentIndex}" };
            var attachmentChunk = new ReaderChunk { Id = $"email:attachment:{messageIndex.ToString("D6", CultureInfo.InvariantCulture)}:{attachmentIndex.ToString("D4", CultureInfo.InvariantCulture)}", Kind = ReaderInputKind.Email, Location = location, Text = attachmentSummary, Markdown = attachmentSummary, Warnings = warnings.Count == 0 ? null : warnings };
            projection.Chunks.Add(attachmentChunk);
            byte[]? payload = attachment.Content;
            projection.Assets.Add(new OfficeDocumentAsset {
                Id = "email-asset-" + projection.Assets.Count.ToString("D6", CultureInfo.InvariantCulture),
                Kind = attachment.EmbeddedDocument != null ? "embedded-message" : attachment.IsInline ? "inline-attachment" : "attachment",
                MediaType = attachment.ContentType,
                Extension = TryExtension(name),
                FileName = name,
                Title = name,
                LengthBytes = attachment.Length,
                PayloadHash = payload == null ? null : Hash(payload),
                PayloadBytes = payload,
                SourceObjectId = attachmentPath,
                Location = location
            });
            if (attachment.EmbeddedDocument != null) {
                AddDocument(attachment.EmbeddedDocument, null, attachmentPath, projection, options, cursor, depth + 1, cancellationToken);
                continue;
            }
            AddAttachmentContent(attachment, name, attachmentPath, subject, attachmentChunk,
                projection, options, cursor, cancellationToken);
        }
    }

    private static bool TryAddSemanticBody(
        string body,
        string bodyKind,
        string logicalPath,
        string subject,
        int messageIndex,
        Projection projection,
        ReaderOptions options,
        EmailDocumentProjectionCursor cursor,
        CancellationToken cancellationToken) {
        string sourceName = bodyKind == "html" ? "email-body.html" : bodyKind == "rtf" ? "email-body.rtf" : string.Empty;
        if (sourceName.Length == 0 || !ReaderNestedContent.CanRead(sourceName)) return false;
        byte[] bytes = bodyKind == "rtf" ? ToBytePreservingRtf(body) : Encoding.UTF8.GetBytes(body);
        try {
            using var stream = new MemoryStream(bytes, writable: false);
            IReadOnlyList<ReaderChunk> nested = ReaderNestedContent.Read(stream, sourceName,
                CloneWithoutHashes(options), cancellationToken);
            for (int index = 0; index < nested.Count; index++) {
                ReaderChunk chunk = nested[index];
                int blockIndex = cursor.NextBlockIndex++;
                chunk.Id = $"email:body:{messageIndex.ToString("D6", CultureInfo.InvariantCulture)}:{bodyKind}:{index.ToString("D4", CultureInfo.InvariantCulture)}";
                chunk.Location = CloneNestedLocation(chunk.Location, logicalPath, subject + " > Body", blockIndex,
                    "email-body-" + bodyKind);
                ClearNestedSource(chunk);
                projection.Chunks.Add(chunk);
            }
            return nested.Count > 0;
        } catch (OperationCanceledException) {
            throw;
        } catch {
            return false;
        }
    }

    private static void AddAttachmentContent(
        EmailAttachment attachment,
        string fileName,
        string attachmentPath,
        string subject,
        ReaderChunk attachmentChunk,
        Projection projection,
        ReaderOptions options,
        EmailDocumentProjectionCursor cursor,
        CancellationToken cancellationToken) {
        if ((attachment.Content == null || attachment.Content.Length == 0) && attachment.ContentSource == null) return;
        string sourceName = ResolveAttachmentSourceName(fileName, attachment.ContentType);
        try {
            using Stream stream = attachment.OpenContentStream();
            IReadOnlyList<ReaderChunk> nested;
            if (ReaderNestedContent.CanRead(sourceName)) {
                nested = ReaderNestedContent.Read(stream, sourceName, CloneWithoutHashes(options), cancellationToken);
            } else if (IsPlainTextAttachment(sourceName, attachment.ContentType)) {
                using var reader = new StreamReader(stream, Encoding.UTF8, true, 4096, leaveOpen: true);
                string text = reader.ReadToEnd();
                nested = BuildPlainAttachmentChunks(text, options.MaxChars);
            } else {
                return;
            }

            for (int index = 0; index < nested.Count; index++) {
                ReaderChunk child = nested[index];
                int blockIndex = cursor.NextBlockIndex++;
                child.Id = $"email:attachment-content:{blockIndex.ToString("D6", CultureInfo.InvariantCulture)}:{index.ToString("D4", CultureInfo.InvariantCulture)}";
                child.Location = CloneNestedLocation(child.Location, attachmentPath,
                    subject + " > Attachment: " + fileName, blockIndex, child.Location.SourceBlockKind);
                ClearNestedSource(child);
                projection.Chunks.Add(child);
            }
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception exception) {
            var warnings = new List<string>(attachmentChunk.Warnings ?? Array.Empty<string>()) {
                $"EMAIL_ATTACHMENT_READER_FAILED: {exception.GetType().Name} while extracting {fileName}."
            };
            attachmentChunk.Warnings = warnings;
        }
    }

    private static IReadOnlyList<ReaderChunk> BuildPlainAttachmentChunks(string text, int maxChars) {
        int limit = Math.Max(256, maxChars);
        var chunks = new List<ReaderChunk>();
        for (int offset = 0, index = 0; offset < text.Length; offset += limit, index++) {
            string part = text.Substring(offset, Math.Min(limit, text.Length - offset));
            chunks.Add(new ReaderChunk {
                Id = "text:" + index.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Text,
                Location = new ReaderLocation { BlockIndex = index, SourceBlockIndex = index, SourceBlockKind = "text" },
                Text = part,
                Markdown = part
            });
        }
        return chunks;
    }

    private static ReaderLocation CloneNestedLocation(ReaderLocation source, string path, string heading, int blockIndex, string? sourceBlockKind) =>
        new ReaderLocation {
            Path = path,
            BlockIndex = blockIndex,
            SourceBlockIndex = source.SourceBlockIndex,
            StartLine = source.StartLine,
            EndLine = source.EndLine,
            NormalizedStartLine = source.NormalizedStartLine,
            NormalizedEndLine = source.NormalizedEndLine,
            HeadingPath = string.IsNullOrWhiteSpace(source.HeadingPath) ? heading : heading + " > " + source.HeadingPath,
            HeadingSlug = source.HeadingSlug,
            SourceBlockKind = string.IsNullOrWhiteSpace(sourceBlockKind) ? source.SourceBlockKind : sourceBlockKind,
            BlockAnchor = source.BlockAnchor,
            Sheet = source.Sheet,
            A1Range = source.A1Range,
            Slide = source.Slide,
            Page = source.Page,
            TableIndex = source.TableIndex
        };

    private static void ClearNestedSource(ReaderChunk chunk) {
        chunk.SourceId = null;
        chunk.SourceHash = null;
        chunk.ChunkHash = null;
        chunk.SourceLastWriteUtc = null;
        chunk.SourceLengthBytes = null;
    }

    private static ReaderOptions CloneWithoutHashes(ReaderOptions source) => new ReaderOptions {
        MaxInputBytes = source.MaxInputBytes,
        OpenXmlMaxCharactersInPart = source.OpenXmlMaxCharactersInPart,
        MaxOpenXmlImageAssets = source.MaxOpenXmlImageAssets,
        OpenPassword = source.OpenPassword,
        MaxOpenXmlImagePlacementsPerRelationship = source.MaxOpenXmlImagePlacementsPerRelationship,
        MaxOpenXmlImageAssetBytes = source.MaxOpenXmlImageAssetBytes,
        MaxOpenXmlImageTotalAssetBytes = source.MaxOpenXmlImageTotalAssetBytes,
        MaxChars = source.MaxChars,
        MaxTableRows = source.MaxTableRows,
        ComputeHashes = false,
        DetectionMode = source.DetectionMode,
        DetectionMaxProbeBytes = source.DetectionMaxProbeBytes,
        DetectionMaxContainerEntries = source.DetectionMaxContainerEntries
    };

    private static byte[] ToBytePreservingRtf(string value) {
        var bytes = new byte[value.Length];
        for (int index = 0; index < value.Length; index++) {
            if (value[index] > byte.MaxValue) throw new InvalidDataException("The preserved RTF body is not byte-preserving.");
            bytes[index] = unchecked((byte)value[index]);
        }
        return bytes;
    }

    private static string ResolveAttachmentSourceName(string fileName, string? contentType) {
        if (!string.IsNullOrWhiteSpace(TryExtension(fileName))) return fileName;
        string? extension = contentType?.Trim().ToLowerInvariant() switch {
            "application/pdf" => ".pdf",
            "text/plain" => ".txt",
            "text/html" => ".html",
            "text/rtf" or "application/rtf" => ".rtf",
            "text/csv" => ".csv",
            "text/markdown" => ".md",
            "message/rfc822" => ".eml",
            "application/vnd.ms-outlook" => ".msg",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => ".docx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => ".xlsx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation" => ".pptx",
            _ => null
        };
        return extension == null ? fileName : fileName + extension;
    }

    private static bool IsPlainTextAttachment(string sourceName, string? contentType) =>
        string.Equals(TryExtension(sourceName), ".txt", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(contentType, "text/plain", StringComparison.OrdinalIgnoreCase);

    private static void AddBody(string body, string bodyKind, string logicalPath, string subject, int messageIndex,
        Projection projection, ReaderOptions options, EmailDocumentProjectionCursor cursor) {
        int limit = Math.Max(256, options.MaxChars);
        for (int offset = 0, partIndex = 0; offset < body.Length; offset += limit, partIndex++) {
            string part = body.Substring(offset, Math.Min(limit, body.Length - offset));
            int blockIndex = cursor.NextBlockIndex++;
            projection.Chunks.Add(new ReaderChunk {
                Id = $"email:body:{messageIndex.ToString("D6", CultureInfo.InvariantCulture)}:{partIndex.ToString("D4", CultureInfo.InvariantCulture)}",
                Kind = ReaderInputKind.Email,
                Location = new ReaderLocation { Path = logicalPath, BlockIndex = blockIndex, SourceBlockIndex = partIndex, HeadingPath = subject + " > Body", SourceBlockKind = bodyKind == "plain" ? "email-body" : "email-body-" + bodyKind, BlockAnchor = $"email-body-m{messageIndex}-{partIndex}" },
                Text = part,
                Markdown = bodyKind == "plain" ? part : "```" + bodyKind + "\n" + part + "\n```",
                Warnings = bodyKind == "plain" ? null : new[] { "EMAIL_" + bodyKind.ToUpperInvariant() + "_BODY_PRESERVED: The source body is retained without executing active content." }
            });
        }
    }

    private static OfficeDocumentReadResult CreateResult(Projection projection, string sourceName, OfficeDocumentSource source) {
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(projection.Chunks, ReaderInputKind.Email, source,
            new[] { OfficeDocumentReaderBuilderEmailExtensions.HandlerId, "officeimo.email", "officeimo.email." + projection.Format.ToString().ToLowerInvariant() }, projection.Assets);
        EmailDocument? primary = projection.Documents.Count == 1 ? projection.Documents[0] : null;
        result.Kind = ReaderInputKind.Email;
        result.Source.Path = sourceName;
        result.Source.Title = primary?.Subject;
        result.Source.Author = primary?.From?.ToString();
        result.Source.Subject = primary?.Subject;
        result.Html = primary?.Body.Html;
        result.Metadata = result.Metadata.Concat(BuildMetadata(projection)).ToArray();
        result.Diagnostics = result.Diagnostics.Concat(projection.Diagnostics.Select(MapDiagnostic)).ToArray();
        return result;
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> BuildMetadata(Projection projection) {
        yield return Metadata("email-format", "Format", projection.Format.ToString(), "string");
        yield return Metadata("email-message-count", "MessageCount", projection.Documents.Count.ToString(CultureInfo.InvariantCulture), "count");
        yield return Metadata("email-attachment-count", "AttachmentCount", projection.Documents.Sum(static document => document.Attachments.Count).ToString(CultureInfo.InvariantCulture), "count");
        for (int index = 0; index < projection.Documents.Count; index++) {
            EmailDocument document = projection.Documents[index];
            string prefix = "email-message-" + index.ToString("D6", CultureInfo.InvariantCulture) + "-";
            string sourceObjectId = "message[" + index.ToString(CultureInfo.InvariantCulture) + "]";
            foreach (OfficeDocumentMetadataEntry entry in MessageMetadata(document, prefix, sourceObjectId)) yield return entry;
            foreach (OfficeDocumentMetadataEntry entry in OutlookItemMetadata(document, prefix, sourceObjectId)) yield return entry;
            EmailMailboxEntry? mailboxEntry = index < projection.MailboxEntries.Count ? projection.MailboxEntries[index] : null;
            if (!string.IsNullOrWhiteSpace(mailboxEntry?.EnvelopeSender))
                yield return Metadata(prefix + "envelope-sender", "EnvelopeSender", mailboxEntry!.EnvelopeSender!, "string", "email.mbox", sourceObjectId);
            if (mailboxEntry?.EnvelopeDate != null)
                yield return Metadata(prefix + "envelope-date", "EnvelopeDate", Date(mailboxEntry.EnvelopeDate)!, "date-time", "email.mbox", sourceObjectId);
        }
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> MessageMetadata(EmailDocument document, string prefix, string sourceObjectId) {
        foreach ((string Name, string? Value, string Type, string Category) item in new[] {
            ("Subject", document.Subject, "string", "email.message"),
            ("OutlookItemKind", document.OutlookItemKind.ToString(), "string", "email.message"),
            ("MessageClass", document.MessageClass, "string", "email.message"),
            ("MessageId", document.MessageId, "string", "email.message"),
            ("From", document.From?.ToString(), "string", "email.address"),
            ("Sender", document.Sender?.ToString(), "string", "email.address"),
            ("To", Join(document, EmailRecipientKind.To), "string", "email.address"),
            ("Cc", Join(document, EmailRecipientKind.Cc), "string", "email.address"),
            ("Bcc", Join(document, EmailRecipientKind.Bcc), "string", "email.address"),
            ("Date", Date(document.Date), "date-time", "email.message"),
            ("ReceivedDate", Date(document.ReceivedDate), "date-time", "email.message"),
            ("AttachmentCount", document.Attachments.Count.ToString(CultureInfo.InvariantCulture), "count", "email.message")
        }) {
            if (!string.IsNullOrWhiteSpace(item.Value)) {
                string id = prefix + ToMetadataId(item.Name);
                yield return Metadata(id, item.Name, item.Value!, item.Type, item.Category, sourceObjectId);
            }
        }
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> OutlookItemMetadata(
        EmailDocument document, string prefix, string sourceObjectId) {
        if (document.Appointment != null) {
            foreach (OfficeDocumentMetadataEntry entry in Entries("email.appointment", sourceObjectId, new[] {
                (prefix + "appointment-start", "Start", Date(document.Appointment.Start), "date-time"),
                (prefix + "appointment-end", "End", Date(document.Appointment.End), "date-time"),
                (prefix + "appointment-location", "Location", document.Appointment.Location, "string"),
                (prefix + "appointment-all-day", "IsAllDay", Bool(document.Appointment.IsAllDay), "boolean"),
                (prefix + "appointment-recurrence", "RecurrencePattern", document.Appointment.RecurrencePattern, "string")
            })) yield return entry;
        }
        if (document.Contact != null) {
            foreach (OfficeDocumentMetadataEntry entry in Entries("email.contact", sourceObjectId, new[] {
                (prefix + "contact-given-name", "GivenName", document.Contact.GivenName, "string"),
                (prefix + "contact-surname", "Surname", document.Contact.Surname, "string"),
                (prefix + "contact-company", "CompanyName", document.Contact.CompanyName, "string"),
                (prefix + "contact-job-title", "JobTitle", document.Contact.JobTitle, "string"),
                (prefix + "contact-email", "Email1Address", document.Contact.Email1.Address, "string")
            })) yield return entry;
        }
        if (document.Task != null) {
            foreach (OfficeDocumentMetadataEntry entry in Entries("email.task", sourceObjectId, new[] {
                (prefix + "task-start", "Start", Date(document.Task.Start), "date-time"),
                (prefix + "task-due", "Due", Date(document.Task.Due), "date-time"),
                (prefix + "task-owner", "Owner", document.Task.Owner, "string"),
                (prefix + "task-complete", "IsComplete", Bool(document.Task.IsComplete), "boolean"),
                (prefix + "task-percent", "PercentComplete", document.Task.PercentComplete?.ToString("0.####", CultureInfo.InvariantCulture), "number")
            })) yield return entry;
        }
        if (document.Journal != null) {
            foreach (OfficeDocumentMetadataEntry entry in Entries("email.journal", sourceObjectId, new[] {
                (prefix + "journal-start", "Start", Date(document.Journal.Start), "date-time"),
                (prefix + "journal-end", "End", Date(document.Journal.End), "date-time"),
                (prefix + "journal-type", "Type", document.Journal.Type, "string")
            })) yield return entry;
        }
        if (document.Note != null) {
            foreach (OfficeDocumentMetadataEntry entry in Entries("email.note", sourceObjectId, new[] {
                (prefix + "note-color", "Color", document.Note.Color?.ToString(CultureInfo.InvariantCulture), "number"),
                (prefix + "note-width", "Width", document.Note.Width?.ToString(CultureInfo.InvariantCulture), "number"),
                (prefix + "note-height", "Height", document.Note.Height?.ToString(CultureInfo.InvariantCulture), "number")
            })) yield return entry;
        }
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> Entries(
        string category, string sourceObjectId,
        IEnumerable<(string Id, string Name, string? Value, string Type)> values) {
        foreach ((string id, string name, string? value, string type) in values) {
            if (!string.IsNullOrWhiteSpace(value)) yield return Metadata(id, name, value!, type, category, sourceObjectId);
        }
    }

    private static string? Bool(bool? value) => value?.ToString().ToLowerInvariant();

    private static string ToMetadataId(string value) {
        var builder = new StringBuilder(value.Length + 4);
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (char.IsUpper(character) && index > 0) builder.Append('-');
            builder.Append(char.ToLowerInvariant(character));
        }
        return builder.ToString();
    }

    private static OfficeDocumentMetadataEntry Metadata(string id, string name, string value, string type, string category = "email.summary", string? sourceObjectId = null) => new OfficeDocumentMetadataEntry { Id = id, Category = category, Name = name, Value = value, ValueType = type, SourceObjectId = sourceObjectId };
    private static OfficeDocumentDiagnostic MapDiagnostic(EmailDiagnostic diagnostic) => new OfficeDocumentDiagnostic { Severity = diagnostic.Severity == EmailDiagnosticSeverity.Error ? OfficeDocumentDiagnosticSeverity.Error : diagnostic.Severity == EmailDiagnosticSeverity.Warning ? OfficeDocumentDiagnosticSeverity.Warning : OfficeDocumentDiagnosticSeverity.Information, Category = OfficeDocumentDiagnosticCategory.Adapter, Code = diagnostic.Code, Message = diagnostic.Message, Source = "OfficeIMO.Reader.Email", IsRecoverable = diagnostic.Severity != EmailDiagnosticSeverity.Error, Location = diagnostic.Location == null ? null : new ReaderLocation { Path = diagnostic.Location } };

    private static string BuildSummary(EmailDocument document, EmailMailboxEntry? entry, string subject) {
        var builder = new StringBuilder().Append("# ").AppendLine(subject);
        Add(builder, "Format", document.Format.ToString()); Add(builder, "Outlook item", document.OutlookItemKind.ToString()); Add(builder, "Message class", document.MessageClass);
        Add(builder, "From", document.From?.ToString()); Add(builder, "Sender", document.Sender?.ToString()); Add(builder, "To", Join(document, EmailRecipientKind.To)); Add(builder, "Cc", Join(document, EmailRecipientKind.Cc)); Add(builder, "Bcc", Join(document, EmailRecipientKind.Bcc));
        Add(builder, "Date", Date(document.Date)); Add(builder, "Received", Date(document.ReceivedDate)); Add(builder, "Message-ID", document.MessageId); Add(builder, "Envelope sender", entry?.EnvelopeSender); Add(builder, "Envelope date", Date(entry?.EnvelopeDate)); Add(builder, "Attachments", document.Attachments.Count.ToString(CultureInfo.InvariantCulture));
        if (document.Appointment != null) { Add(builder, "Start", Date(document.Appointment.Start)); Add(builder, "End", Date(document.Appointment.End)); Add(builder, "Location", document.Appointment.Location); Add(builder, "Recurrence", document.Appointment.RecurrencePattern); }
        if (document.Task != null) { Add(builder, "Task start", Date(document.Task.Start)); Add(builder, "Task due", Date(document.Task.Due)); Add(builder, "Task owner", document.Task.Owner); }
        if (document.Contact != null) { Add(builder, "Given name", document.Contact.GivenName); Add(builder, "Surname", document.Contact.Surname); Add(builder, "Company", document.Contact.CompanyName); Add(builder, "Email", document.Contact.Email1.Address); }
        return builder.ToString().TrimEnd();
    }

    private static string BuildAttachmentSummary(EmailAttachment attachment, string name) { var builder = new StringBuilder().Append("## Attachment: ").AppendLine(name); Add(builder, "Content type", attachment.ContentType); Add(builder, "Length", attachment.Length.ToString(CultureInfo.InvariantCulture)); Add(builder, "Content-ID", attachment.ContentId); Add(builder, "Inline", attachment.IsInline ? "true" : "false"); Add(builder, "Embedded item", attachment.EmbeddedDocument == null ? null : "true"); return builder.ToString().TrimEnd(); }
    private static void Add(StringBuilder builder, string name, string? value) { if (!string.IsNullOrWhiteSpace(value)) builder.Append("- ").Append(name).Append(": ").AppendLine(value!.Replace("\r", " ").Replace("\n", " ").Trim()); }
    private static string Join(EmailDocument document, EmailRecipientKind kind) => string.Join(", ", document.Recipients.Where(recipient => recipient.Kind == kind).Select(recipient => recipient.Address.ToString()).Where(static value => !string.IsNullOrWhiteSpace(value)));
    private static string? Date(DateTimeOffset? value) => value?.ToString("O", CultureInfo.InvariantCulture);
    private static string Limit(string value, int maxChars, List<string> warnings, string label) { int limit = Math.Max(256, maxChars); if (value.Length <= limit) return value; warnings.Add(label + " was truncated due to ReaderOptions.MaxChars."); return value.Substring(0, limit); }
    private static IReadOnlyList<string>? DiagnosticsToWarnings(IReadOnlyList<EmailDiagnostic> diagnostics) { string[] values = diagnostics.Where(static item => item.Severity != EmailDiagnosticSeverity.Information).Select(static item => item.Code + ": " + item.Message).ToArray(); return values.Length == 0 ? null : values; }
    private static void EnrichChunks(IEnumerable<ReaderChunk> chunks, OfficeDocumentSource source, bool computeHashes) { foreach (ReaderChunk chunk in chunks) { chunk.SourceId = source.SourceId; chunk.SourceHash = source.SourceHash; chunk.SourceLengthBytes = source.LengthBytes; chunk.SourceLastWriteUtc = source.LastWriteUtc; chunk.TokenEstimate = chunk.Text.Length == 0 ? 0 : Math.Max(1, (chunk.Text.Length + 3) / 4); if (computeHashes) chunk.ChunkHash = Hash(chunk.Text + "\n" + chunk.Markdown); } }
    private static string NormalizeSourceKey(string value) => Path.DirectorySeparatorChar == '\\' ? value.ToLowerInvariant() : value;
    private static string? TryExtension(string name) { try { string extension = Path.GetExtension(name); return string.IsNullOrWhiteSpace(extension) ? null : extension; } catch { return null; } }
    private static long? TryLength(string path) { try { return File.Exists(path) ? new FileInfo(path).Length : (long?)null; } catch { return null; } }
    private static DateTime? TryLastWrite(string path) { try { return File.Exists(path) ? new FileInfo(path).LastWriteTimeUtc : (DateTime?)null; } catch { return null; } }
    private static long? TryLength(Stream stream) { try { return stream.CanSeek ? stream.Length : (long?)null; } catch { return null; } }
    private static string? TryHashFile(string path) { try { using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete); return Hash(stream); } catch { return null; } }
    private static string? TryHashStream(Stream stream) { if (!stream.CanSeek) return null; long position = stream.Position; try { stream.Position = 0; return Hash(stream); } catch { return null; } finally { stream.Position = position; } }
    private static string Hash(string? value) => Hash(Encoding.UTF8.GetBytes(value ?? string.Empty));
    private static string Hash(byte[] bytes) { using SHA256 algorithm = SHA256.Create(); return ToHex(algorithm.ComputeHash(bytes)); }
    private static string Hash(Stream stream) { using SHA256 algorithm = SHA256.Create(); return ToHex(algorithm.ComputeHash(stream)); }
    private static string ToHex(byte[] bytes) { var builder = new StringBuilder(bytes.Length * 2); foreach (byte value in bytes) builder.Append(value.ToString("x2", CultureInfo.InvariantCulture)); return builder.ToString(); }

    private sealed class Projection {
        internal Projection(string sourceName, EmailFileFormat format) { SourceName = sourceName; Format = format; }
        internal string SourceName { get; }
        internal EmailFileFormat Format { get; }
        internal List<EmailDocument> Documents { get; } = new List<EmailDocument>();
        internal List<EmailMailboxEntry?> MailboxEntries { get; } = new List<EmailMailboxEntry?>();
        internal List<EmailDiagnostic> Diagnostics { get; } = new List<EmailDiagnostic>();
        internal List<ReaderChunk> Chunks { get; } = new List<ReaderChunk>();
        internal List<OfficeDocumentAsset> Assets { get; } = new List<OfficeDocumentAsset>();
    }
}

internal sealed class EmailDocumentProjectionCursor {
    internal int NextMessageIndex { get; set; }
    internal int NextBlockIndex { get; set; }
    internal bool DiagnosticsAttached { get; set; }
}
