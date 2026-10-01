using OfficeIMO.Email;

namespace OfficeIMO.Reader.Email;

internal static partial class EmailReaderProjection {
    private static bool IsMailAttachment(EmailAttachment attachment, string name) =>
        attachment.ContentType?.Equals("message/rfc822", StringComparison.OrdinalIgnoreCase) == true ||
        attachment.ContentType?.Equals("message/global", StringComparison.OrdinalIgnoreCase) == true ||
        attachment.ContentType?.Equals("application/vnd.ms-outlook", StringComparison.OrdinalIgnoreCase) == true ||
        attachment.ContentType?.Equals("application/ms-tnef", StringComparison.OrdinalIgnoreCase) == true ||
        attachment.ContentType?.Equals("application/mbox", StringComparison.OrdinalIgnoreCase) == true ||
        TryExtension(name)?.ToLowerInvariant() is ".eml" or ".msg" or ".oft" or ".tnef" or ".emlx" or
            ".mbox" or ".mbx" or ".pst" or ".ost" or ".olm";

    private static void AddAttachmentContent(
        EmailAttachment attachment, string fileName, string attachmentPath, string subject,
        ReaderChunk attachmentChunk, Projection projection, ReaderOptions options,
        EmailDocumentProjectionCursor cursor, CancellationToken cancellationToken) {
        string sourceName = ResolveAttachmentSourceName(fileName, attachment.ContentType);
        bool hasHandler = ReaderNestedContent.CanRead(sourceName);
        bool plainText = IsPlainTextAttachment(sourceName, attachment.ContentType);
        if (!hasHandler && !plainText) {
            projection.AttachmentSkipped++;
            AttachmentOutcome(projection, attachmentChunk, attachmentPath, "EMAIL_ATTACHMENT_READER_UNSUPPORTED",
                "No registered Reader handler or plain-text fallback supports this attachment.");
            return;
        }
        if ((attachment.Content == null || attachment.Content.Length == 0) && attachment.ContentSource == null) {
            projection.AttachmentSkipped++;
            AttachmentOutcome(projection, attachmentChunk, attachmentPath, "EMAIL_ATTACHMENT_READER_CONTENT_UNAVAILABLE",
                "The attachment has no available payload to extract.");
            return;
        }

        projection.AttachmentAttempted++;
        try {
            bool rtf = string.Equals(TryExtension(sourceName), ".rtf", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(attachment.ContentType, "text/rtf", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(attachment.ContentType, "application/rtf", StringComparison.OrdinalIgnoreCase);
            bool text = !rtf && (plainText || attachment.ContentType?.StartsWith("text/", StringComparison.OrdinalIgnoreCase) == true);
            EmailAttachmentTextResult? decoded = text
                ? EmailAttachmentTextReader.Read(attachment, options.MaxInputBytes ?? 16L * 1024L * 1024L, cancellationToken)
                : null;
            if (decoded != null) {
                foreach (EmailDiagnostic diagnostic in decoded.Diagnostics) {
                    projection.Diagnostics.Add(new EmailDiagnostic(diagnostic.Code, diagnostic.Message, diagnostic.Severity, attachmentPath));
                }
                string[] warnings = decoded.Diagnostics.Select(item => item.Code + ": " + item.Message).ToArray();
                if (warnings.Length > 0)
                    attachmentChunk.Warnings = (attachmentChunk.Warnings ?? Array.Empty<string>()).Concat(warnings).ToArray();
            }
            using Stream stream = decoded != null
                ? new MemoryStream(EncodeAttachmentProjection(decoded.Text, sourceName), writable: false)
                : attachment.OpenContentStream();
            IReadOnlyList<ReaderChunk> nested = hasHandler
                ? ReaderNestedContent.Read(stream, sourceName, CloneWithoutHashes(options), cancellationToken)
                : BuildPlainAttachmentChunks(decoded!.Text, options.MaxChars);

            for (int index = 0; index < nested.Count; index++) {
                ReaderChunk child = nested[index];
                int blockIndex = cursor.NextBlockIndex++;
                child.Id = $"email:attachment-content:{blockIndex.ToString("D6", CultureInfo.InvariantCulture)}:{index.ToString("D4", CultureInfo.InvariantCulture)}";
                child.Location = CloneNestedLocation(child.Location, attachmentPath,
                    subject + " > Attachment: " + fileName, blockIndex, child.Location.SourceBlockKind);
                ClearNestedSource(child);
                projection.Chunks.Add(child);
            }
            if (!nested.Any(chunk => !string.IsNullOrWhiteSpace(chunk.Text) || !string.IsNullOrWhiteSpace(chunk.Markdown))) {
                projection.AttachmentEmpty++;
                AttachmentOutcome(projection, attachmentChunk, attachmentPath, "EMAIL_ATTACHMENT_READER_EMPTY",
                    "Extraction completed without readable text.");
            } else {
                projection.AttachmentSucceeded++;
                AttachmentOutcome(projection, attachmentChunk, attachmentPath, "EMAIL_ATTACHMENT_READER_SUCCEEDED",
                    "Extraction produced readable attachment content.", includeChunkWarning: false);
            }
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception exception) {
            projection.AttachmentFailed++;
            AttachmentOutcome(projection, attachmentChunk, attachmentPath, "EMAIL_ATTACHMENT_READER_FAILED",
                exception.GetType().Name + " while extracting the attachment.", EmailDiagnosticSeverity.Warning);
        }
    }

    private static void AttachmentOutcome(Projection projection, ReaderChunk chunk, string path,
        string code, string message, EmailDiagnosticSeverity severity = EmailDiagnosticSeverity.Information,
        bool includeChunkWarning = true) {
        projection.Diagnostics.Add(new EmailDiagnostic(code, message, severity, path));
        if (includeChunkWarning)
            chunk.Warnings = (chunk.Warnings ?? Array.Empty<string>()).Concat(new[] { code + ": " + message }).ToArray();
    }
}
