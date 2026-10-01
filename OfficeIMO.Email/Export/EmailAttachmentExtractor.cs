using OfficeIMO.Core.Internal;
using OfficeIMO.Email.Store;
using OfficeIMO.Internal;
using System.Security.Cryptography;

namespace OfficeIMO.Email;

/// <summary>Exports attachment bytes to a caller-controlled directory with portable names and atomic commits.</summary>
public static class EmailAttachmentExtractor {
    /// <summary>
    /// Extracts attachments without overwriting existing files, opening linked paths, or modifying the model.
    /// Keep the destination under the caller's control while extraction runs. The returned manifest contains
    /// source names and paths; treat it as mail data. Earlier completed files survive cancellation or later failures.
    /// Source operations run on a worker so asynchronous content providers cannot capture a blocked UI context.
    /// </summary>
    public static EmailAttachmentExtractionResult Extract(EmailDocument document, string destinationDirectory,
        EmailAttachmentExtractionOptions? options = null, CancellationToken cancellationToken = default) =>
        Task.Run(() => ExtractAsync(document, destinationDirectory, options, cancellationToken), cancellationToken)
            .GetAwaiter().GetResult();

    /// <summary>Asynchronously stages source content and exports attachments under the same bounds as <see cref="Extract"/>.</summary>
    public static async Task<EmailAttachmentExtractionResult> ExtractAsync(EmailDocument document, string destinationDirectory,
        EmailAttachmentExtractionOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (destinationDirectory == null) throw new ArgumentNullException(nameof(destinationDirectory));
        cancellationToken.ThrowIfCancellationRequested();
        var effective = options ?? new EmailAttachmentExtractionOptions();
        Directory.CreateDirectory(destinationDirectory);
        using var directoryHandle = OfficePathIdentity.OpenDirectoryForIdentity(destinationDirectory, out string root);
        var entries = new List<EmailAttachmentExtractionEntry>();
        var diagnostics = new List<EmailDiagnostic>();
        var ancestors = new HashSet<EmailDocument>();
        long consumed = 0;
        int visitedAttachments = 0;
        bool truncated = false;
        await Visit(document, string.Empty, 0).ConfigureAwait(false);
        return new EmailAttachmentExtractionResult(entries.AsReadOnly(), consumed, truncated, diagnostics.AsReadOnly());

        async Task Visit(EmailDocument current, string prefix, int depth) {
            if (!ancestors.Add(current)) {
                truncated = true;
                diagnostics.Add(new EmailDiagnostic("EMAIL_EXTRACTION_CYCLE", "An embedded-message cycle was not traversed.", location: prefix));
                return;
            }
            try {
                for (int index = 0; index < current.Attachments.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (visitedAttachments >= effective.MaxAttachments || consumed >= effective.MaxTotalBytes) {
                        truncated = true;
                        diagnostics.Add(new EmailDiagnostic("EMAIL_EXTRACTION_BUDGET", "Attachment count or total byte budget stopped traversal.", location: prefix));
                        return;
                    }
                    EmailAttachment attachment = current.Attachments[index];
                    visitedAttachments++;
                    string sourcePath = prefix + index.ToString(CultureInfo.InvariantCulture);
                    var itemDiagnostics = new List<EmailDiagnostic>();
                    string? outputPath = null;
                    string? digest = null;
                    long written = 0;
                    bool included = (effective.IncludeInline || !attachment.IsInline) && (effective.IncludeHidden || !attachment.IsHidden);
                    try {
                        if (!included) {
                            itemDiagnostics.Add(new EmailDiagnostic("EMAIL_EXTRACTION_SKIPPED", "Attachment excluded by inline or hidden policy.", location: sourcePath));
                        } else if (attachment.EmbeddedDocument == null && attachment.Content == null && attachment.ContentSource == null) {
                            itemDiagnostics.Add(new EmailDiagnostic("EMAIL_EXTRACTION_CONTENT_UNAVAILABLE", "Decoded content is unavailable; linked attachment paths are never opened.", location: sourcePath));
                        } else {
                            long maximum = Math.Min(effective.MaxAttachmentBytes, effective.MaxTotalBytes - consumed);
                            string name = EmailStoreExportPathBuilder.SanitizeSegmentByUtf8Bytes(attachment.FileName, 160, 180, "attachment");
                            if (attachment.EmbeddedDocument != null) name = Path.GetFileNameWithoutExtension(name) + ".eml";
                            string path = Path.Combine(root, entries.Count.ToString("D4", CultureInfo.InvariantCulture) + "-" +
                                EmailStoreExportPathBuilder.GetStableHash(sourcePath) + "-" + name);
                            OfficePathIdentity.EnsurePathMatchesOpenedDirectory(root, directoryHandle);
                            await OfficeFileCommit.WriteAsync(path, async (output, token) => {
                                using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
                                if (attachment.EmbeddedDocument != null) {
                                    int remainingDepth = effective.MaxDepth - depth - 1;
                                    if (remainingDepth < 0 && attachment.EmbeddedDocument.Attachments.Count > 0)
                                        throw new EmailLimitExceededException(nameof(effective.MaxDepth), depth + 1, effective.MaxDepth);
                                    remainingDepth = Math.Max(0, remainingDepth);
                                    int stagingLimit = effective.MaxAttachments - visitedAttachments;
                                    ReserveEmbeddedVisits(attachment.EmbeddedDocument, depth + 1, effective.MaxDepth, () => {
                                        cancellationToken.ThrowIfCancellationRequested();
                                        if (visitedAttachments >= effective.MaxAttachments)
                                            throw new EmailLimitExceededException(nameof(effective.MaxAttachments), visitedAttachments + 1, effective.MaxAttachments);
                                        visitedAttachments++;
                                    });
                                    var preflightWriter = new EmailDocumentWriter(new EmailWriterOptions(usePreservedRawSource: true,
                                        maxNestedMessageDepth: remainingDepth, maxOutputBytes: maximum));
                                    EmailConversionReport conversion = preflightWriter.AnalyzeConversion(attachment.EmbeddedDocument);
                                    if (!conversion.CanWrite) {
                                        itemDiagnostics.AddRange(conversion.Diagnostics);
                                        throw new InvalidDataException("Embedded message serialization was blocked; see its diagnostics.");
                                    }
                                    using var staging = await EmailAttachmentStaging.CreateAsync(attachment.EmbeddedDocument, maximum,
                                        cancellationToken, count => consumed += count, remainingDepth, stagingLimit).ConfigureAwait(false);
                                    using var scope = staging.EnterScope();
                                    long outputMaximum = Math.Min(effective.MaxAttachmentBytes, effective.MaxTotalBytes - consumed);
                                    if (outputMaximum <= 0) throw new EmailLimitExceededException(nameof(effective.MaxTotalBytes), consumed + 1, effective.MaxTotalBytes);
                                    var writer = new EmailDocumentWriter(new EmailWriterOptions(usePreservedRawSource: true,
                                        maxNestedMessageDepth: remainingDepth, maxOutputBytes: outputMaximum));
                                    using var observed = new EmailExtractionWriteStream(output, cancellationToken, count => {
                                        consumed += count; written += count;
                                    }, hash);
                                    EmailWriteResult result = writer.Write(attachment.EmbeddedDocument, observed);
                                    itemDiagnostics.AddRange(result.Diagnostics);
                                    if (result.HasErrors) throw new InvalidDataException("Embedded message serialization was blocked; see its diagnostics.");
                                    cancellationToken.ThrowIfCancellationRequested();
                                } else {
                                    using Stream input = await attachment.OpenContentStreamAsync(cancellationToken).ConfigureAwait(false);
                                    var buffer = new byte[81920];
                                    while (true) {
                                        cancellationToken.ThrowIfCancellationRequested();
                                        long remaining = maximum - written;
                                        int read = await input.ReadAsync(buffer, 0, remaining >= buffer.Length ? buffer.Length : (int)remaining + 1, cancellationToken).ConfigureAwait(false);
                                        if (read == 0) break;
                                        consumed += read;
                                        cancellationToken.ThrowIfCancellationRequested();
                                        if (read > remaining) throw new EmailLimitExceededException(nameof(effective.MaxAttachmentBytes), written + read, maximum);
                                        await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
                                        hash.AppendData(buffer, 0, read);
                                        written += read;
                                    }
                                }
                                cancellationToken.ThrowIfCancellationRequested();
                                OfficePathIdentity.EnsurePathMatchesOpenedDirectory(root, directoryHandle);
                                digest = BitConverter.ToString(hash.GetHashAndReset()).Replace("-", string.Empty).ToLowerInvariant();
                            }, OfficeFileCommit.ConflictPolicy.FailIfExists, cancellationToken).ConfigureAwait(false);
                            outputPath = path;
                        }
                    } catch (Exception exception) when (exception is IOException || exception is InvalidDataException || exception is UnauthorizedAccessException || exception is InvalidOperationException) {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (exception is EmailLimitExceededException) truncated = true;
                        itemDiagnostics.Add(exception is EmailLimitExceededException limit
                            ? EmailDiagnostic.FromLimit(limit, "Attachment extraction", sourcePath)
                            : new EmailDiagnostic("EMAIL_EXTRACTION_FAILED", exception.Message, EmailDiagnosticSeverity.Error, sourcePath));
                    }
                    entries.Add(new EmailAttachmentExtractionEntry(sourcePath, attachment.FileName, outputPath,
                        outputPath == null ? 0 : written, outputPath == null ? null : digest, itemDiagnostics.AsReadOnly()));
                    if (included && effective.RecurseEmbeddedMessages && attachment.EmbeddedDocument != null) {
                        if (depth >= effective.MaxDepth) {
                            truncated = true;
                            diagnostics.Add(new EmailDiagnostic("EMAIL_EXTRACTION_DEPTH", "Nested attachment traversal reached its depth limit.", location: sourcePath));
                        } else await Visit(attachment.EmbeddedDocument, sourcePath + "/", depth + 1).ConfigureAwait(false);
                    }
                }
            } finally { ancestors.Remove(current); }
        }
    }

    private static void ReserveEmbeddedVisits(EmailDocument document, int depth, int maxDepth, Action visit) {
        var active = new HashSet<EmailDocument>();
        Reserve(document, depth);
        void Reserve(EmailDocument current, int currentDepth) {
            if (!active.Add(current)) throw new InvalidDataException("An embedded-message cycle cannot be exported.");
            try {
                foreach (EmailAttachment attachment in current.Attachments) {
                    if (currentDepth > maxDepth)
                        throw new EmailLimitExceededException("MaxDepth", currentDepth, maxDepth);
                    visit();
                    if (attachment.EmbeddedDocument != null) Reserve(attachment.EmbeddedDocument, currentDepth + 1);
                }
            } finally { active.Remove(current); }
        }
    }
}
