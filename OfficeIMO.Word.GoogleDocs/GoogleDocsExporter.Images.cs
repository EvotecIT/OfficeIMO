using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using System.IO;

namespace OfficeIMO.Word.GoogleDocs {
    public sealed partial class GoogleDocsExporter : IGoogleDocsExporter {
        private static async Task<IReadOnlyDictionary<GoogleDocsInlineImage, string>> UploadInlineImagesAsync(
            GoogleDriveClient driveClient,
            GoogleDocsBatch batch,
            IList<GoogleDriveTemporaryContentLease> leases,
            CancellationToken cancellationToken) {
            var imageUris = new Dictionary<GoogleDocsInlineImage, string>();
            foreach (var image in EnumerateInlineImages(batch)) {
                if (!TryResolveImageUploadPayload(image, out var uploadName, out var mimeType, out var bytes, out var diagnosticMessage)) {
                    batch.Report.Add(
                        TranslationSeverity.Warning,
                        "InlineImages",
                        diagnosticMessage);
                    continue;
                }

                var lease = await GoogleDriveTemporaryContentLease.CreatePublicReadLeaseAsync(
                    driveClient,
                    bytes,
                    new GoogleDriveUploadOptions {
                        Name = uploadName,
                        ContentType = mimeType,
                    },
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);
                leases.Add(lease);
                imageUris[image] = lease.PublicUri;
            }

            return imageUris;
        }

        private static async Task CleanupTemporaryInlineImagesAsync(
            IReadOnlyList<GoogleDriveTemporaryContentLease> leases,
            TranslationReport report,
            CancellationToken cancellationToken) {
            if (leases.Count == 0) {
                return;
            }

            int deletedCount = 0;
            foreach (var lease in leases.Reverse()) {
                var cleanup = await lease.CleanupAsync(cancellationToken).ConfigureAwait(false);
                deletedCount += cleanup.Entries.Count(entry => entry.Status == GoogleDriveCleanupStatus.Deleted);
            }

            if (deletedCount > 0) {
                report.Add(
                    TranslationSeverity.Info,
                    "InlineImages",
                    $"Deleted {deletedCount} temporary public Drive staging file(s) after Google Docs fetched the inline images.",
                    code: "DOCS.IMAGE.TEMPORARY_LEASE_CLEANED",
                    action: TranslationAction.Preserve,
                    count: deletedCount);
            }
        }


        private static IEnumerable<GoogleDocsInlineImage> EnumerateInlineImages(GoogleDocsBatch batch) {
            var seen = new HashSet<GoogleDocsInlineImage>();
            foreach (var paragraphRequest in batch.Requests.OfType<GoogleDocsInsertParagraphRequest>()) {
                foreach (var image in EnumerateParagraphImages(paragraphRequest.Paragraph)) {
                    if (seen.Add(image)) {
                        yield return image;
                    }
                }
            }

            foreach (var tableRequest in batch.Requests.OfType<GoogleDocsInsertTableRequest>()) {
                foreach (var image in EnumerateTableImages(tableRequest.Table)) {
                    if (seen.Add(image)) {
                        yield return image;
                    }
                }
            }

            foreach (var segment in batch.Segments) {
                foreach (var request in segment.Requests) {
                    switch (request) {
                        case GoogleDocsInsertParagraphRequest paragraphRequest:
                            foreach (var image in EnumerateParagraphImages(paragraphRequest.Paragraph)) {
                                if (seen.Add(image)) {
                                    yield return image;
                                }
                            }
                            break;
                        case GoogleDocsInsertTableRequest tableRequest:
                            foreach (var image in EnumerateTableImages(tableRequest.Table)) {
                                if (seen.Add(image)) {
                                    yield return image;
                                }
                            }
                            break;
                    }
                }
            }
        }

        private static IEnumerable<GoogleDocsInlineImage> EnumerateParagraphImages(GoogleDocsParagraph paragraph) {
            foreach (var run in paragraph.Runs) {
                if (run.InlineImage != null) {
                    yield return run.InlineImage;
                }

                if (run.Footnote == null) {
                    continue;
                }

                foreach (var footnoteParagraph in run.Footnote.Paragraphs) {
                    foreach (var image in EnumerateParagraphImages(footnoteParagraph)) {
                        yield return image;
                    }
                }
            }
        }

        private static IEnumerable<GoogleDocsInlineImage> EnumerateTableImages(GoogleDocsTable table) {
            foreach (var row in table.Rows) {
                foreach (var cell in row.Cells) {
                    foreach (var paragraph in cell.Paragraphs) {
                        foreach (var image in EnumerateParagraphImages(paragraph)) {
                            yield return image;
                        }
                    }
                }
            }
        }


        private static bool TryResolveImageUploadPayload(
            GoogleDocsInlineImage image,
            out string fileName,
            out string mimeType,
            out byte[] bytes,
            out string diagnosticMessage) {
            fileName = image.FileName ?? string.Empty;
            mimeType = image.ContentType ?? string.Empty;
            bytes = Array.Empty<byte>();
            diagnosticMessage = string.Empty;

            if (image.Bytes != null && image.Bytes.Length > 0) {
                bytes = image.Bytes;
                if (string.IsNullOrWhiteSpace(fileName)) {
                    fileName = string.IsNullOrWhiteSpace(image.FilePath) ? "officeimo-inline-image" : Path.GetFileName(image.FilePath);
                }

                if (string.IsNullOrWhiteSpace(mimeType) && !TryGetImageMimeType(fileName, out mimeType)) {
                    diagnosticMessage = "An inline image was embedded in the Word document, but its content type could not be inferred for the current Google Docs upload slice.";
                    return false;
                }

                return true;
            }

            if (!string.IsNullOrWhiteSpace(image.FilePath) && File.Exists(image.FilePath)) {
                var existingFilePath = image.FilePath!;
                fileName = string.IsNullOrWhiteSpace(fileName) ? Path.GetFileName(existingFilePath) : fileName;
                if (string.IsNullOrWhiteSpace(mimeType) && !TryGetImageMimeType(existingFilePath, out mimeType)) {
                    diagnosticMessage = $"Inline image '{existingFilePath}' uses an unsupported extension for the current Google Docs image upload slice, so export kept the readable placeholder.";
                    return false;
                }

                bytes = File.ReadAllBytes(existingFilePath);
                return true;
            }

            if (!string.IsNullOrWhiteSpace(image.FilePath)) {
                diagnosticMessage = $"Inline image file '{image.FilePath}' was not found, so Google Docs export kept the readable placeholder instead of a native image.";
                return false;
            }

            diagnosticMessage = "A Word inline image did not expose embedded bytes or a local file path, so Google Docs export kept the readable placeholder instead of a native image.";
            return false;
        }

        private static bool TryGetImageMimeType(string filePath, out string mimeType) {
            switch (Path.GetExtension(filePath).ToLowerInvariant()) {
                case ".png":
                    mimeType = "image/png";
                    return true;
                case ".jpg":
                case ".jpeg":
                    mimeType = "image/jpeg";
                    return true;
                case ".gif":
                    mimeType = "image/gif";
                    return true;
                case ".bmp":
                    mimeType = "image/bmp";
                    return true;
                default:
                    mimeType = string.Empty;
                    return false;
            }
        }

    }
}
