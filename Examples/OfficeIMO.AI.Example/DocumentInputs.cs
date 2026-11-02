using OfficeIMO.AI;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Image;
using OfficeIMO.Reader.Pdf;

internal static class DocumentInputs {
    public static async Task<byte[]> ReadFileAsync(string path, int maxBytes, CancellationToken cancellationToken) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length < 1 || stream.Length > maxBytes) throw new InvalidDataException("File exceeds the input budget or is empty.");
        byte[] bytes = new byte[checked((int)stream.Length)];
        await stream.ReadExactlyAsync(bytes, cancellationToken);
        if (stream.ReadByte() != -1) throw new InvalidDataException("Source grew while being captured.");
        return bytes;
    }

    public static async Task<OfficeAiDocument> ReadAsync(byte[] bytes, string name, bool includeImages,
        IReadOnlyList<int> pages, OfficeAiLimits limits, CancellationToken cancellationToken) {
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddPdfHandler().AddImageHandler().Build();
        OfficeDocumentReadResult result = await reader.ReadDocumentAsync(bytes, name, new ReaderOptions {
            MaxInputBytes = limits.MaxInputBytes, MaxChars = Math.Min(limits.MaxDocumentCharacters, 8000),
            MaxTableRows = Math.Min(limits.MaxTableCells, 10_000)
        }, cancellationToken);
        if (result.Diagnostics.Any(item => item.Severity == OfficeDocumentDiagnosticSeverity.Error))
            throw new InvalidDataException("The reader reported a blocking source diagnostic.");
        var images = new List<OfficeAiImage>();
        bool pdf = result.Kind == ReaderInputKind.Pdf;
        if (includeImages && pdf) {
            using PdfDocument document = PdfDocument.Load(new MemoryStream(bytes), new PdfLoadOptions {
                Limits = new PdfReadLimits { MaxInputBytes = limits.MaxInputBytes }, PermissionPolicy = PdfPermissionPolicy.Enforce
            });
            PdfPageSelector? selection = pages.Count == 0 ? null : PdfPageSelector.Parse(string.Join(",", pages));
            IReadOnlyList<OfficeImageExportResult> rendered = document.ExportImages(
                new PdfImageExportOptions {
                    Format = OfficeImageExportFormat.Png, PageSelector = selection,
                    TargetDpi = 120, MaximumOutputCount = limits.MaxPages, MaximumRasterPixels = limits.MaxImagePixels,
                    MaximumTotalRasterPixels = limits.MaxImagePixels * Math.Min(limits.MaxRequests, limits.MaxPages),
                    MaximumTotalEncodedBytes = limits.MaxInputBytes, RenderTimeout = limits.Timeout
                }, out OfficeIMO.OfficeConversionReport renderReport, cancellationToken);
            for (int index = 0; index < rendered.Count; index++) {
                byte[] page = rendered[index].Bytes;
                if (!OfficeImageReader.TryValidateContent(page, ".png", out OfficeImageInfo info))
                    throw new InvalidDataException("The exported PDF page has no valid raster metadata.");
                int number = pages.Count == 0 ? index + 1 : pages[index];
                images.Add(new("image-page-" + number, number, info.MimeType, page, info.Width, info.Height));
            }
            result.Diagnostics = result.Diagnostics.Concat(renderReport.Diagnostics.Select(diagnostic => new OfficeDocumentDiagnostic {
                Code = diagnostic.Code, Message = diagnostic.Message, Source = diagnostic.Format,
                Severity = diagnostic.Severity switch {
                    OfficeIMO.OfficeDiagnosticSeverity.Error => OfficeDocumentDiagnosticSeverity.Error,
                    OfficeIMO.OfficeDiagnosticSeverity.Warning => OfficeDocumentDiagnosticSeverity.Warning,
                    _ => OfficeDocumentDiagnosticSeverity.Information
                },
                Attributes = new Dictionary<string, string> { ["location"] = diagnostic.Location ?? string.Empty }
            })).ToArray();
        } else if (result.Kind == ReaderInputKind.Unknown
            && result.CapabilitiesUsed.Contains(OfficeDocumentReaderBuilderImageExtensions.HandlerId, StringComparer.Ordinal)
            && result.Assets.Any(asset => asset.PayloadBytes is not null && asset.MediaType?.StartsWith("image/", StringComparison.Ordinal) == true)) {
            // Standalone image metadata is not recognized page text.
            result.Blocks = Array.Empty<OfficeDocumentBlock>(); result.Chunks = Array.Empty<ReaderChunk>();
            if (includeImages) {
                OfficeDocumentAsset asset = result.Assets.First(item => item.PayloadBytes is not null);
                images.Add(new("image-page-1", 1, asset.MediaType!, asset.PayloadBytes!, asset.Width!.Value, asset.Height!.Value));
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeAiDocument.FromReadResult(bytes, result, images, limits);
    }
}
