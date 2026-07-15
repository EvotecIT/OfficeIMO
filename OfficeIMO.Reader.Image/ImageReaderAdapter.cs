using OfficeIMO.Drawing;
using System.Text;

namespace OfficeIMO.Reader.Image;

internal static class ImageReaderAdapter {
    internal static readonly string[] Extensions = {
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".svg",
        ".emf", ".wmf", ".ico", ".pcx", ".webp"
    };

    internal static OfficeDocumentReadResult ReadDocument(
        string path,
        ReaderOptions readerOptions,
        ReaderImageOptions imageOptions,
        CancellationToken cancellationToken) {
        ReaderAdapterInputSnapshot input = DocumentReaderEngine.ReadAdapterInput(path, readerOptions, cancellationToken);
        return BuildDocument(input, readerOptions, imageOptions, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(
        Stream stream,
        string? sourceName,
        ReaderOptions readerOptions,
        ReaderImageOptions imageOptions,
        CancellationToken cancellationToken) {
        ReaderAdapterInputSnapshot input = DocumentReaderEngine.ReadAdapterInput(
            stream,
            string.IsNullOrWhiteSpace(sourceName) ? "image.bin" : sourceName,
            readerOptions,
            cancellationToken);
        return BuildDocument(input, readerOptions, imageOptions, cancellationToken);
    }

    private static OfficeDocumentReadResult BuildDocument(
        ReaderAdapterInputSnapshot input,
        ReaderOptions readerOptions,
        ReaderImageOptions imageOptions,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string sourceName = input.Source.Path ?? "image.bin";
        if (!OfficeImageReader.TryIdentify(input.Bytes, fileName: null, out OfficeImageInfo info) ||
            info.Format == OfficeImageFormat.Unknown) {
            throw new NotSupportedException("Image format is not supported: " + sourceName);
        }

        string assetId = "image-0000";
        string extension = OfficeImageInfo.GetDefaultExtension(info.Format);
        string payloadHash = OfficeDocumentAssetHash.ComputeSha256Hex(input.Bytes);
        var location = new ReaderLocation {
            Path = sourceName,
            BlockIndex = 0,
            SourceBlockIndex = 0,
            SourceBlockKind = "image",
            BlockAnchor = assetId
        };
        var asset = new OfficeDocumentAsset {
            Id = assetId,
            Kind = "image",
            MediaType = info.MimeType,
            Extension = extension,
            FileName = OfficeDocumentAssetNaming.BuildFileName(assetId, extension),
            Title = Path.GetFileName(sourceName),
            Width = info.Width > 0 ? info.Width : null,
            Height = info.Height > 0 ? info.Height : null,
            LengthBytes = input.Bytes.LongLength,
            PayloadHash = payloadHash,
            PayloadBytes = imageOptions.IncludePayload ? input.Bytes : null,
            SourceObjectId = "source-image",
            Location = location
        };
        string markdown = BuildMarkdown(
            sourceName,
            info,
            input.Bytes.LongLength,
            assetId,
            imageOptions.IncludePayload,
            imageOptions.CreateOcrCandidate);
        var chunk = new ReaderChunk {
            Id = "image-metadata-0000",
            Kind = ReaderInputKind.Unknown,
            Location = location,
            Text = BuildPlainText(sourceName, info, input.Bytes.LongLength),
            Markdown = markdown,
            Visuals = new[] {
                new ReaderVisual {
                    Kind = "image",
                    Language = info.MimeType,
                    Content = Path.GetFileName(sourceName),
                    PayloadHash = payloadHash,
                    Location = location
                }
            }
        };
        DocumentReaderEngine.ApplyAdapterSource(chunk, input, readerOptions.ComputeHashes);
        IReadOnlyList<OfficeDocumentOcrCandidate> ocrCandidates = imageOptions.CreateOcrCandidate
            ? new[] {
                new OfficeDocumentOcrCandidate {
                    Id = "image-ocr-0000",
                    Kind = "image",
                    Reason = "Standalone image is available for optional OCR; no OCR engine was run.",
                    Confidence = 1D,
                    AssetId = assetId,
                    ImageCount = 1,
                    TextBlockCount = 0,
                    Location = location
                }
            }
            : Array.Empty<OfficeDocumentOcrCandidate>();
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            new[] { chunk },
            ReaderInputKind.Unknown,
            input.Source,
            new[] { OfficeDocumentReaderBuilderImageExtensions.HandlerId, "officeimo.drawing.image-identification" },
            new[] { asset },
            ocrCandidates);
        result.Source.Title = Path.GetFileName(sourceName);
        result.Metadata = result.Metadata.Concat(BuildMetadata(info, input.Bytes.LongLength)).ToArray();
        result.Visuals = chunk.Visuals ?? Array.Empty<ReaderVisual>();
        return result;
    }

    private static string BuildMarkdown(
        string sourceName,
        OfficeImageInfo info,
        long length,
        string assetId,
        bool includePayload,
        bool ocrCandidate) {
        var markdown = new StringBuilder();
        markdown.Append("# ").AppendLine(Path.GetFileName(sourceName));
        markdown.AppendLine();
        markdown.Append("- Format: ").AppendLine(info.Format.ToString());
        markdown.Append("- Media type: ").AppendLine(info.MimeType);
        if (info.Width > 0 && info.Height > 0) {
            markdown.Append("- Dimensions: ")
                .Append(info.Width.ToString(CultureInfo.InvariantCulture))
                .Append(" × ")
                .Append(info.Height.ToString(CultureInfo.InvariantCulture))
                .AppendLine(" px");
        }
        markdown.Append("- Size: ").Append(length.ToString(CultureInfo.InvariantCulture)).AppendLine(" bytes");
        markdown.AppendLine();
        if (includePayload) {
            markdown.Append("Image bytes are available as asset `").Append(assetId).Append('`').AppendLine();
        } else {
            markdown.AppendLine("Image metadata is available; source bytes were not retained as an asset payload.");
        }
        if (ocrCandidate) markdown.AppendLine("OCR has not been run; the result includes an optional OCR candidate.");
        return markdown.ToString().TrimEnd();
    }

    private static string BuildPlainText(string sourceName, OfficeImageInfo info, long length) {
        string dimensions = info.Width > 0 && info.Height > 0
            ? info.Width.ToString(CultureInfo.InvariantCulture) + " x " + info.Height.ToString(CultureInfo.InvariantCulture) + " pixels; "
            : string.Empty;
        return "Image " + Path.GetFileName(sourceName) + "; " + info.Format + "; " + dimensions +
            length.ToString(CultureInfo.InvariantCulture) + " bytes.";
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> BuildMetadata(OfficeImageInfo info, long length) {
        yield return Metadata("image-format", "Format", info.Format, "string");
        yield return Metadata("image-media-type", "MediaType", info.MimeType, "string");
        yield return Metadata("image-length", "LengthBytes", length, "number");
        if (info.Width > 0) yield return Metadata("image-width", "WidthPixels", info.Width, "number");
        if (info.Height > 0) yield return Metadata("image-height", "HeightPixels", info.Height, "number");
        if (info.Width > 0 || info.Height > 0) {
            yield return Metadata("image-dpi-x", "DpiX", info.DpiX, "number");
            yield return Metadata("image-dpi-y", "DpiY", info.DpiY, "number");
        }
        if (info.AspectRatio.HasValue) {
            yield return Metadata("image-aspect-ratio", "AspectRatio", info.AspectRatio.Value, "number");
        }
    }

    private static OfficeDocumentMetadataEntry Metadata(string id, string name, object value, string valueType) {
        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "image.properties",
            Name = name,
            Value = Convert.ToString(value, CultureInfo.InvariantCulture),
            ValueType = valueType
        };
    }
}
