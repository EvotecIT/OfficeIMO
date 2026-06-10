using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Reader;

public static partial class DocumentReader {
    private static IReadOnlyList<OfficeDocumentAsset> ReadOpenXmlImageAssets(string path, ReaderInputKind kind, ReaderOptions opt, CancellationToken cancellationToken) {
        if (kind != ReaderInputKind.Word && kind != ReaderInputKind.PowerPoint && kind != ReaderInputKind.Excel) {
            return Array.Empty<OfficeDocumentAsset>();
        }

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return ReadOpenXmlImageAssets(stream, path, kind, opt, cancellationToken);
    }

    private static IReadOnlyList<OfficeDocumentAsset> ReadOpenXmlImageAssets(Stream stream, string sourceName, ReaderInputKind kind, ReaderOptions opt, CancellationToken cancellationToken) {
        if (kind != ReaderInputKind.Word && kind != ReaderInputKind.PowerPoint && kind != ReaderInputKind.Excel) {
            return Array.Empty<OfficeDocumentAsset>();
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (stream.CanSeek) {
            stream.Position = 0;
        }

        var assets = new List<OfficeDocumentAsset>();
        OpenSettings? openSettings = CreateOpenSettings(opt);
        if (kind == ReaderInputKind.Word) {
            using WordprocessingDocument document = openSettings == null
                ? WordprocessingDocument.Open(stream, false)
                : WordprocessingDocument.Open(stream, false, openSettings);
            CollectWordImageAssets(document, sourceName, assets, cancellationToken);
        } else if (kind == ReaderInputKind.PowerPoint) {
            using PresentationDocument document = openSettings == null
                ? PresentationDocument.Open(stream, false)
                : PresentationDocument.Open(stream, false, openSettings);
            CollectPowerPointImageAssets(document, sourceName, assets, cancellationToken);
        } else if (kind == ReaderInputKind.Excel) {
            using SpreadsheetDocument document = openSettings == null
                ? SpreadsheetDocument.Open(stream, false)
                : SpreadsheetDocument.Open(stream, false, openSettings);
            CollectExcelImageAssets(document, sourceName, assets, cancellationToken);
        }

        return assets.Count == 0 ? Array.Empty<OfficeDocumentAsset>() : assets;
    }

    private static void CollectWordImageAssets(WordprocessingDocument document, string sourceName, List<OfficeDocumentAsset> assets, CancellationToken cancellationToken) {
        if (document.MainDocumentPart == null) {
            return;
        }

        int assetIndex = assets.Count;
        var seenImages = new HashSet<Uri>();
        var visitedParts = new HashSet<Uri>();
        CollectImageParts(
            document.MainDocumentPart,
            sourceName,
            ReaderInputKind.Word,
            slideNumber: null,
            sheetNumber: null,
            sheetName: null,
            assets,
            seenImages,
            visitedParts,
            ref assetIndex,
            cancellationToken);
    }

    private static void CollectPowerPointImageAssets(PresentationDocument document, string sourceName, List<OfficeDocumentAsset> assets, CancellationToken cancellationToken) {
        PresentationPart? presentationPart = document.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList == null) {
            return;
        }

        int assetIndex = assets.Count;
        var seenImages = new HashSet<Uri>();
        var visitedParts = new HashSet<Uri>();
        int slideNumber = 1;
        foreach (SlideId slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>()) {
            cancellationToken.ThrowIfCancellationRequested();

            string? relationshipId = slideId.RelationshipId?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                slideNumber++;
                continue;
            }

            if (presentationPart.GetPartById(relationshipId!) is SlidePart slidePart) {
                CollectImageParts(
                    slidePart,
                    sourceName,
                    ReaderInputKind.PowerPoint,
                    slideNumber,
                    sheetNumber: null,
                    sheetName: null,
                    assets,
                    seenImages,
                    visitedParts,
                    ref assetIndex,
                    cancellationToken);
            }

            slideNumber++;
        }
    }

    private static void CollectExcelImageAssets(SpreadsheetDocument document, string sourceName, List<OfficeDocumentAsset> assets, CancellationToken cancellationToken) {
        WorkbookPart? workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook?.Sheets == null) {
            return;
        }

        int assetIndex = assets.Count;
        var seenImages = new HashSet<Uri>();
        var visitedParts = new HashSet<Uri>();
        int sheetNumber = 1;
        foreach (Sheet sheet in workbookPart.Workbook.Sheets.Elements<Sheet>()) {
            cancellationToken.ThrowIfCancellationRequested();

            string? relationshipId = sheet.Id?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                sheetNumber++;
                continue;
            }

            if (workbookPart.GetPartById(relationshipId!) is WorksheetPart worksheetPart) {
                IReadOnlyDictionary<string, OpenXmlImageAssetMetadata> imageMetadata = BuildExcelImageMetadata(worksheetPart.DrawingsPart);
                CollectImageParts(
                    worksheetPart,
                    sourceName,
                    ReaderInputKind.Excel,
                    slideNumber: null,
                    sheetNumber,
                    sheet.Name?.Value,
                    assets,
                    seenImages,
                    visitedParts,
                    (container, imageRelationshipId) => ResolveExcelImageMetadata(container, worksheetPart.DrawingsPart, imageMetadata, imageRelationshipId),
                    ref assetIndex,
                    cancellationToken);
            }

            sheetNumber++;
        }
    }

    private static void CollectImageParts(
        OpenXmlPartContainer container,
        string sourceName,
        ReaderInputKind kind,
        int? slideNumber,
        int? sheetNumber,
        string? sheetName,
        List<OfficeDocumentAsset> assets,
        HashSet<Uri> seenImages,
        HashSet<Uri> visitedParts,
        ref int assetIndex,
        CancellationToken cancellationToken) {
        CollectImageParts(container, sourceName, kind, slideNumber, sheetNumber, sheetName, assets, seenImages, visitedParts, resolveMetadata: null, ref assetIndex, cancellationToken);
    }

    private static void CollectImageParts(
        OpenXmlPartContainer container,
        string sourceName,
        ReaderInputKind kind,
        int? slideNumber,
        int? sheetNumber,
        string? sheetName,
        List<OfficeDocumentAsset> assets,
        HashSet<Uri> seenImages,
        HashSet<Uri> visitedParts,
        Func<OpenXmlPartContainer, string, OpenXmlImageAssetMetadata?>? resolveMetadata,
        ref int assetIndex,
        CancellationToken cancellationToken) {
        foreach (IdPartPair pair in container.Parts) {
            cancellationToken.ThrowIfCancellationRequested();

            OpenXmlPart part = pair.OpenXmlPart;
            if (part is ImagePart imagePart) {
                Uri uri = imagePart.Uri;
                if (seenImages.Add(uri)) {
                    assets.Add(BuildOpenXmlImageAsset(imagePart, pair.RelationshipId, sourceName, kind, slideNumber, sheetNumber, sheetName, assetIndex, resolveMetadata?.Invoke(container, pair.RelationshipId)));
                    assetIndex++;
                }

                continue;
            }

            if (part is OpenXmlPartContainer childContainer) {
                if (visitedParts.Add(part.Uri)) {
                    CollectImageParts(childContainer, sourceName, kind, slideNumber, sheetNumber, sheetName, assets, seenImages, visitedParts, resolveMetadata, ref assetIndex, cancellationToken);
                }
            }
        }
    }

    private static OfficeDocumentAsset BuildOpenXmlImageAsset(ImagePart imagePart, string relationshipId, string sourceName, ReaderInputKind kind, int? slideNumber, int? sheetNumber, string? sheetName, int assetIndex, OpenXmlImageAssetMetadata? metadata = null) {
        byte[] payload;
        using (Stream imageStream = imagePart.GetStream(FileMode.Open, FileAccess.Read)) {
            using var memory = new MemoryStream();
            imageStream.CopyTo(memory);
            payload = memory.ToArray();
        }

        string kindStem = kind.ToString().ToLowerInvariant();
        int? width = null;
        int? height = null;
        if (OfficeDocumentImageDimensions.TryReadPixelDimensions(payload, imagePart.ContentType, out int detectedWidth, out int detectedHeight)) {
            width = detectedWidth;
            height = detectedHeight;
        }

        string assetId;
        if (slideNumber.HasValue) {
            assetId = string.Concat(kindStem, "-slide-", slideNumber.Value.ToString("D4", CultureInfo.InvariantCulture), "-image-", assetIndex.ToString("D4", CultureInfo.InvariantCulture));
        } else if (sheetNumber.HasValue) {
            assetId = string.Concat(kindStem, "-sheet-", sheetNumber.Value.ToString("D4", CultureInfo.InvariantCulture), "-image-", assetIndex.ToString("D4", CultureInfo.InvariantCulture));
        } else {
            assetId = string.Concat(kindStem, "-image-", assetIndex.ToString("D4", CultureInfo.InvariantCulture));
        }
        string? extension = ResolveImageExtension(imagePart.ContentType, imagePart.Uri);

        return new OfficeDocumentAsset {
            Id = assetId,
            Kind = "image",
            MediaType = imagePart.ContentType,
            Extension = extension,
            FileName = OfficeDocumentAssetNaming.BuildFileName(assetId, extension),
            AltText = metadata?.AltText,
            Title = metadata?.Title,
            Width = width,
            Height = height,
            LengthBytes = payload.Length,
            PayloadHash = OfficeDocumentAssetHash.ComputeSha256Hex(payload),
            PayloadBytes = payload,
            SourceObjectId = relationshipId + "|" + imagePart.Uri,
            Location = new ReaderLocation {
                Path = sourceName,
                Slide = slideNumber,
                Sheet = sheetName,
                SourceBlockKind = "image",
                BlockAnchor = assetId
            }
        };
    }

    private sealed class OpenXmlImageAssetMetadata {
        public string? AltText { get; set; }

        public string? Title { get; set; }
    }

    private static IReadOnlyDictionary<string, OpenXmlImageAssetMetadata> BuildExcelImageMetadata(DrawingsPart? drawingsPart) {
        Xdr.WorksheetDrawing? drawing = drawingsPart?.WorksheetDrawing;
        if (drawing == null) {
            return new Dictionary<string, OpenXmlImageAssetMetadata>(StringComparer.Ordinal);
        }

        var metadata = new Dictionary<string, OpenXmlImageAssetMetadata>(StringComparer.Ordinal);
        foreach (Xdr.Picture picture in drawing.Descendants<Xdr.Picture>()) {
            string? relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                continue;
            }

            Xdr.NonVisualDrawingProperties? properties = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
            string? altText = NormalizeOptionalAssetText(properties?.Description?.Value);
            string? title = NormalizeOptionalAssetText(properties?.Title?.Value);
            if (altText == null && title == null) {
                continue;
            }

            metadata[relationshipId!] = new OpenXmlImageAssetMetadata {
                AltText = altText,
                Title = title
            };
        }

        return metadata;
    }

    private static OpenXmlImageAssetMetadata? ResolveExcelImageMetadata(OpenXmlPartContainer container, DrawingsPart? drawingsPart, IReadOnlyDictionary<string, OpenXmlImageAssetMetadata> metadata, string relationshipId) {
        if (drawingsPart == null || !ReferenceEquals(container, drawingsPart)) {
            return null;
        }

        return metadata.TryGetValue(relationshipId, out OpenXmlImageAssetMetadata? value) ? value : null;
    }

    private static string? NormalizeOptionalAssetText(string? value) {
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static string? ResolveImageExtension(string? contentType, Uri uri) {
        string? extension = contentType?.Trim().ToLowerInvariant() switch {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/jpg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/tiff" => ".tiff",
            "image/svg+xml" => ".svg",
            "image/x-emf" => ".emf",
            "image/x-wmf" => ".wmf",
            _ => null
        };

        if (!string.IsNullOrWhiteSpace(extension)) {
            return extension;
        }

        string uriExtension = Path.GetExtension(uri.ToString());
        return string.IsNullOrWhiteSpace(uriExtension) ? null : uriExtension;
    }
}
