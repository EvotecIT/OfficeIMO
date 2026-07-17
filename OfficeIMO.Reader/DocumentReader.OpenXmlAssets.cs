using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DPic = DocumentFormat.OpenXml.Drawing.Pictures;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static IReadOnlyList<OfficeDocumentAsset> ReadOpenXmlImageAssets(string path, ReaderInputKind kind, ReaderOptions opt, CancellationToken cancellationToken) {
        if (kind != ReaderInputKind.Word && kind != ReaderInputKind.PowerPoint && kind != ReaderInputKind.Excel) {
            return Array.Empty<OfficeDocumentAsset>();
        }

        if (kind == ReaderInputKind.PowerPoint
            && (!string.IsNullOrEmpty(opt.OpenPassword)
                || IsLegacyPowerPointExtension(path))) {
            using PowerPointPresentation presentation =
                LoadPowerPointForReader(path, opt);
            return CollectProjectedPowerPointImageAssets(presentation, path,
                opt, cancellationToken);
        }
        if (IsLegacyBinaryOfficeExtension(path)) {
            return Array.Empty<OfficeDocumentAsset>();
        }

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return ReadOpenXmlImageAssets(stream, path, kind, opt, cancellationToken);
    }

    private static IReadOnlyList<OfficeDocumentAsset> ReadOpenXmlImageAssets(Stream stream, string sourceName, ReaderInputKind kind, ReaderOptions opt, CancellationToken cancellationToken) {
        if (kind != ReaderInputKind.Word && kind != ReaderInputKind.PowerPoint && kind != ReaderInputKind.Excel) {
            return Array.Empty<OfficeDocumentAsset>();
        }

        if (kind == ReaderInputKind.PowerPoint
            && (!string.IsNullOrEmpty(opt.OpenPassword)
                || IsLegacyPowerPointExtension(sourceName)
                || IsLegacyPowerPointCompound(stream, opt))) {
            using PowerPointPresentation presentation =
                LoadPowerPointForReader(stream, opt);
            return CollectProjectedPowerPointImageAssets(presentation,
                sourceName, opt, cancellationToken);
        }
        if (IsLegacyBinaryOfficeExtension(sourceName)) {
            return Array.Empty<OfficeDocumentAsset>();
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (stream.CanSeek) {
            stream.Position = 0;
        }

        var assets = new List<OfficeDocumentAsset>();
        var payloadCache = new Dictionary<Uri, OpenXmlImagePayload>();
        long totalPayloadBytes = 0;
        OpenSettings? openSettings = CreateOpenSettings(opt);
        if (kind == ReaderInputKind.Word) {
            using WordprocessingDocument document = openSettings == null
                ? WordprocessingDocument.Open(stream, false)
                : WordprocessingDocument.Open(stream, false, openSettings);
            CollectWordImageAssets(document, sourceName, opt, assets, payloadCache, ref totalPayloadBytes, cancellationToken);
        } else if (kind == ReaderInputKind.PowerPoint) {
            using PresentationDocument document = openSettings == null
                ? PresentationDocument.Open(stream, false)
                : PresentationDocument.Open(stream, false, openSettings);
            CollectPowerPointImageAssets(document, sourceName, opt, assets, payloadCache, ref totalPayloadBytes, cancellationToken);
        } else if (kind == ReaderInputKind.Excel) {
            try {
                using SpreadsheetDocument document = openSettings == null
                    ? SpreadsheetDocument.Open(stream, false)
                    : SpreadsheetDocument.Open(stream, false, openSettings);
                CollectExcelImageAssets(document, sourceName, opt, assets, payloadCache, ref totalPayloadBytes, cancellationToken);
            } catch (Exception exception) when (ShouldSkipExcelImageAssetsAfterPasswordedOpenFailure(exception, opt)) {
                return Array.Empty<OfficeDocumentAsset>();
            }
        }

        return assets.Count == 0 ? Array.Empty<OfficeDocumentAsset>() : assets;
    }

    private static bool IsLegacyPowerPointCompound(Stream stream,
        ReaderOptions options) {
        if (!stream.CanSeek) return false;
        long position = stream.Position;
        try {
            DetectionCandidate candidate = InspectOfficeCompound(stream,
                position, options.DetectionMaxContainerEntries);
            return candidate.Kind == ReaderInputKind.PowerPoint
                && string.Equals(candidate.MediaType,
                    "application/vnd.ms-powerpoint",
                    StringComparison.OrdinalIgnoreCase);
        } finally {
            stream.Position = position;
        }
    }

    private static IReadOnlyList<OfficeDocumentAsset>
        CollectProjectedPowerPointImageAssets(
            PowerPointPresentation presentation, string sourceName,
            ReaderOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var assets = new List<OfficeDocumentAsset>();
        var payloadCache = new Dictionary<Uri, OpenXmlImagePayload>();
        long totalPayloadBytes = 0;
        CollectPowerPointImageAssets(presentation.OpenXmlDocument,
            sourceName, options, assets, payloadCache,
            ref totalPayloadBytes, cancellationToken);
        return assets.Count == 0
            ? Array.Empty<OfficeDocumentAsset>()
            : assets;
    }

    private static bool ShouldSkipExcelImageAssetsAfterPasswordedOpenFailure(Exception exception, ReaderOptions opt) {
        return !string.IsNullOrEmpty(opt.OpenPassword)
            && (exception is OpenXmlPackageException || exception is FileFormatException);
    }

    private static void CollectWordImageAssets(WordprocessingDocument document, string sourceName, ReaderOptions opt, List<OfficeDocumentAsset> assets, Dictionary<Uri, OpenXmlImagePayload> payloadCache, ref long totalPayloadBytes, CancellationToken cancellationToken) {
        if (document.MainDocumentPart == null) {
            return;
        }

        int assetIndex = assets.Count;
        var seenImagePlacements = new HashSet<string>(StringComparer.Ordinal);
        var visitedParts = new HashSet<Uri>();
        CollectImageParts(
            document.MainDocumentPart,
            sourceName,
            ReaderInputKind.Word,
            slideNumber: null,
            sheetNumber: null,
            sheetName: null,
            assets,
            seenImagePlacements,
            visitedParts,
            opt,
            payloadCache,
            ref totalPayloadBytes,
            ref assetIndex,
            cancellationToken);
    }

    private static void CollectPowerPointImageAssets(PresentationDocument document, string sourceName, ReaderOptions opt, List<OfficeDocumentAsset> assets, Dictionary<Uri, OpenXmlImagePayload> payloadCache, ref long totalPayloadBytes, CancellationToken cancellationToken) {
        PresentationPart? presentationPart = document.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList == null) {
            return;
        }

        int assetIndex = assets.Count;
        var seenImagePlacements = new HashSet<string>(StringComparer.Ordinal);
        int slideNumber = 1;
        foreach (SlideId slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>()) {
            cancellationToken.ThrowIfCancellationRequested();

            string? relationshipId = slideId.RelationshipId?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                slideNumber++;
                continue;
            }

            if (presentationPart.GetPartById(relationshipId!) is SlidePart slidePart) {
                var visitedParts = new HashSet<Uri>();
                CollectImageParts(
                    slidePart,
                    sourceName,
                    ReaderInputKind.PowerPoint,
                    slideNumber,
                    sheetNumber: null,
                    sheetName: null,
                    assets,
                    seenImagePlacements,
                    visitedParts,
                    opt,
                    payloadCache,
                    ref totalPayloadBytes,
                    ref assetIndex,
                    cancellationToken);
            }

            slideNumber++;
        }
    }

    private static void CollectExcelImageAssets(SpreadsheetDocument document, string sourceName, ReaderOptions opt, List<OfficeDocumentAsset> assets, Dictionary<Uri, OpenXmlImagePayload> payloadCache, ref long totalPayloadBytes, CancellationToken cancellationToken) {
        WorkbookPart? workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook?.Sheets == null) {
            return;
        }

        int assetIndex = assets.Count;
        var seenImagePlacements = new HashSet<string>(StringComparer.Ordinal);
        string? selectedSheetName = string.IsNullOrWhiteSpace(opt.ExcelSheetName) ? null : opt.ExcelSheetName!.Trim();
        int sheetNumber = 1;
        foreach (Sheet sheet in workbookPart.Workbook.Sheets.Elements<Sheet>()) {
            cancellationToken.ThrowIfCancellationRequested();

            string? sheetName = sheet.Name?.Value;
            if (selectedSheetName != null && !string.Equals(sheetName, selectedSheetName, StringComparison.OrdinalIgnoreCase)) {
                sheetNumber++;
                continue;
            }

            string? relationshipId = sheet.Id?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                sheetNumber++;
                continue;
            }

            if (workbookPart.GetPartById(relationshipId!) is WorksheetPart worksheetPart) {
                IReadOnlyDictionary<string, IReadOnlyList<OpenXmlImageAssetMetadata>> imageMetadata = BuildExcelImageMetadata(worksheetPart.DrawingsPart);
                (int Row1, int Column1, int Row2, int Column2)? selectedRange = TryParseExcelAssetRange(opt.ExcelA1Range);
                var visitedParts = new HashSet<Uri>();
                CollectImageParts(
                    worksheetPart,
                    sourceName,
                    ReaderInputKind.Excel,
                    slideNumber: null,
                    sheetNumber,
                    sheetName,
                    assets,
                    seenImagePlacements,
                    visitedParts,
                    opt,
                    (container, imageRelationshipId) => ResolveExcelImageMetadata(container, worksheetPart.DrawingsPart, imageMetadata, imageRelationshipId),
                    metadata => IsExcelImageInSelectedRange(metadata, selectedRange),
                    payloadCache,
                    ref totalPayloadBytes,
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
        HashSet<string> seenImagePlacements,
        HashSet<Uri> visitedParts,
        ReaderOptions opt,
        Dictionary<Uri, OpenXmlImagePayload> payloadCache,
        ref long totalPayloadBytes,
        ref int assetIndex,
        CancellationToken cancellationToken) {
        CollectImageParts(container, sourceName, kind, slideNumber, sheetNumber, sheetName, assets, seenImagePlacements, visitedParts, opt, resolveMetadata: null, shouldIncludeMetadata: null, payloadCache, ref totalPayloadBytes, ref assetIndex, cancellationToken);
    }

    private static void CollectImageParts(
        OpenXmlPartContainer container,
        string sourceName,
        ReaderInputKind kind,
        int? slideNumber,
        int? sheetNumber,
        string? sheetName,
        List<OfficeDocumentAsset> assets,
        HashSet<string> seenImagePlacements,
        HashSet<Uri> visitedParts,
        ReaderOptions opt,
        Func<OpenXmlPartContainer, string, IReadOnlyList<OpenXmlImageAssetMetadata>?>? resolveMetadata,
        Func<OpenXmlImageAssetMetadata?, bool>? shouldIncludeMetadata,
        Dictionary<Uri, OpenXmlImagePayload> payloadCache,
        ref long totalPayloadBytes,
        ref int assetIndex,
        CancellationToken cancellationToken) {
        foreach (IdPartPair pair in container.Parts) {
            cancellationToken.ThrowIfCancellationRequested();

            OpenXmlPart part = pair.OpenXmlPart;
            if (part is ImagePart imagePart) {
                IReadOnlyList<OpenXmlImageAssetMetadata>? metadataPlacements = resolveMetadata?.Invoke(container, pair.RelationshipId)
                    ?? ResolveOpenXmlImageMetadataPlacements(container, pair.RelationshipId, opt);
                int placementCount = metadataPlacements?.Count ?? Math.Max(1, CountImageRelationshipPlacements(container, pair.RelationshipId, opt));
                EnsureOpenXmlImagePlacementCount(opt, placementCount, pair.RelationshipId);
                for (int placementIndex = 0; placementIndex < placementCount; placementIndex++) {
                    OpenXmlImageAssetMetadata? metadata = metadataPlacements != null && placementIndex < metadataPlacements.Count
                        ? metadataPlacements[placementIndex]
                        : null;
                    if (shouldIncludeMetadata != null && !shouldIncludeMetadata(metadata)) {
                        continue;
                    }

                    if (!seenImagePlacements.Add(BuildOpenXmlImagePlacementKey(kind, slideNumber, sheetNumber, container, pair.RelationshipId, imagePart, placementIndex))) {
                        continue;
                    }

                    EnsureOpenXmlImageAssetCount(opt, assets.Count + 1);
                    OpenXmlImagePayload payload = GetOpenXmlImagePayload(imagePart, opt, payloadCache, ref totalPayloadBytes);
                    assets.Add(BuildOpenXmlImageAsset(imagePart, pair.RelationshipId, sourceName, kind, slideNumber, sheetNumber, sheetName, assetIndex, payload, metadata));
                    assetIndex++;
                }

                continue;
            }

            if (part is OpenXmlPartContainer childContainer) {
                if (ShouldTraverseRelatedPart(kind, opt, part) && visitedParts.Add(part.Uri)) {
                    CollectImageParts(childContainer, sourceName, kind, slideNumber, sheetNumber, sheetName, assets, seenImagePlacements, visitedParts, opt, resolveMetadata, shouldIncludeMetadata, payloadCache, ref totalPayloadBytes, ref assetIndex, cancellationToken);
                }
            }
        }
    }

    private static string BuildOpenXmlImagePlacementKey(ReaderInputKind kind, int? slideNumber, int? sheetNumber, OpenXmlPartContainer container, string relationshipId, ImagePart imagePart, int placementIndex) {
        string containerUri = container is OpenXmlPart part ? part.Uri.ToString() : "package";
        return string.Concat(
            kind.ToString(),
            "|slide:", slideNumber?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            "|sheet:", sheetNumber?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            "|container:", containerUri,
            "|relationship:", relationshipId,
            "|image:", imagePart.Uri,
            "|placement:", placementIndex.ToString(CultureInfo.InvariantCulture));
    }

    private static OfficeDocumentAsset BuildOpenXmlImageAsset(ImagePart imagePart, string relationshipId, string sourceName, ReaderInputKind kind, int? slideNumber, int? sheetNumber, string? sheetName, int assetIndex, OpenXmlImagePayload payload, OpenXmlImageAssetMetadata? metadata = null) {
        string kindStem = kind.ToString().ToLowerInvariant();
        string assetId;
        if (slideNumber.HasValue) {
            assetId = string.Concat(kindStem, "-slide-", slideNumber.Value.ToString("D4", CultureInfo.InvariantCulture), "-image-", assetIndex.ToString("D4", CultureInfo.InvariantCulture));
        } else if (sheetNumber.HasValue) {
            assetId = string.Concat(kindStem, "-sheet-", sheetNumber.Value.ToString("D4", CultureInfo.InvariantCulture), "-image-", assetIndex.ToString("D4", CultureInfo.InvariantCulture));
        } else {
            assetId = string.Concat(kindStem, "-image-", assetIndex.ToString("D4", CultureInfo.InvariantCulture));
        }

        return new OfficeDocumentAsset {
            Id = assetId,
            Kind = "image",
            MediaType = imagePart.ContentType,
            Extension = payload.Extension,
            FileName = OfficeDocumentAssetNaming.BuildFileName(assetId, payload.Extension),
            AltText = metadata?.AltText,
            Title = metadata?.Title,
            Width = payload.Width,
            Height = payload.Height,
            LengthBytes = payload.Bytes.Length,
            PayloadHash = payload.Hash,
            PayloadBytes = payload.Bytes,
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

    private static OpenXmlImagePayload GetOpenXmlImagePayload(ImagePart imagePart, ReaderOptions opt, Dictionary<Uri, OpenXmlImagePayload> payloadCache, ref long totalPayloadBytes) {
        if (payloadCache.TryGetValue(imagePart.Uri, out OpenXmlImagePayload? cached)) {
            return cached;
        }

        byte[] payload;
        using (Stream imageStream = imagePart.GetStream(FileMode.Open, FileAccess.Read)) {
            payload = CopyOpenXmlImagePayload(imageStream, opt.MaxOpenXmlImageAssetBytes);
        }

        if (opt.MaxOpenXmlImageTotalAssetBytes.HasValue) {
            long nextTotal = checked(totalPayloadBytes + payload.LongLength);
            if (nextTotal > opt.MaxOpenXmlImageTotalAssetBytes.Value) {
                throw new IOException($"OpenXML image asset extraction exceeds MaxOpenXmlImageTotalAssetBytes ({nextTotal.ToString(CultureInfo.InvariantCulture)} > {opt.MaxOpenXmlImageTotalAssetBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }

            totalPayloadBytes = nextTotal;
        } else {
            totalPayloadBytes = checked(totalPayloadBytes + payload.LongLength);
        }

        int? width = null;
        int? height = null;
        if (OfficeDocumentImageDimensions.TryReadPixelDimensions(payload, imagePart.ContentType, out int detectedWidth, out int detectedHeight)) {
            width = detectedWidth;
            height = detectedHeight;
        }

        var value = new OpenXmlImagePayload(
            payload,
            OfficeDocumentAssetHash.ComputeSha256Hex(payload),
            ResolveImageExtension(imagePart.ContentType, imagePart.Uri),
            width,
            height);
        payloadCache.Add(imagePart.Uri, value);
        return value;
    }

    private static byte[] CopyOpenXmlImagePayload(Stream imageStream, long? maxBytes) {
        if (maxBytes.HasValue && imageStream.CanSeek) {
            long remaining = imageStream.Length - imageStream.Position;
            if (remaining > maxBytes.Value) {
                throw new IOException($"OpenXML image asset exceeds MaxOpenXmlImageAssetBytes ({remaining.ToString(CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }
        }

        using var memory = new MemoryStream();
        var buffer = new byte[81920];
        long total = 0;
        while (true) {
            int read = imageStream.Read(buffer, 0, buffer.Length);
            if (read == 0) {
                break;
            }

            total += read;
            if (maxBytes.HasValue && total > maxBytes.Value) {
                throw new IOException($"OpenXML image asset exceeds MaxOpenXmlImageAssetBytes ({total.ToString(CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }

            memory.Write(buffer, 0, read);
        }

        return memory.ToArray();
    }

    private static void EnsureOpenXmlImageAssetCount(ReaderOptions opt, int nextCount) {
        if (opt.MaxOpenXmlImageAssets.HasValue && nextCount > opt.MaxOpenXmlImageAssets.Value) {
            throw new IOException($"OpenXML image asset extraction exceeds MaxOpenXmlImageAssets ({nextCount.ToString(CultureInfo.InvariantCulture)} > {opt.MaxOpenXmlImageAssets.Value.ToString(CultureInfo.InvariantCulture)}).");
        }
    }

    private static void EnsureOpenXmlImagePlacementCount(ReaderOptions opt, int count, string relationshipId) {
        if (opt.MaxOpenXmlImagePlacementsPerRelationship.HasValue && count > opt.MaxOpenXmlImagePlacementsPerRelationship.Value) {
            throw new IOException($"OpenXML image relationship '{relationshipId}' exceeds MaxOpenXmlImagePlacementsPerRelationship ({count.ToString(CultureInfo.InvariantCulture)} > {opt.MaxOpenXmlImagePlacementsPerRelationship.Value.ToString(CultureInfo.InvariantCulture)}).");
        }
    }

    private sealed class OpenXmlImageAssetMetadata {
        public string? AltText { get; set; }

        public string? Title { get; set; }

        public int? AnchorRow { get; set; }

        public int? AnchorColumn { get; set; }
    }

    private sealed class OpenXmlImagePayload {
        public OpenXmlImagePayload(byte[] bytes, string hash, string? extension, int? width, int? height) {
            Bytes = bytes;
            Hash = hash;
            Extension = extension;
            Width = width;
            Height = height;
        }

        public byte[] Bytes { get; }
        public string Hash { get; }
        public string? Extension { get; }
        public int? Width { get; }
        public int? Height { get; }
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<OpenXmlImageAssetMetadata>> BuildExcelImageMetadata(DrawingsPart? drawingsPart) {
        Xdr.WorksheetDrawing? drawing = drawingsPart?.WorksheetDrawing;
        if (drawing == null) {
            return new Dictionary<string, IReadOnlyList<OpenXmlImageAssetMetadata>>(StringComparer.Ordinal);
        }

        var metadata = new Dictionary<string, List<OpenXmlImageAssetMetadata>>(StringComparer.Ordinal);
        foreach (Xdr.Picture picture in drawing.Descendants<Xdr.Picture>()) {
            string? relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                continue;
            }

            Xdr.NonVisualDrawingProperties? properties = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
            string? altText = NormalizeOptionalAssetText(properties?.Description?.Value);
            string? title = NormalizeOptionalAssetText(properties?.Title?.Value);
            if (altText == null && title == null) {
                // Keep anchor metadata even when the picture has no descriptive text.
            }

            if (!metadata.TryGetValue(relationshipId!, out List<OpenXmlImageAssetMetadata>? placements)) {
                placements = new List<OpenXmlImageAssetMetadata>();
                metadata[relationshipId!] = placements;
            }

            placements.Add(new OpenXmlImageAssetMetadata {
                AltText = altText,
                Title = title,
                AnchorRow = GetExcelImageAnchorRow(picture),
                AnchorColumn = GetExcelImageAnchorColumn(picture)
            });
        }

        return metadata.ToDictionary(
            pair => pair.Key,
            pair => (IReadOnlyList<OpenXmlImageAssetMetadata>)pair.Value.ToArray(),
            StringComparer.Ordinal);
    }

    private static IReadOnlyList<OpenXmlImageAssetMetadata>? ResolveExcelImageMetadata(OpenXmlPartContainer container, DrawingsPart? drawingsPart, IReadOnlyDictionary<string, IReadOnlyList<OpenXmlImageAssetMetadata>> metadata, string relationshipId) {
        if (drawingsPart == null || !ReferenceEquals(container, drawingsPart)) {
            return null;
        }

        return metadata.TryGetValue(relationshipId, out IReadOnlyList<OpenXmlImageAssetMetadata>? value) ? value : null;
    }

    private static IReadOnlyList<OpenXmlImageAssetMetadata>? ResolveOpenXmlImageMetadataPlacements(OpenXmlPartContainer container, string relationshipId, ReaderOptions opt) {
        OpenXmlPartRootElement? root = (container as OpenXmlPart)?.RootElement;
        if (root == null) {
            return null;
        }

        List<OpenXmlImageAssetMetadata>? placements = null;
        foreach (A.Blip blip in root.Descendants<A.Blip>()) {
            if (!string.Equals(blip.Embed?.Value, relationshipId, StringComparison.Ordinal) &&
                !string.Equals(blip.Link?.Value, relationshipId, StringComparison.Ordinal)) {
                continue;
            }

            placements ??= new List<OpenXmlImageAssetMetadata>();
            EnsureOpenXmlImagePlacementCount(opt, placements.Count + 1, relationshipId);
            placements.Add(ResolveOpenXmlImageMetadata(blip));
        }

        return placements == null || placements.Count == 0 ? null : placements;
    }

    private static OpenXmlImageAssetMetadata ResolveOpenXmlImageMetadata(A.Blip blip) {
        DW.DocProperties? wordProperties = blip.Ancestors<DW.Inline>().FirstOrDefault()?.DocProperties
            ?? blip.Ancestors<DW.Anchor>().FirstOrDefault()?.GetFirstChild<DW.DocProperties>();
        DPic.NonVisualDrawingProperties? drawingProperties = blip.Ancestors<DPic.Picture>().FirstOrDefault()?.NonVisualPictureProperties?.NonVisualDrawingProperties;
        NonVisualDrawingProperties? presentationProperties = blip.Ancestors<DocumentFormat.OpenXml.Presentation.Picture>().FirstOrDefault()?.NonVisualPictureProperties?.NonVisualDrawingProperties;

        return new OpenXmlImageAssetMetadata {
            AltText = NormalizeOptionalAssetText(
                wordProperties?.Description?.Value ??
                drawingProperties?.Description?.Value ??
                presentationProperties?.Description?.Value),
            Title = NormalizeOptionalAssetText(
                wordProperties?.Title?.Value ??
                drawingProperties?.Title?.Value ??
                presentationProperties?.Title?.Value)
        };
    }

    private static bool ShouldTraverseRelatedPart(ReaderInputKind kind, ReaderOptions opt, OpenXmlPart part) {
        if (kind == ReaderInputKind.Word && !opt.IncludeWordFootnotes && (part is FootnotesPart || part is EndnotesPart)) {
            return false;
        }

        if (kind == ReaderInputKind.PowerPoint && !opt.IncludePowerPointNotes && part is NotesSlidePart) {
            return false;
        }

        return true;
    }

    private static int CountImageRelationshipPlacements(OpenXmlPartContainer container, string relationshipId, ReaderOptions opt) {
        OpenXmlPartRootElement? root = (container as OpenXmlPart)?.RootElement;
        if (root == null) {
            return 0;
        }

        int count = 0;
        foreach (A.Blip blip in root.Descendants<A.Blip>()) {
            if (string.Equals(blip.Embed?.Value, relationshipId, StringComparison.Ordinal) ||
                string.Equals(blip.Link?.Value, relationshipId, StringComparison.Ordinal)) {
                count++;
                EnsureOpenXmlImagePlacementCount(opt, count, relationshipId);
            }
        }

        return count;
    }

    private static (int Row1, int Column1, int Row2, int Column2)? TryParseExcelAssetRange(string? a1Range) {
        if (string.IsNullOrWhiteSpace(a1Range)) {
            return null;
        }

        return A1.TryParseRange(a1Range!, out int row1, out int column1, out int row2, out int column2)
            ? (row1, column1, row2, column2)
            : null;
    }

    private static bool IsExcelImageInSelectedRange(OpenXmlImageAssetMetadata? metadata, (int Row1, int Column1, int Row2, int Column2)? selectedRange) {
        if (selectedRange == null) {
            return true;
        }

        if (metadata?.AnchorRow == null || metadata.AnchorColumn == null) {
            return false;
        }

        var range = selectedRange.Value;
        return metadata.AnchorRow.Value >= range.Row1 &&
            metadata.AnchorRow.Value <= range.Row2 &&
            metadata.AnchorColumn.Value >= range.Column1 &&
            metadata.AnchorColumn.Value <= range.Column2;
    }

    private static int? GetExcelImageAnchorRow(Xdr.Picture picture) {
        Xdr.FromMarker? marker = picture.Ancestors<Xdr.TwoCellAnchor>().FirstOrDefault()?.FromMarker
            ?? picture.Ancestors<Xdr.OneCellAnchor>().FirstOrDefault()?.FromMarker;
        return TryParseZeroBasedMarkerIndex(marker?.RowId?.Text);
    }

    private static int? GetExcelImageAnchorColumn(Xdr.Picture picture) {
        Xdr.FromMarker? marker = picture.Ancestors<Xdr.TwoCellAnchor>().FirstOrDefault()?.FromMarker
            ?? picture.Ancestors<Xdr.OneCellAnchor>().FirstOrDefault()?.FromMarker;
        return TryParseZeroBasedMarkerIndex(marker?.ColumnId?.Text);
    }

    private static int? TryParseZeroBasedMarkerIndex(string? value) {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int zeroBased) && zeroBased >= 0
            ? zeroBased + 1
            : null;
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
