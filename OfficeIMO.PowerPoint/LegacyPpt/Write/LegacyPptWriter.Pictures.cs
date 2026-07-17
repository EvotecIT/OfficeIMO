using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort OfficeArtBStoreContainer = 0xF001;

        internal static bool TryReadPictureCatalog(
            PowerPointPresentation presentation,
            out LegacyPptWriterPictureCatalog catalog,
            out LegacyPptFeature failureFeature,
            out string? reason) {
            if (presentation == null) throw new ArgumentNullException(
                nameof(presentation));
            catalog = new LegacyPptWriterPictureCatalog();
            var materializedLayoutPictures = new HashSet<OpenXmlElement>(
                ReferenceComparer.Instance);
            foreach (PowerPointSlide slide in presentation.Slides) {
                IReadOnlyList<PowerPointShape> shapes = ReadSlideShapesForWrite(
                    slide, out string? shapeReason);
                if (shapeReason != null) {
                    failureFeature = LegacyPptFeature.RasterPictures;
                    reason = shapeReason;
                    return false;
                }
                shapes = FlattenShapeTreeForWrite(shapes,
                    out shapeReason);
                if (shapeReason != null) {
                    failureFeature = LegacyPptFeature.Groups;
                    reason = shapeReason;
                    return false;
                }
                foreach (PowerPointShape shape in shapes) {
                    if (shape is PowerPointPicture
                        && IsLayoutShape(shape)) {
                        materializedLayoutPictures.Add(shape.Element);
                    }
                    byte[] imageBytes;
                    string contentType;
                    if (shape is PowerPointPicture picture
                        && picture is not PowerPointMedia) {
                        failureFeature = LegacyPptFeature.RasterPictures;
                        if (!TryReadPicture(picture, out imageBytes,
                                out string? pictureContentType, out reason)) {
                            return false;
                        }
                        contentType = pictureContentType!;
                    } else if (shape is PowerPointChart chart) {
                        failureFeature = LegacyPptFeature.Charts;
                        if (!TryRenderChartPicture(chart, out imageBytes,
                                out reason)) {
                            return false;
                        }
                        contentType = "image/png";
                    } else if (shape is PowerPointSmartArt smartArt) {
                        failureFeature = LegacyPptFeature.SmartArt;
                        if (!TryRenderSmartArtPicture(smartArt, out imageBytes,
                                out reason)) {
                            return false;
                        }
                        contentType = "image/png";
                    } else if (shape is PowerPointOleObject ole) {
                        failureFeature = LegacyPptFeature.EmbeddedOle;
                        if (!TryReadOlePreview(slide.SlidePart, ole,
                                out PowerPointPicture? preview,
                                out imageBytes,
                                out string? previewContentType,
                                out reason)) {
                            return false;
                        }
                        if (preview == null) continue;
                        contentType = previewContentType!;
                    } else {
                        continue;
                    }
                    if (!catalog.TryAdd(shape, imageBytes, contentType,
                            out reason)) {
                        return false;
                    }
                }
                if (!TryReadBackground(slide,
                        out LegacyPptWriterBackground? slideBackground,
                        out reason)
                    || !TryAddBackgroundPicture(catalog, slideBackground,
                        out reason)) {
                    failureFeature = LegacyPptFeature.Backgrounds;
                    return false;
                }
                NotesSlidePart? notesPart = slide.SlidePart.NotesSlidePart;
                if (notesPart != null
                    && ShouldWriteNotesPage(slide, out _)
                    && (!TryReadBackground(notesPart,
                            out LegacyPptWriterBackground? notesBackground,
                            out reason)
                        || !TryAddBackgroundPicture(catalog, notesBackground,
                            out reason))) {
                    failureFeature = LegacyPptFeature.Backgrounds;
                    return false;
                }
            }

            PresentationPart? presentationPart = presentation.OpenXmlDocument
                .PresentationPart;
            if (!TryValidateMaterializedLayoutPictures(presentation,
                    materializedLayoutPictures, out reason)) {
                failureFeature = LegacyPptFeature.Layouts;
                return false;
            }
            foreach (SlideMasterPart masterPart in presentationPart?
                         .SlideMasterParts ?? Enumerable.Empty<SlideMasterPart>()) {
                IReadOnlyList<PowerPointShape> masterShapes =
                    ReadMasterShapesForWrite(masterPart,
                        out string? masterShapeReason);
                if (masterShapeReason != null) {
                    failureFeature = LegacyPptFeature.Masters;
                    reason = masterShapeReason;
                    return false;
                }
                if (!TryAddMasterPictures(catalog, masterShapes,
                        out failureFeature, out reason)) {
                    return false;
                }
                if (!TryReadBackground(masterPart,
                        out LegacyPptWriterBackground? masterBackground,
                        out reason)
                    || !TryAddBackgroundPicture(catalog, masterBackground,
                        out reason)) {
                    failureFeature = LegacyPptFeature.Backgrounds;
                    return false;
                }
            }
            NotesMasterPart? notesMasterPart = presentationPart?
                .NotesMasterPart;
            if (notesMasterPart != null) {
                IReadOnlyList<PowerPointShape> notesMasterShapes =
                    ReadMasterShapesForWrite(notesMasterPart,
                        out string? notesMasterShapeReason);
                if (notesMasterShapeReason != null) {
                    failureFeature = LegacyPptFeature.Masters;
                    reason = notesMasterShapeReason;
                    return false;
                }
                if (!TryAddMasterPictures(catalog, notesMasterShapes,
                        out failureFeature, out reason)) {
                    return false;
                }
                if (!TryReadBackground(notesMasterPart,
                        out LegacyPptWriterBackground? notesMasterBackground,
                        out reason)
                    || !TryAddBackgroundPicture(catalog,
                        notesMasterBackground, out reason)) {
                    failureFeature = LegacyPptFeature.Backgrounds;
                    return false;
                }
            }
            HandoutMasterPart? handoutMasterPart = presentationPart?
                .HandoutMasterPart;
            if (handoutMasterPart != null) {
                IReadOnlyList<PowerPointShape> handoutMasterShapes =
                    ReadMasterShapesForWrite(handoutMasterPart,
                        out string? handoutMasterShapeReason);
                if (handoutMasterShapeReason != null) {
                    failureFeature = LegacyPptFeature.Masters;
                    reason = handoutMasterShapeReason;
                    return false;
                }
                if (!TryAddMasterPictures(catalog, handoutMasterShapes,
                        out failureFeature, out reason)) {
                    return false;
                }
                if (!TryReadBackground(handoutMasterPart,
                        out LegacyPptWriterBackground? handoutBackground,
                        out reason)
                    || !TryAddBackgroundPicture(catalog,
                        handoutBackground, out reason)) {
                    failureFeature = LegacyPptFeature.Backgrounds;
                    return false;
                }
            }
            failureFeature = LegacyPptFeature.RasterPictures;
            reason = null;
            return true;
        }

        private static bool TryAddBackgroundPicture(
            LegacyPptWriterPictureCatalog catalog,
            LegacyPptWriterBackground? background, out string? reason) {
            if (background?.PictureFill == null) {
                reason = null;
                return true;
            }
            return catalog.TryAdd(background.PictureFill,
                background.PictureBytes, background.PictureContentType!,
                out reason);
        }

        internal static bool TryRenderChartPicture(PowerPointChart chart,
            out byte[] pngBytes, out string? reason) {
            if (chart == null) throw new ArgumentNullException(nameof(chart));
            pngBytes = Array.Empty<byte>();
            if (!chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot)) {
                reason = "The chart data or chart kind cannot be projected through the shared OfficeIMO chart renderer.";
                return false;
            }
            try {
                OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot,
                    useMinimumCanvas: false);
                return TryRasterizeStaticVisual(drawing, "chart",
                    out pngBytes, out reason);
            } catch (Exception exception) when (exception is ArgumentException
                                                or InvalidOperationException
                                                or OverflowException) {
                reason = $"The chart cannot be rendered as a static binary PowerPoint visual: {exception.Message}";
                return false;
            }
        }

        internal static bool TryRenderTablePicture(PowerPointTable table,
            out byte[] pngBytes, out string? reason) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            pngBytes = Array.Empty<byte>();
            if (!PowerPointSlideImageRenderer.TryCreateTableDrawing(table,
                    out OfficeDrawing drawing, out reason)) {
                return false;
            }
            try {
                return TryRasterizeStaticVisual(drawing, "table",
                    out pngBytes, out reason);
            } catch (Exception exception) when (exception is ArgumentException
                                                or InvalidOperationException
                                                or OverflowException) {
                reason = $"The table cannot be rendered as a static binary PowerPoint visual: {exception.Message}";
                return false;
            }
        }

        internal static bool TryRenderSmartArtPicture(
            PowerPointSmartArt smartArt, out byte[] pngBytes,
            out string? reason) {
            if (smartArt == null) throw new ArgumentNullException(nameof(smartArt));
            pngBytes = Array.Empty<byte>();
            if (!smartArt.TryGetOfficeDiagramSnapshot(
                    out OfficeDiagramSnapshot snapshot)) {
                reason = "The SmartArt data model has no readable semantic nodes for static conversion.";
                return false;
            }
            try {
                OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(snapshot);
                return TryRasterizeStaticVisual(drawing, "SmartArt diagram",
                    out pngBytes, out reason);
            } catch (Exception exception) when (exception is ArgumentException
                                                or InvalidOperationException
                                                or OverflowException) {
                reason = $"The SmartArt diagram cannot be rendered as a static binary PowerPoint visual: {exception.Message}";
                return false;
            }
        }

        private static bool TryRasterizeStaticVisual(OfficeDrawing drawing,
            string ownerName, out byte[] pngBytes, out string? reason) {
            pngBytes = Array.Empty<byte>();
            double pixelAreaAtTwoX = checked(drawing.Width * drawing.Height
                * 4D);
            double scale = pixelAreaAtTwoX <= 16_000_000D
                ? 2D
                : Math.Sqrt(16_000_000D
                    / (drawing.Width * drawing.Height));
            if (double.IsNaN(scale) || double.IsInfinity(scale)
                || scale <= 0D) {
                reason = $"The {ownerName} dimensions cannot be rasterized safely for binary PowerPoint conversion.";
                return false;
            }
            pngBytes = OfficeDrawingRasterRenderer.ToPng(drawing, scale,
                OfficeColor.White);
            reason = null;
            return true;
        }

        internal static bool TryValidatePictureForWrite(
            PowerPointPicture picture, out string? reason) {
            if (picture == null) throw new ArgumentNullException(nameof(picture));
            if (picture is PowerPointMedia) {
                reason = "Media poster frames are encoded through the binary media catalog, not as standalone pictures.";
                return false;
            }
            if (picture.Element is not P.Picture source
                || source.BlipFill?.Blip == null) {
                reason = "The picture has no embedded DrawingML image reference.";
                return false;
            }
            if (source.BlipFill.Blip.Link?.Value != null) {
                reason = "Linked pictures cannot be converted to an embedded binary picture without losing the link.";
                return false;
            }
            if (!TryReadPictureEffects(source.BlipFill.Blip,
                    out _, out reason)) {
                return false;
            }
            if (!TryReadPictureProtectionForWrite(picture, out _,
                    out reason)) {
                return false;
            }
            if (source.BlipFill.GetFirstChild<A.Tile>() != null) {
                reason = "Tiled picture fills cannot be represented by a binary picture frame without changing the visual result.";
                return false;
            }
            A.FillRectangle? fillRectangle = source.BlipFill
                .GetFirstChild<A.Stretch>()?.GetFirstChild<A.FillRectangle>();
            if (fillRectangle is { HasAttributes: true }
                || fillRectangle is { HasChildren: true }) {
                reason = "A picture with a custom stretch rectangle cannot be represented losslessly by the native binary writer.";
                return false;
            }
            reason = null;
            return true;
        }

        private static bool TryReadPicture(PowerPointPicture picture,
            out byte[] imageBytes, out string? contentType,
            out string? reason) {
            imageBytes = Array.Empty<byte>();
            contentType = null;
            if (!TryValidatePictureForWrite(picture, out reason)) return false;
            contentType = NormalizePictureContentType(picture.ContentType);
            if (contentType == null) {
                reason = $"The embedded picture content type '{picture.ContentType ?? "(missing)"}' has no native raster BLIP mapping.";
                return false;
            }
            try {
                imageBytes = picture.GetImageBytes();
                _ = OfficeArtBlipStoreEntryWriter.CreateBlipRecord(imageBytes,
                    contentType);
            } catch (Exception exception) when (exception is InvalidOperationException
                                                or IOException
                                                or NotSupportedException
                                                or ArgumentException
                                                or OverflowException) {
                imageBytes = Array.Empty<byte>();
                reason = $"The embedded {contentType} picture cannot be written as an OfficeArt BLIP: {exception.Message}";
                return false;
            }
            reason = null;
            return true;
        }

        private static string? NormalizePictureContentType(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            return value!.Trim().ToLowerInvariant() switch {
                "image/png" or "image/x-png" => "image/png",
                "image/jpeg" or "image/jpg" => "image/jpeg",
                "image/bmp" or "image/x-ms-bmp" => "image/bmp",
                "image/tiff" or "image/tif" => "image/tiff",
                "image/x-emf" or "image/emf" => "image/x-emf",
                "image/x-wmf" or "image/wmf" => "image/x-wmf",
                _ => null
            };
        }

        private static byte[] BuildPictureFoptRecord(PowerPointPicture picture,
            uint oneBasedStoreIndex) {
            var properties = new List<LegacyPptWriterFoptProperty>(16);
            AddShapeFormattingProperties(properties, picture);
            AddPictureFormatProperties(properties, picture);
            properties.Add(new LegacyPptWriterFoptProperty(0x4104,
                oneBasedStoreIndex));
            return BuildFoptRecord(properties);
        }

        internal static byte[] BuildPreservedPictureFoptRecord(
            LegacyPptRecord prototype, PowerPointPicture picture) {
            if (prototype == null) throw new ArgumentNullException(
                nameof(prototype));
            if (picture == null) throw new ArgumentNullException(
                nameof(picture));
            return BuildPreservedShapeFoptRecord(prototype, picture,
                rewriteShapeTransform: false,
                rewriteShapeGeometry: false,
                rewriteShapeVisualStyle: false,
                rewritePictureFormatting: true)
                ?? throw new InvalidOperationException(
                    "A picture FOPT cannot be empty because it owns the BLIP reference.");
        }

        private static byte[]? BuildPictureTertiaryFoptRecord(
            PowerPointPicture picture) =>
            BuildPreservedPictureTertiaryFoptRecord(null, picture);

        internal static byte[]? BuildPreservedPictureTertiaryFoptRecord(
            LegacyPptRecord? prototype, PowerPointPicture picture) =>
            BuildPreservedTertiaryFoptRecord(prototype, picture,
                shapeVisibility: null);

        internal static byte[]? BuildPreservedTertiaryFoptRecord(
            LegacyPptRecord? prototype,
            PowerPointPicture? pictureFormatting,
            PowerPointShape? shapeVisibility) {
            if (pictureFormatting == null && shapeVisibility == null) {
                throw new ArgumentException(
                    "At least one tertiary shape-property family must be rewritten.");
            }
            var properties = prototype == null
                ? new List<LegacyPptWriterFoptProperty>()
                : ReadFoptProperties(prototype).ToList();
            if (pictureFormatting != null) {
                properties = properties.Where(property =>
                        property.PropertyId != 0x011A)
                    .ToList();
                P.Picture source = (P.Picture)pictureFormatting.Element;
                if (!TryReadPictureEffects(source.BlipFill!.Blip!,
                        out LegacyPptWriterPictureEffects effects,
                        out string? reason)) {
                    throw new NotSupportedException(reason);
                }
                if (effects.RecolorColor.HasValue) {
                    properties.Add(new LegacyPptWriterFoptProperty(0x011A,
                        PackOfficeArtColor(effects.RecolorColor.Value)));
                }
            }
            if (shapeVisibility != null) {
                IReadOnlyList<LegacyPptWriterFoptProperty> sourceProperties =
                    properties.ToArray();
                properties = properties.Where(property =>
                        property.PropertyId != 0x03BF)
                    .ToList();
                if (!TryReadShapeVisibilityForWrite(shapeVisibility,
                        out IReadOnlyList<LegacyPptWriterFoptProperty>
                            visibility, out string? reason)) {
                    throw new NotSupportedException(reason);
                }
                properties.AddRange(visibility);
                const uint hiddenMask = (1U << 14) | (1U << 30);
                PreserveBooleanPropertyBits(sourceProperties, properties,
                    0x03BF, hiddenMask);
            }
            return properties.Count == 0
                ? null
                : BuildFoptRecord(properties, OfficeArtTertiaryFopt);
        }

        private static void AddPictureFormatProperties(
            ICollection<LegacyPptWriterFoptProperty> properties,
            PowerPointPicture picture,
            uint preservedBooleanProperties = 0) {
            P.Picture source = (P.Picture)picture.Element;
            AddPictureProtectionProperties(properties, picture);
            A.SourceRectangle? crop = source.BlipFill?.SourceRectangle;
            AddPictureCropProperty(properties, 0x0100,
                crop?.Top?.Value);
            AddPictureCropProperty(properties, 0x0101,
                crop?.Bottom?.Value);
            AddPictureCropProperty(properties, 0x0102,
                crop?.Left?.Value);
            AddPictureCropProperty(properties, 0x0103,
                crop?.Right?.Value);
            if (!TryReadPictureEffects(source.BlipFill!.Blip!,
                    out LegacyPptWriterPictureEffects effects,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            AddPictureEffectProperties(properties, effects,
                preservedBooleanProperties);
        }

        private static byte[] BuildStaticVisualFoptRecord(
            PowerPointShape shape, uint oneBasedStoreIndex) {
            var properties = new List<LegacyPptWriterFoptProperty>(8);
            AddShapeFormattingProperties(properties, shape);
            properties.Add(new LegacyPptWriterFoptProperty(0x4104,
                oneBasedStoreIndex));
            return BuildFoptRecord(properties);
        }

        private static byte[] BuildOlePreviewFoptRecord(
            PowerPointOleObject shape, PowerPointPicture preview,
            uint oneBasedStoreIndex) {
            var properties = new List<LegacyPptWriterFoptProperty>(16);
            AddShapeTransformProperties(properties, shape);
            AddShapeVisualStyleProperties(properties, preview);
            if (!TryReadShapeMetadataForWrite(shape,
                    out IReadOnlyList<LegacyPptWriterFoptProperty> metadata,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            properties.AddRange(metadata);
            if (!TryReadShapeVisibilityForWrite(shape,
                    out IReadOnlyList<LegacyPptWriterFoptProperty> visibility,
                    out reason)) {
                throw new NotSupportedException(reason);
            }
            properties.AddRange(visibility);
            AddPictureFormatProperties(properties, preview);
            properties.Add(new LegacyPptWriterFoptProperty(0x4104,
                oneBasedStoreIndex));
            return BuildFoptRecord(properties);
        }

        private static void AddPictureCropProperty(
            ICollection<LegacyPptWriterFoptProperty> properties,
            ushort propertyId, int? openXmlRatio) {
            if (!openXmlRatio.HasValue || openXmlRatio.Value == 0) return;
            double scaled = openXmlRatio.Value / 100000D * 65536D;
            if (scaled < int.MinValue || scaled > int.MaxValue) {
                throw new NotSupportedException(
                    "The picture crop value exceeds the signed OfficeArt 16.16 range.");
            }
            int value = checked((int)Math.Round(scaled,
                MidpointRounding.AwayFromZero));
            properties.Add(new LegacyPptWriterFoptProperty(propertyId,
                unchecked((uint)value)));
        }

        internal sealed class LegacyPptWriterPictureCatalog {
            private const int MaximumStoreEntryCount = 0x0FFF;
            private readonly Dictionary<OpenXmlElement, LegacyPptWriterPicture>
                _pictures = new(ReferenceComparer.Instance);
            private readonly Dictionary<string, List<LegacyPptWriterPicture>>
                _entriesByHash = new(StringComparer.Ordinal);
            private readonly List<LegacyPptWriterPicture> _entries = new();
            private readonly int _baseStoreEntryCount;
            private readonly uint _baseDelayedStreamOffset;

            internal LegacyPptWriterPictureCatalog(int baseStoreEntryCount = 0,
                uint baseDelayedStreamOffset = 0) {
                if (baseStoreEntryCount < 0
                    || baseStoreEntryCount > MaximumStoreEntryCount) {
                    throw new ArgumentOutOfRangeException(
                        nameof(baseStoreEntryCount));
                }
                _baseStoreEntryCount = baseStoreEntryCount;
                _baseDelayedStreamOffset = baseDelayedStreamOffset;
            }

            internal IReadOnlyList<LegacyPptWriterPicture> Entries => _entries;

            internal LegacyPptWriterPicture Get(PowerPointShape shape) =>
                Get(shape.Element);

            internal LegacyPptWriterPicture Get(OpenXmlElement element) =>
                _pictures.TryGetValue(element,
                    out LegacyPptWriterPicture? value)
                    ? value
                    : throw new InvalidOperationException(
                        "The picture shape has no BLIP store catalog entry.");

            internal bool TryAdd(PowerPointShape shape, byte[] imageBytes,
                string contentType, out string? reason) =>
                TryAdd(shape.Element, imageBytes, contentType, out reason);

            internal bool TryAdd(OpenXmlElement element, byte[] imageBytes,
                string contentType, out string? reason) {
                if (_pictures.TryGetValue(element,
                        out LegacyPptWriterPicture? existing)) {
                    existing.AddReference();
                    reason = null;
                    return true;
                }
                string hash = ComputePictureHash(contentType, imageBytes);
                if (!_entriesByHash.TryGetValue(hash,
                        out List<LegacyPptWriterPicture>? candidates)) {
                    candidates = new List<LegacyPptWriterPicture>();
                    _entriesByHash.Add(hash, candidates);
                }
                LegacyPptWriterPicture? entry = candidates.FirstOrDefault(
                    candidate => candidate.ContentType == contentType
                        && candidate.ImageBytes.SequenceEqual(imageBytes));
                if (entry == null) {
                    if (_baseStoreEntryCount + _entries.Count
                        >= MaximumStoreEntryCount) {
                        reason = "A binary PowerPoint drawing group cannot contain more than 4,095 distinct picture-store entries.";
                        return false;
                    }
                    entry = new LegacyPptWriterPicture(
                        checked((uint)(_baseStoreEntryCount
                            + _entries.Count) + 1U), imageBytes,
                        contentType);
                    _entries.Add(entry);
                    candidates.Add(entry);
                } else {
                    entry.AddReference();
                }
                _pictures.Add(element, entry);
                reason = null;
                return true;
            }

            internal byte[] BuildStore() {
                if (_baseStoreEntryCount != 0
                    || _baseDelayedStreamOffset != 0) {
                    throw new InvalidOperationException(
                        "A picture catalog based on an existing BLIP store can build only appended FBSE records.");
                }
                byte[][] entries = BuildDelayedStoreEntries();
                return BuildContainer(OfficeArtBStoreContainer,
                    checked((ushort)entries.Length), entries);
            }

            internal byte[][] BuildDelayedStoreEntries() {
                var entries = new byte[_entries.Count][];
                uint delayedStreamOffset = _baseDelayedStreamOffset;
                for (int index = 0; index < _entries.Count; index++) {
                    LegacyPptWriterPicture entry = _entries[index];
                    entries[index] = OfficeArtBlipStoreEntryWriter.CreateDelayed(
                        entry.ImageBytes, entry.ContentType,
                        delayedStreamOffset, entry.ReferenceCount);
                    delayedStreamOffset = checked(delayedStreamOffset
                        + (uint)entry.BlipRecord.Length);
                }
                return entries;
            }

            internal byte[] BuildPicturesStream() {
                int length = _entries.Sum(entry => entry.BlipRecord.Length);
                var result = new byte[length];
                int offset = 0;
                foreach (LegacyPptWriterPicture entry in _entries) {
                    Buffer.BlockCopy(entry.BlipRecord, 0, result, offset,
                        entry.BlipRecord.Length);
                    offset = checked(offset + entry.BlipRecord.Length);
                }
                return result;
            }

            private static string ComputePictureHash(string contentType,
                byte[] imageBytes) {
                using SHA256 sha256 = SHA256.Create();
                return contentType + ":"
                    + Convert.ToBase64String(sha256.ComputeHash(imageBytes));
            }
        }

        internal sealed class LegacyPptWriterPicture {
            private readonly byte[] _imageBytes;

            internal LegacyPptWriterPicture(uint oneBasedStoreIndex,
                byte[] imageBytes, string contentType) {
                OneBasedStoreIndex = oneBasedStoreIndex;
                _imageBytes = (byte[])imageBytes.Clone();
                ContentType = contentType;
                BlipRecord = OfficeArtBlipStoreEntryWriter.CreateBlipRecord(
                    _imageBytes, contentType);
                ReferenceCount = 1;
            }

            internal uint OneBasedStoreIndex { get; }
            internal byte[] ImageBytes => _imageBytes;
            internal string ContentType { get; }
            internal byte[] BlipRecord { get; }
            internal uint ReferenceCount { get; private set; }

            internal void AddReference() => ReferenceCount = checked(
                ReferenceCount + 1U);
        }
    }
}
