using System.Security.Cryptography;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using static OfficeIMO.PowerPoint.LegacyPpt.Write.LegacyPptWriterXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort OfficeArtBStoreContainer = 0xF001;

        internal static bool TryReadPictureCatalog(
            LegacyPptPackageView view,
            PowerPointDocument presentation,
            LegacyPptWriterFontCatalog tableFonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            bool convertUnsupportedTables,
            out LegacyPptWriterPictureCatalog catalog,
            out LegacyPptFeature failureFeature,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(view);
            ArgumentNullException.ThrowIfNull(presentation);
            ArgumentNullException.ThrowIfNull(tableFonts);
            ArgumentNullException.ThrowIfNull(pictureBullets);
            catalog = new LegacyPptWriterPictureCatalog();
            var materializedLayoutPictures = new HashSet<string>(
                StringComparer.Ordinal);
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
                LegacyPptPartView slidePart = view.Part(slide.PartRef.Name);
                foreach (PowerPointShape shape in shapes) {
                    if (shape is PowerPointImage
                        && IsLayoutShape(shape)) {
                        materializedLayoutPictures.Add(ElementPath(
                            shape.OwnerPartRef.Name, shape.XmlElement));
                    }
                    byte[] imageBytes;
                    string contentType;
                    if (shape is PowerPointImage picture
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
                    } else if (convertUnsupportedTables
                        && shape is PowerPointTable table
                        && !TryReadTableForWrite(table, tableFonts,
                            pictureBullets, out _)) {
                        failureFeature = LegacyPptFeature.Tables;
                        if (!TryRenderTablePicture(table, out imageBytes,
                                out reason)) {
                            return false;
                        }
                        contentType = "image/png";
                    } else if (shape is PowerPointOleObject ole) {
                        failureFeature = LegacyPptFeature.EmbeddedOle;
                        if (!TryReadOlePreview(view, slidePart, ole,
                                out LegacyPptWriterOlePreview? preview,
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
                LegacyPptPartView notesPart = slidePart.NotesSlide;
                if (!notesPart.IsNone
                    && ShouldWriteNotesPage(slide, out _)
                    && (!SdkTryReadBackground(presentation, notesPart,
                            out LegacyPptWriterBackground? notesBackground,
                            out reason)
                        || !TryAddBackgroundPicture(catalog, notesBackground,
                            out reason))) {
                    failureFeature = LegacyPptFeature.Backgrounds;
                    return false;
                }
            }

            if (!TryValidateMaterializedLayoutPictures(presentation,
                    materializedLayoutPictures, out reason)) {
                failureFeature = LegacyPptFeature.Layouts;
                return false;
            }
            foreach (LegacyPptPartView masterPart in view.SlideMasterParts) {
                if (!TryAddMasterPartPictures(presentation, catalog,
                        masterPart, tableFonts, pictureBullets,
                        convertUnsupportedTables, out failureFeature,
                        out reason)) {
                    return false;
                }
            }
            if (!view.NotesMaster.IsNone
                && !TryAddMasterPartPictures(presentation, catalog,
                    view.NotesMaster, tableFonts, pictureBullets,
                    convertUnsupportedTables, out failureFeature,
                    out reason)) {
                return false;
            }
            if (!view.HandoutMaster.IsNone
                && !TryAddMasterPartPictures(presentation, catalog,
                    view.HandoutMaster, tableFonts, pictureBullets,
                    convertUnsupportedTables, out failureFeature,
                    out reason)) {
                return false;
            }
            failureFeature = LegacyPptFeature.RasterPictures;
            reason = null;
            return true;
        }

        private static bool TryAddMasterPartPictures(
            PowerPointDocument presentation,
            LegacyPptWriterPictureCatalog catalog,
            LegacyPptPartView masterPart,
            LegacyPptWriterFontCatalog tableFonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            bool convertUnsupportedTables,
            out LegacyPptFeature failureFeature, out string? reason) {
            IReadOnlyList<PowerPointShape> masterShapes =
                SdkReadMasterShapesForWrite(presentation, masterPart,
                    out string? masterShapeReason);
            if (masterShapeReason != null) {
                failureFeature = LegacyPptFeature.Masters;
                reason = masterShapeReason;
                return false;
            }
            if (!TryAddMasterPictures(catalog, masterShapes,
                    tableFonts, pictureBullets, convertUnsupportedTables,
                    out failureFeature, out reason)) {
                return false;
            }
            if (!SdkTryReadBackground(presentation, masterPart,
                    out LegacyPptWriterBackground? masterBackground,
                    out reason)
                || !TryAddBackgroundPicture(catalog, masterBackground,
                    out reason)) {
                failureFeature = LegacyPptFeature.Backgrounds;
                return false;
            }
            return true;
        }

        private static bool TryAddBackgroundPicture(
            LegacyPptWriterPictureCatalog catalog,
            LegacyPptWriterBackground? background, out string? reason) {
            OfficeXmlNode pictureFill = background == null
                ? default
                : SdkBackgroundPictureFill(background);
            if (pictureFill.IsNone) {
                reason = null;
                return true;
            }
            return catalog.TryAdd(pictureFill,
                background!.PictureBytes, background.PictureContentType!,
                out reason);
        }

        internal static bool TryRenderChartPicture(PowerPointChart chart,
            out byte[] pngBytes, out string? reason) {
            ArgumentNullException.ThrowIfNull(chart);
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
            ArgumentNullException.ThrowIfNull(table);
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
            ArgumentNullException.ThrowIfNull(smartArt);
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
            try {
                var options = new OfficeImageExportOptions {
                    Scale = 2D,
                    MaximumRasterPixels = MaximumStaticVisualRasterPixels,
                    RasterOverflowBehavior = OfficeRasterOverflowBehavior.ReduceScale
                };
                OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
                    drawing.Width, drawing.Height,
                    OfficeImageExportFormat.Png, options,
                    MaximumStaticVisualRasterPixels, ownerName);
                pngBytes = OfficeDrawingRasterRenderer.ToPng(drawing,
                    plan.Limit.Scale, OfficeColor.White);
                reason = null;
                return true;
            } catch (Exception exception) when (exception
                is ArgumentException or InvalidOperationException
                    or OverflowException or OfficeImageExportLimitException) {
                reason = $"The {ownerName} dimensions cannot be rasterized safely for binary PowerPoint conversion.";
                return false;
            }
        }

        internal static bool TryValidatePictureForWrite(
            PowerPointImage picture, out string? reason) {
            ArgumentNullException.ThrowIfNull(picture);
            if (picture is PowerPointMedia) {
                reason = "Media poster frames are encoded through the binary media catalog, not as standalone pictures.";
                return false;
            }
            return TryValidatePictureForWrite(ShapeXml(picture), out reason);
        }

        private static bool TryValidatePictureForWrite(OfficeXmlNode source,
            out string? reason) {
            OfficeXmlNode blipFill = source.Child(LegacyPptXml.P, "blipFill");
            OfficeXmlNode blip = blipFill.Child(LegacyPptXml.A, "blip");
            if (!source.Is(LegacyPptXml.P, "pic") || blip.IsNone) {
                reason = "The picture has no embedded DrawingML image reference.";
                return false;
            }
            if (blip.Attribute(LegacyPptXml.R, "link") != null) {
                reason = "Linked pictures cannot be converted to an embedded binary picture without losing the link.";
                return false;
            }
            if (!TryReadPictureEffects(blip, out _, out reason)) {
                return false;
            }
            if (!TryReadPictureProtectionForWrite(source, out _,
                    out reason)) {
                return false;
            }
            if (!blipFill.Child(LegacyPptXml.A, "tile").IsNone) {
                reason = "Tiled picture fills cannot be represented by a binary picture frame without changing the visual result.";
                return false;
            }
            OfficeXmlNode fillRectangle = blipFill
                .Child(LegacyPptXml.A, "stretch")
                .Child(LegacyPptXml.A, "fillRect");
            if (!fillRectangle.IsNone && (HasAttributes(fillRectangle)
                    || HasChildren(fillRectangle))) {
                reason = "A picture with a custom stretch rectangle cannot be represented losslessly by the native binary writer.";
                return false;
            }
            reason = null;
            return true;
        }

        private static bool TryReadPicture(PowerPointImage picture,
            out byte[] imageBytes, out string? contentType,
            out string? reason) {
            imageBytes = Array.Empty<byte>();
            contentType = null;
            if (!TryValidatePictureForWrite(picture, out reason)) return false;
            return TryReadPictureBytes(picture.ContentType,
                () => picture.GetImageBytes(MaximumPictureBytes),
                out imageBytes, out contentType, out reason);
        }

        private static bool TryReadPicture(LegacyPptWriterOlePreview preview,
            out byte[] imageBytes, out string? contentType,
            out string? reason) {
            imageBytes = Array.Empty<byte>();
            contentType = null;
            if (!TryValidatePictureForWrite(preview.Element, out reason)) {
                return false;
            }
            LegacyPptPartView imagePart = preview.ImagePart;
            return TryReadPictureBytes(imagePart.IsNone ? null : imagePart.ContentType,
                () => imagePart.IsNone
                    ? throw new InvalidOperationException("Picture has no embedded image part.")
                    : OfficeStreamReader.ReadAllBytes(
                        new MemoryStream(imagePart.GetBytes(), writable: false),
                        MaximumPictureBytes),
                out imageBytes, out contentType, out reason);
        }

        private static bool TryReadPictureBytes(string? sourceContentType,
            Func<byte[]> readBytes, out byte[] imageBytes,
            out string? contentType, out string? reason) {
            imageBytes = Array.Empty<byte>();
            contentType = NormalizePictureContentType(sourceContentType);
            if (contentType == null) {
                reason = $"The embedded picture content type '{sourceContentType ?? "(missing)"}' has no native raster BLIP mapping.";
                return false;
            }
            try {
                imageBytes = readBytes();
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

        private static byte[] BuildPictureFoptRecord(PowerPointImage picture,
            uint oneBasedStoreIndex) {
            var properties = new List<LegacyPptWriterFoptProperty>(16);
            AddShapeFormattingProperties(properties, picture);
            AddPictureFormatProperties(properties, ShapeXml(picture));
            properties.Add(new LegacyPptWriterFoptProperty(0x4104,
                oneBasedStoreIndex));
            return BuildFoptRecord(properties);
        }

        internal static byte[] BuildPreservedPictureFoptRecord(
            LegacyPptRecord prototype, PowerPointImage picture) {
            ArgumentNullException.ThrowIfNull(prototype);
            ArgumentNullException.ThrowIfNull(picture);
            return BuildPreservedShapeFoptRecord(prototype, picture,
                rewriteShapeTransform: false,
                rewriteShapeGeometry: false,
                rewriteShapeVisualStyle: false,
                rewritePictureFormatting: true)
                ?? throw new InvalidOperationException(
                    "A picture FOPT cannot be empty because it owns the BLIP reference.");
        }

        private static byte[]? BuildPictureTertiaryFoptRecord(
            PowerPointImage picture) =>
            BuildPreservedPictureTertiaryFoptRecord(null, picture);

        private static byte[]? BuildPictureTertiaryFoptRecord(
            OfficeXmlNode picture) {
            var properties = new List<LegacyPptWriterFoptProperty>();
            AddPictureRecolorProperty(properties, picture);
            return properties.Count == 0
                ? null
                : BuildFoptRecord(properties, OfficeArtTertiaryFopt);
        }

        private static void AddPictureRecolorProperty(
            ICollection<LegacyPptWriterFoptProperty> properties,
            OfficeXmlNode picture) {
            if (!TryReadPictureEffects(picture
                        .Child(LegacyPptXml.P, "blipFill")
                        .Child(LegacyPptXml.A, "blip"),
                    out LegacyPptWriterPictureEffects effects,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            if (effects.RecolorColor.HasValue) {
                properties.Add(new LegacyPptWriterFoptProperty(0x011A,
                    PackOfficeArtColor(effects.RecolorColor.Value)));
            }
        }

        internal static byte[]? BuildPreservedPictureTertiaryFoptRecord(
            LegacyPptRecord? prototype, PowerPointImage picture) =>
            BuildPreservedTertiaryFoptRecord(prototype, picture,
                shapeVisibility: null);

        internal static byte[]? BuildPreservedTertiaryFoptRecord(
            LegacyPptRecord? prototype,
            PowerPointImage? pictureFormatting,
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
                AddPictureRecolorProperty(properties,
                    ShapeXml(pictureFormatting));
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
            OfficeXmlNode source,
            uint preservedBooleanProperties = 0) {
            AddPictureProtectionProperties(properties, source);
            OfficeXmlNode blipFill = source.Child(LegacyPptXml.P, "blipFill");
            OfficeXmlNode crop = blipFill.Child(LegacyPptXml.A, "srcRect");
            AddPictureCropProperty(properties, 0x0100, crop.Int("t"));
            AddPictureCropProperty(properties, 0x0101, crop.Int("b"));
            AddPictureCropProperty(properties, 0x0102, crop.Int("l"));
            AddPictureCropProperty(properties, 0x0103, crop.Int("r"));
            if (!TryReadPictureEffects(blipFill.Child(LegacyPptXml.A, "blip"),
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
            PowerPointOleObject shape, LegacyPptWriterOlePreview preview,
            uint oneBasedStoreIndex) {
            var properties = new List<LegacyPptWriterFoptProperty>(16);
            AddShapeTransformProperties(properties, shape);
            if (!TryReadShapeVisualStyle(preview.Element,
                    out IReadOnlyList<LegacyPptWriterFoptProperty> style,
                    out string? styleReason)) {
                throw new NotSupportedException(styleReason);
            }
            foreach (LegacyPptWriterFoptProperty property in style) {
                properties.Add(property);
            }
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
            AddPictureFormatProperties(properties, preview.Element);
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
            // Keyed by the shape's write key or by the background fill element.
            private readonly Dictionary<object, LegacyPptWriterPicture>
                _pictures = new();
            private readonly Dictionary<string, List<LegacyPptWriterPicture>>
                _entriesByHash = new(StringComparer.Ordinal);
            private readonly List<LegacyPptWriterPicture> _entries = new();
            private readonly int _baseStoreEntryCount;
            private readonly uint _baseDelayedStreamOffset;
            private long _totalPictureBytes;

            internal LegacyPptWriterPictureCatalog(int baseStoreEntryCount = 0,
                uint baseDelayedStreamOffset = 0) {
                ArgumentOutOfRangeException.ThrowIfNegative(baseStoreEntryCount);
                ArgumentOutOfRangeException.ThrowIfGreaterThan(baseStoreEntryCount, MaximumStoreEntryCount);
                _baseStoreEntryCount = baseStoreEntryCount;
                _baseDelayedStreamOffset = baseDelayedStreamOffset;
            }

            internal IReadOnlyList<LegacyPptWriterPicture> Entries => _entries;

            internal LegacyPptWriterPicture Get(PowerPointShape shape) =>
                Get(ShapeKey(shape));

            internal bool Contains(PowerPointShape shape) =>
                _pictures.ContainsKey(ShapeKey(shape));

            internal LegacyPptWriterPicture Get(OfficeXmlNode element) =>
                Get((object)element);

            private LegacyPptWriterPicture Get(object key) =>
                _pictures.TryGetValue(key,
                    out LegacyPptWriterPicture? value)
                    ? value
                    : throw new InvalidOperationException(
                        "The picture shape has no BLIP store catalog entry.");

            internal bool TryAdd(PowerPointShape shape, byte[] imageBytes,
                string contentType, out string? reason) =>
                TryAdd(ShapeKey(shape), imageBytes, contentType, out reason);

            internal bool TryAdd(OfficeXmlNode element, byte[] imageBytes,
                string contentType, out string? reason) =>
                TryAdd((object)element, imageBytes, contentType, out reason);

            private bool TryAdd(object element, byte[] imageBytes,
                string contentType, out string? reason) {
                if (_pictures.TryGetValue(element,
                        out LegacyPptWriterPicture? existing)) {
                    existing.AddReference();
                    reason = null;
                    return true;
                }
                if (imageBytes.Length > MaximumPictureBytes) {
                    reason = $"A binary PowerPoint picture cannot exceed {MaximumPictureBytes} bytes.";
                    return false;
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
                    long nextTotal = checked(_totalPictureBytes
                        + imageBytes.Length);
                    if (nextTotal > MaximumTotalPictureBytes) {
                        reason = $"Binary PowerPoint picture payloads cannot exceed {MaximumTotalPictureBytes} aggregate bytes.";
                        return false;
                    }
                    entry = new LegacyPptWriterPicture(
                        checked((uint)(_baseStoreEntryCount
                            + _entries.Count) + 1U), imageBytes,
                        contentType);
                    _entries.Add(entry);
                    candidates.Add(entry);
                    _totalPictureBytes = nextTotal;
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
