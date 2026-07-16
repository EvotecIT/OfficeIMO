using System.Security.Cryptography;
using DocumentFormat.OpenXml;
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
            IEnumerable<PowerPointSlide> slides,
            out LegacyPptWriterPictureCatalog catalog,
            out LegacyPptFeature failureFeature,
            out string? reason) {
            if (slides == null) throw new ArgumentNullException(nameof(slides));
            catalog = new LegacyPptWriterPictureCatalog();
            foreach (PowerPointSlide slide in slides) {
                IReadOnlyList<PowerPointShape> shapes = ReadSlideShapesForWrite(
                    slide, out string? shapeReason);
                if (shapeReason != null) {
                    failureFeature = LegacyPptFeature.RasterPictures;
                    reason = shapeReason;
                    return false;
                }
                foreach (PowerPointShape shape in shapes) {
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
                    } else if (shape is PowerPointTable table) {
                        failureFeature = LegacyPptFeature.Tables;
                        if (!TryRenderTablePicture(table, out imageBytes,
                                out reason)) {
                            return false;
                        }
                        contentType = "image/png";
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
                    } else {
                        continue;
                    }
                    if (!catalog.TryAdd(shape, imageBytes, contentType,
                            out reason)) {
                        return false;
                    }
                }
            }
            failureFeature = LegacyPptFeature.RasterPictures;
            reason = null;
            return true;
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
            var properties = new List<LegacyPptWriterFoptProperty>(8);
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
            IReadOnlyList<LegacyPptWriterFoptProperty> sourceProperties =
                ReadFoptProperties(prototype);
            const uint rewrittenBooleanMask = (1U << 18) | (1U << 17)
                | (1U << 2) | (1U << 1);
            LegacyPptWriterFoptProperty? booleanProperty = sourceProperties
                .LastOrDefault(property => property.PropertyId == 0x013F);
            uint preservedBooleanProperties = (booleanProperty?.Value ?? 0U)
                & ~rewrittenBooleanMask;
            List<LegacyPptWriterFoptProperty> properties = sourceProperties
                .Where(property =>
                    property.PropertyId is not (>= 0x0100 and <= 0x0103)
                    and not 0x0107 and not 0x0108 and not 0x0109
                    and not 0x013F)
                .ToList();
            AddPictureFormatProperties(properties, picture,
                preservedBooleanProperties);
            return BuildFoptRecord(properties);
        }

        private static byte[]? BuildPictureTertiaryFoptRecord(
            PowerPointPicture picture) =>
            BuildPreservedPictureTertiaryFoptRecord(null, picture);

        internal static byte[]? BuildPreservedPictureTertiaryFoptRecord(
            LegacyPptRecord? prototype, PowerPointPicture picture) {
            if (picture == null) throw new ArgumentNullException(
                nameof(picture));
            var properties = prototype == null
                ? new List<LegacyPptWriterFoptProperty>()
                : ReadFoptProperties(prototype).Where(property =>
                    property.PropertyId != 0x011A).ToList();
            P.Picture source = (P.Picture)picture.Element;
            if (!TryReadPictureEffects(source.BlipFill!.Blip!,
                    out LegacyPptWriterPictureEffects effects,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            if (effects.RecolorColor.HasValue) {
                properties.Add(new LegacyPptWriterFoptProperty(0x011A,
                    PackOfficeArtColor(effects.RecolorColor.Value)));
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
            uint oneBasedStoreIndex) => BuildFoptRecord(new[] {
                new LegacyPptWriterFoptProperty(0x4104,
                    oneBasedStoreIndex)
            });

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
            private readonly Dictionary<OpenXmlElement, LegacyPptWriterPicture>
                _pictures = new(ReferenceComparer.Instance);
            private readonly Dictionary<string, List<LegacyPptWriterPicture>>
                _entriesByHash = new(StringComparer.Ordinal);
            private readonly List<LegacyPptWriterPicture> _entries = new();

            internal IReadOnlyList<LegacyPptWriterPicture> Entries => _entries;

            internal LegacyPptWriterPicture Get(PowerPointShape shape) =>
                _pictures.TryGetValue(shape.Element,
                    out LegacyPptWriterPicture? value)
                    ? value
                    : throw new InvalidOperationException(
                        "The picture shape has no BLIP store catalog entry.");

            internal bool TryAdd(PowerPointShape shape, byte[] imageBytes,
                string contentType, out string? reason) {
                string hash = ComputePictureHash(contentType, imageBytes);
                if (!_entriesByHash.TryGetValue(hash,
                        out List<LegacyPptWriterPicture>? candidates)) {
                    candidates = new List<LegacyPptWriterPicture>();
                    _entriesByHash.Add(hash, candidates);
                }
                LegacyPptWriterPicture? entry = candidates.FirstOrDefault(
                    candidate => candidate.ContentType == contentType
                        && candidate.ImageBytes.AsSpan().SequenceEqual(imageBytes));
                if (entry == null) {
                    if (_entries.Count >= ushort.MaxValue) {
                        reason = "A binary PowerPoint drawing group cannot contain more than 65,535 distinct picture-store entries.";
                        return false;
                    }
                    entry = new LegacyPptWriterPicture(
                        checked((uint)_entries.Count + 1U), imageBytes,
                        contentType);
                    _entries.Add(entry);
                    candidates.Add(entry);
                } else {
                    entry.AddReference();
                }
                _pictures.Add(shape.Element, entry);
                reason = null;
                return true;
            }

            internal byte[] BuildStore() {
                var entries = new byte[_entries.Count][];
                uint delayedStreamOffset = 0;
                for (int index = 0; index < _entries.Count; index++) {
                    LegacyPptWriterPicture entry = _entries[index];
                    entries[index] = OfficeArtBlipStoreEntryWriter.CreateDelayed(
                        entry.ImageBytes, entry.ContentType,
                        delayedStreamOffset, entry.ReferenceCount);
                    delayedStreamOffset = checked(delayedStreamOffset
                        + (uint)entry.BlipRecord.Length);
                }
                return BuildContainer(OfficeArtBStoreContainer,
                    checked((ushort)entries.Length), entries);
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
