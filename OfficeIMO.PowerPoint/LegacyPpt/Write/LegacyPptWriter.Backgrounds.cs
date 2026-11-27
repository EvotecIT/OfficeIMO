using System.Collections.ObjectModel;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Core.Internal;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort OfficeArtFopt = 0xF00B;
        private const ushort OfficeArtTertiaryFopt = 0xF122;
        private const uint OfficeArtBackgroundShapeFlag = 1U << 10;

        // Callers in LegacyPptWritePreflight.cs and LegacyPptWriter.Pictures.cs not yet passing a view.
        internal static bool TryReadBackground(PowerPointSlide slide,
            out LegacyPptWriterBackground? background, out string? reason) {
            ArgumentNullException.ThrowIfNull(slide);
            PowerPointDocument presentation = OwnerDocumentForWrite(slide);
            using LegacyPptPackageView view = LegacyPptPackageView.Capture(presentation);
            return TryReadBackground(view, presentation, slide, out background, out reason);
        }

        /// <summary>The background a slide shows: its own, or else its layout's.</summary>
        internal static bool TryReadBackground(LegacyPptPackageView view,
            PowerPointDocument presentation, PowerPointSlide slide,
            out LegacyPptWriterBackground? background, out string? reason) {
            ArgumentNullException.ThrowIfNull(slide);
            return TryReadSlideBackground(presentation, SlidePartView(view, slide),
                out background, out reason);
        }

        private static bool TryReadSlideBackground(PowerPointDocument presentation,
            LegacyPptPartView slidePart,
            out LegacyPptWriterBackground? background, out string? reason) {
            LegacyPptPartView ownerPart = slidePart;
            OfficeXmlNode source = CommonSlideData(slidePart).Child(Pml, "bg");
            if (source.IsNone) {
                ownerPart = slidePart.SlideLayout;
                source = CommonSlideData(ownerPart).Child(Pml, "bg");
            }
            if (source.IsNone || ownerPart.IsNone) {
                background = null;
                reason = null;
                return true;
            }
            return TryReadBackground(presentation, ownerPart, source, "slide or layout",
                out background, out reason);
        }

        /// <summary>
        /// The background of a slide master, notes master, handout master, title master layout or notes page part; a
        /// slide part reads its layout's background when it has none of its own.
        /// </summary>
        internal static bool TryReadBackground(PowerPointDocument presentation,
            LegacyPptPartView part,
            out LegacyPptWriterBackground? background, out string? reason) {
            ArgumentNullException.ThrowIfNull(presentation);
            if (part.IsNone) throw new ArgumentNullException(nameof(part));
            if (part.IsSlide) return TryReadSlideBackground(presentation, part, out background, out reason);
            string ownerName = part.IsSlideMaster ? "slide master"
                : part.IsNotesMaster ? "notes master"
                : part.IsHandoutMaster ? "handout master"
                : part.IsSlideLayout ? "title master layout"
                : "notes page";
            OfficeXmlNode source = CommonSlideData(part).Child(Pml, "bg");
            if (source.IsNone) {
                background = null;
                reason = null;
                return true;
            }
            return TryReadBackground(presentation, part, source, ownerName,
                out background, out reason);
        }

        private static bool TryReadBackground(PowerPointDocument presentation,
            LegacyPptPartView ownerPart, OfficeXmlNode source, string ownerName,
            out LegacyPptWriterBackground? background, out string? reason) {
            OfficeXmlNode colorScheme = GetBackgroundColorScheme(ownerPart);
            OfficeXmlNode placeholderColor = source.Child(Pml, "bgRef");
            OfficeXmlNode fill = GetBackgroundFill(ownerPart, source,
                out LegacyPptPartView fillOwnerPart);
            if (fill.IsNone) {
                if (!placeholderColor.IsNone) {
                    background = null;
                    reason = $"The {ownerName} background style reference cannot be resolved from its active theme.";
                    return false;
                }
                background = null;
                reason = null;
                return true;
            }
            if (fill.Is(Dml, "grpFill")) {
                background = null;
                reason = null;
                return true;
            }
            if (fill.Is(Dml, "noFill")) {
                background = LegacyPptWriterBackground.NoFill();
                reason = null;
                return true;
            }
            if (fill.Is(Dml, "solidFill")) {
                OfficeColor? color = OfficeXmlThemeColors.ResolveColor(
                    fill, colorScheme, placeholderColor);
                if (!color.HasValue) {
                    background = null;
                    reason = $"The {ownerName} solid background color cannot be resolved for binary PowerPoint writing.";
                    return false;
                }
                background = LegacyPptWriterBackground.Solid(color.Value);
                reason = null;
                return true;
            }
            if (fill.Is(Dml, "gradFill")) {
                return TryReadGradientBackground(fill, colorScheme,
                    placeholderColor, ownerName, out background, out reason);
            }
            if (fill.Is(Dml, "blipFill")) {
                return TryReadPictureBackground(presentation, fillOwnerPart, fill,
                    ownerName, out background, out reason);
            }

            background = null;
            reason = fill.Is(Dml, "pattFill")
                ? $"The {ownerName} pattern background is not yet encoded by the native binary writer."
                : $"The {ownerName} background fill type '{fill.LocalName}' is not supported by the native binary writer.";
            return false;
        }

        private static bool TryReadPictureBackground(PowerPointDocument presentation,
            LegacyPptPartView ownerPart, OfficeXmlNode fill, string ownerName,
            out LegacyPptWriterBackground? background, out string? reason) {
            background = null;
            OfficeXmlNode blip = fill.Child(Dml, "blip");
            string? relationshipId = blip.Attribute(LegacyPptXml.R, "embed");
            if (blip.IsNone || string.IsNullOrWhiteSpace(relationshipId)) {
                reason = $"The {ownerName} picture background has no embedded image relationship.";
                return false;
            }
            if (!string.IsNullOrWhiteSpace(blip.Attribute(LegacyPptXml.R, "link"))) {
                reason = $"The {ownerName} picture background is linked; embedding it in binary PowerPoint would lose the link.";
                return false;
            }
            if (HasChildElements(blip)) {
                reason = $"The {ownerName} picture background uses image effects that have no lossless classic background-fill mapping.";
                return false;
            }
            if (!fill.Child(Dml, "srcRect").IsNone) {
                reason = $"The {ownerName} picture background uses source cropping that is not yet encoded by the native binary writer.";
                return false;
            }
            if (!fill.Child(Dml, "tile").IsNone) {
                reason = $"The {ownerName} picture background uses tiled placement and cannot be converted to one stretched classic picture fill without changing its appearance.";
                return false;
            }
            OfficeXmlNode stretch = fill.Child(Dml, "stretch");
            OfficeXmlNode fillRectangle = stretch.Child(Dml, "fillRect");
            bool otherChildren = false;
            foreach (OfficeXmlNode child in fill.Children()) {
                if (!child.Is(Dml, "blip") && !child.Is(Dml, "stretch")) otherChildren = true;
            }
            if (stretch.IsNone || fillRectangle.IsNone
                || HasAttributes(fillRectangle) || HasChildElements(fillRectangle)
                || otherChildren) {
                reason = $"The {ownerName} picture background does not use the plain full-frame stretch mapping supported by classic binary PowerPoint.";
                return false;
            }

            if (!ownerPart.TryGetRelationship(relationshipId, out LegacyPptRelationshipView relationship)
                || relationship.IsExternal || relationship.TargetPart.IsNone) {
                reason = $"The {ownerName} picture background relationship cannot be resolved.";
                return false;
            }
            if (!relationship.IsType(LegacyPptPackageView.ImageRelationshipType)) {
                reason = $"The {ownerName} picture background relationship does not target an image part.";
                return false;
            }
            LegacyPptPartView imagePart = relationship.TargetPart;
            if (imagePart.Relationships.Count > 0) {
                reason = $"Related parts on the {ownerName} picture background have no binary PowerPoint mapping.";
                return false;
            }
            string? contentType = NormalizePictureContentType(
                imagePart.ContentType);
            if (contentType == null) {
                reason = $"The {ownerName} picture background content type '{imagePart.ContentType}' has no native OfficeArt BLIP mapping.";
                return false;
            }
            byte[] imageBytes;
            try {
                imageBytes = imagePart.ReadBytes(64 * 1024 * 1024);
                _ = OfficeArtBlipStoreEntryWriter.CreateBlipRecord(imageBytes,
                    contentType);
            } catch (Exception exception) when (exception is IOException
                                                or InvalidDataException
                                                or InvalidOperationException
                                                or NotSupportedException
                                                or UnauthorizedAccessException
                                                or ArgumentException
                                                or OverflowException) {
                reason = $"The {ownerName} picture background cannot be written as an OfficeArt BLIP: {exception.Message}";
                return false;
            }
            background = LegacyPptWriterBackground.Picture(presentation, ownerPart.Name,
                fill, imageBytes, contentType);
            reason = null;
            return true;
        }

        private static bool TryReadGradientBackground(OfficeXmlNode gradient,
            OfficeXmlNode colorScheme, OfficeXmlNode placeholderColor,
            string ownerName, out LegacyPptWriterBackground? background,
            out string? reason) {
            OfficeXmlNode linear = gradient.Child(Dml, "lin");
            OfficeXmlNode path = gradient.Child(Dml, "path");
            if (linear.IsNone == path.IsNone) {
                background = null;
                reason = $"The {ownerName} gradient background must contain exactly one linear or supported path geometry.";
                return false;
            }
            string? flip = gradient.Attribute("flip");
            if (flip != null && !string.Equals(flip, "none", StringComparison.Ordinal)) {
                background = null;
                reason = $"The {ownerName} gradient background uses tile flipping that the binary PowerPoint writer cannot reproduce losslessly.";
                return false;
            }
            foreach (OfficeXmlNode child in gradient.Children()) {
                if (child.Is(Dml, "gsLst") || child == linear || child == path) continue;
                background = null;
                reason = $"The {ownerName} gradient background uses a custom tile rectangle that the binary PowerPoint format cannot reproduce through this writer.";
                return false;
            }
            if (!path.IsNone && HasChildElements(path)) {
                background = null;
                reason = $"The {ownerName} path gradient uses a custom fill rectangle that has no lossless binary PowerPoint mapping.";
                return false;
            }
            OfficeXmlNode[] sourceStops = gradient.Child(Dml, "gsLst")
                .Children(Dml, "gs").ToList()
                .OrderBy(stop => stop.Int("pos") ?? 0)
                .ToArray();
            if (sourceStops.Length < 2 || sourceStops.Length > ushort.MaxValue) {
                background = null;
                reason = $"The {ownerName} gradient background needs at least two resolvable stops.";
                return false;
            }

            var stops = new List<LegacyPptWriterGradientStop>(sourceStops.Length);
            foreach (OfficeXmlNode sourceStop in sourceStops) {
                int position = sourceStop.Int("pos") ?? 0;
                OfficeColor? color = OfficeXmlThemeColors.ResolveColor(
                    sourceStop, colorScheme, placeholderColor);
                if (position < 0 || position > 100000 || !color.HasValue) {
                    background = null;
                    reason = $"The {ownerName} gradient background contains an invalid position or unresolved color.";
                    return false;
                }
                stops.Add(new LegacyPptWriterGradientStop(color.Value,
                    position / 100000D));
            }

            if (!TryGetGradientOpacityRamp(stops, out byte foregroundAlpha,
                    out byte backgroundAlpha)) {
                background = null;
                reason = $"The {ownerName} gradient uses non-linear stop opacity; binary PowerPoint can store only one linear opacity ramp across the gradient.";
                return false;
            }

            double legacyAngle = 0D;
            uint fillType;
            string? pathKind = path.Attribute("path");
            if (!linear.IsNone) {
                double openXmlAngle = (linear.Int("ang") ?? 0) / 60000D;
                legacyAngle = NormalizeBackgroundAngle(270D - openXmlAngle);
                fillType = linear.Bool("scaled") == true ? 7U : 4U;
            } else if (pathKind == "circle") {
                fillType = 5U;
            } else if (pathKind == "shape") {
                fillType = 6U;
            } else {
                background = null;
                reason = $"The {ownerName} path gradient geometry '{pathKind}' has no lossless binary PowerPoint mapping.";
                return false;
            }
            background = LegacyPptWriterBackground.Gradient(fillType,
                legacyAngle, foregroundAlpha, backgroundAlpha, stops);
            reason = null;
            return true;
        }

        private static bool TryGetGradientOpacityRamp(
            IReadOnlyList<LegacyPptWriterGradientStop> stops,
            out byte foregroundAlpha, out byte backgroundAlpha) {
            foregroundAlpha = byte.MaxValue;
            backgroundAlpha = byte.MaxValue;
            if (stops.Count == 0) return false;

            LegacyPptWriterGradientStop first = stops[0];
            LegacyPptWriterGradientStop last = stops[stops.Count - 1];
            if (stops.All(stop => stop.Color.A == first.Color.A)) {
                foregroundAlpha = first.Color.A;
                backgroundAlpha = first.Color.A;
                return true;
            }

            double positionRange = last.Position - first.Position;
            if (positionRange <= 0D) return false;
            double slope = (last.Color.A - first.Color.A) / positionRange;
            double foreground = first.Color.A - slope * first.Position;
            double background = foreground + slope;
            if (foreground < -0.5D || foreground > 255.5D
                || background < -0.5D || background > 255.5D) {
                return false;
            }

            byte rampForegroundAlpha = checked((byte)Math.Max(0D, Math.Min(255D,
                Math.Round(foreground, MidpointRounding.ToEven))));
            byte rampBackgroundAlpha = checked((byte)Math.Max(0D, Math.Min(255D,
                Math.Round(background, MidpointRounding.ToEven))));
            bool representable = stops.All(stop => {
                double expected = rampForegroundAlpha
                    + (rampBackgroundAlpha - rampForegroundAlpha) * stop.Position;
                return Math.Abs(stop.Color.A
                    - Math.Round(expected, MidpointRounding.ToEven)) <= 1D;
            });
            if (!representable) return false;
            foregroundAlpha = rampForegroundAlpha;
            backgroundAlpha = rampBackgroundAlpha;
            return true;
        }

        private static OfficeXmlNode GetBackgroundFill(LegacyPptPartView ownerPart,
            OfficeXmlNode source, out LegacyPptPartView fillOwnerPart) {
            fillOwnerPart = ownerPart;
            OfficeXmlNode properties = source.Child(Pml, "bgPr");
            if (!properties.IsNone && HasChildElements(properties)) {
                return properties.FirstChild;
            }
            OfficeXmlNode reference = source.Child(Pml, "bgRef");
            uint? referenceIndex = reference.UInt("idx");
            if (!referenceIndex.HasValue) return default;
            OfficeXmlNode formatScheme = GetBackgroundFormatScheme(
                ownerPart, out LegacyPptPartView themePart);
            if (formatScheme.IsNone || themePart.IsNone) return default;
            fillOwnerPart = themePart;
            uint index = referenceIndex.Value;
            if (index >= 1001U) {
                return NthChild(formatScheme.Child(Dml, "bgFillStyleLst"), index - 1001U);
            }
            if (index < 1U) return default;
            return NthChild(formatScheme.Child(Dml, "fillStyleLst"), index - 1U);
        }

        private static OfficeXmlNode NthChild(OfficeXmlNode parent, uint index) {
            uint position = 0;
            foreach (OfficeXmlNode child in parent.Children()) {
                if (position == index) return child;
                position++;
            }
            return default;
        }

        private static OfficeXmlNode GetBackgroundColorScheme(
            LegacyPptPartView ownerPart) {
            foreach (LegacyPptPartView themePart in GetBackgroundThemeSources(ownerPart)) {
                OfficeXmlNode scheme = GetThemeSchemeElement(themePart, "clrScheme");
                if (!scheme.IsNone) return scheme;
            }
            return default;
        }

        private static OfficeXmlNode GetBackgroundFormatScheme(
            LegacyPptPartView ownerPart, out LegacyPptPartView themeSourcePart) {
            foreach (LegacyPptPartView candidate in
                     GetBackgroundThemeSources(ownerPart)) {
                OfficeXmlNode scheme = GetThemeSchemeElement(candidate, "fmtScheme");
                if (scheme.IsNone) continue;
                themeSourcePart = candidate;
                return scheme;
            }
            themeSourcePart = default;
            return default;
        }

        // A theme override holds its schemes at its root; a theme inside a:themeElements.
        private static OfficeXmlNode GetThemeSchemeElement(LegacyPptPartView themePart,
            string localName) =>
            themePart.IsThemeOverride ? ThemeOverrideElement(themePart, localName)
            : themePart.IsTheme ? ThemeElement(themePart, localName)
            : default;

        private static IEnumerable<LegacyPptPartView> GetBackgroundThemeSources(
            LegacyPptPartView ownerPart) {
            if (ownerPart.IsSlide) {
                if (!ownerPart.ThemeOverride.IsNone) yield return ownerPart.ThemeOverride;
                LegacyPptPartView layout = ownerPart.SlideLayout;
                if (!layout.ThemeOverride.IsNone) yield return layout.ThemeOverride;
                if (!layout.SlideMaster.Theme.IsNone) yield return layout.SlideMaster.Theme;
            } else if (ownerPart.IsSlideLayout) {
                if (!ownerPart.ThemeOverride.IsNone) yield return ownerPart.ThemeOverride;
                if (!ownerPart.SlideMaster.Theme.IsNone) yield return ownerPart.SlideMaster.Theme;
            } else if (ownerPart.IsNotesSlide) {
                if (!ownerPart.ThemeOverride.IsNone) yield return ownerPart.ThemeOverride;
                LegacyPptPartView notesMaster = ownerPart.RelatedPart(
                    LegacyPptPackageView.NotesMasterRelationshipType);
                if (!notesMaster.Theme.IsNone) yield return notesMaster.Theme;
            } else if (ownerPart.IsSlideMaster || ownerPart.IsNotesMaster
                       || ownerPart.IsHandoutMaster) {
                if (!ownerPart.Theme.IsNone) yield return ownerPart.Theme;
            }
        }

        private static byte[] BuildBackgroundDrawingRecord(LegacyPptRecord drawing,
            LegacyPptWriterBackground background,
            LegacyPptWriterPictureCatalog? pictureCatalog = null) {
            var children = new List<byte[]>(drawing.Children.Count);
            bool wroteBackground = false;
            foreach (LegacyPptRecord child in drawing.Children) {
                if (child.Type != OfficeArtDgContainer) {
                    children.Add(child.CopyRecordBytes());
                    continue;
                }
                var drawingChildren = new List<byte[]>(child.Children.Count);
                foreach (LegacyPptRecord drawingChild in child.Children) {
                    if (IsBackgroundShapeRecord(drawingChild)) {
                        drawingChildren.Add(BuildBackgroundShapeRecord(
                            drawingChild, background, pictureCatalog));
                        wroteBackground = true;
                    } else {
                        drawingChildren.Add(drawingChild.CopyRecordBytes());
                    }
                }
                children.Add(BuildContainer(child.Type, child.Instance, drawingChildren));
            }
            if (!wroteBackground) {
                throw new InvalidDataException(
                    "The embedded binary PowerPoint master has no OfficeArt background shape.");
            }
            return BuildContainer(drawing.Type, drawing.Instance, children);
        }

        private static bool IsBackgroundShapeRecord(LegacyPptRecord record) {
            if (record.Type != OfficeArtSpContainer) return false;
            LegacyPptRecord? fsp = record.Children.FirstOrDefault(child =>
                child.Type == OfficeArtFsp);
            return fsp != null && fsp.PayloadLength >= 8
                && (fsp.ReadUInt32(4) & OfficeArtBackgroundShapeFlag) != 0;
        }

        internal static int? ReadBackgroundPictureStoreIndex(
            LegacyPptRecord prototype) {
            ArgumentNullException.ThrowIfNull(prototype);
            LegacyPptRecord? drawing = prototype.Children.FirstOrDefault(child =>
                child.Type == RecordDrawing);
            LegacyPptRecord? backgroundShape = drawing?.DescendantsAndSelf()
                .Where(record => record.Type == OfficeArtSpContainer)
                .LastOrDefault(IsBackgroundShapeRecord);
            LegacyPptRecord? fopt = backgroundShape?.Children.FirstOrDefault(
                child => child.Type == OfficeArtFopt);
            if (fopt == null) return null;
            LegacyPptWriterFoptProperty? property = ReadFoptProperties(fopt)
                .LastOrDefault(candidate => candidate.PropertyId == 0x0186
                    && (candidate.OperationId & 0x4000) != 0);
            return property == null || property.Value == 0
                || property.Value > int.MaxValue
                ? null
                : checked((int)property.Value);
        }

        private static byte[] BuildBackgroundShapeRecord(LegacyPptRecord prototype,
            LegacyPptWriterBackground background,
            LegacyPptWriterPictureCatalog? pictureCatalog = null) {
            var children = new List<byte[]>(prototype.Children.Count + 1);
            bool wroteFopt = false;
            foreach (LegacyPptRecord child in prototype.Children) {
                if (child.Type == OfficeArtFopt) {
                    children.Add(BuildBackgroundFoptRecord(child, background,
                        pictureCatalog));
                    wroteFopt = true;
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (!wroteFopt) {
                children.Insert(Math.Min(1, children.Count),
                    BuildBackgroundFoptRecord(null, background,
                        pictureCatalog));
            }
            return BuildContainer(OfficeArtSpContainer, prototype.Instance, children);
        }

        private static byte[] BuildBackgroundFoptRecord(LegacyPptRecord? prototype,
            LegacyPptWriterBackground background,
            LegacyPptWriterPictureCatalog? pictureCatalog) {
            List<LegacyPptWriterFoptProperty> properties = prototype == null
                ? new List<LegacyPptWriterFoptProperty>()
                : ReadFoptProperties(prototype).Where(property =>
                    property.PropertyId < 0x0180 || property.PropertyId > 0x01BF)
                    .ToList();
            properties.Add(new LegacyPptWriterFoptProperty(0x0180,
                background.FillType));
            if (background.RequiresPictureCatalog) {
                LegacyPptWriterPicture picture = pictureCatalog?.Get(background)
                    ?? throw new NotSupportedException(
                        "A preservation-aware picture-background edit requires an updated binary BLIP store.");
                properties.Add(new LegacyPptWriterFoptProperty(0x4186,
                    picture.OneBasedStoreIndex));
            } else if (background.Filled) {
                LegacyPptWriterGradientStop first = background.Stops[0];
                LegacyPptWriterGradientStop last = background.Stops[background.Stops.Count - 1];
                properties.Add(new LegacyPptWriterFoptProperty(0x0181,
                    PackOfficeArtColor(first.Color)));
                properties.Add(new LegacyPptWriterFoptProperty(0x0182,
                    PackOfficeArtOpacity(background.ForegroundAlpha)));
                properties.Add(new LegacyPptWriterFoptProperty(0x0183,
                    PackOfficeArtColor(last.Color)));
                properties.Add(new LegacyPptWriterFoptProperty(0x0184,
                    PackOfficeArtOpacity(background.BackgroundAlpha)));
                if (background.FillType is >= 4U and <= 8U) {
                    if (background.FillType is 4U or 7U) {
                        properties.Add(new LegacyPptWriterFoptProperty(0x018B,
                            unchecked((uint)checked((int)Math.Round(
                                background.AngleDegrees * 65536D,
                                MidpointRounding.AwayFromZero)))));
                    }
                    properties.Add(new LegacyPptWriterFoptProperty(0x018C, 0));
                    byte[] shadeColors = BuildGradientStopArray(background.Stops);
                    properties.Add(new LegacyPptWriterFoptProperty(0x8197,
                        checked((uint)shadeColors.Length), shadeColors));
                }
            }
            properties.Add(new LegacyPptWriterFoptProperty(0x01BF,
                background.Filled ? 0x00100010U : 0x00100000U));
            return BuildFoptRecord(properties);
        }

        private static IReadOnlyList<LegacyPptWriterFoptProperty> ReadFoptProperties(
            LegacyPptRecord record) {
            int fixedLength = checked(record.Instance * 6);
            if (fixedLength > record.PayloadLength) {
                throw new InvalidDataException(
                    "The embedded binary PowerPoint FOPT table is truncated.");
            }
            int complexOffset = fixedLength;
            var result = new List<LegacyPptWriterFoptProperty>(record.Instance);
            for (int index = 0; index < record.Instance; index++) {
                ushort operationId = record.ReadUInt16(index * 6);
                uint value = record.ReadUInt32(index * 6 + 2);
                byte[]? complexData = null;
                if ((operationId & 0x8000) != 0) {
                    if (value > int.MaxValue
                        || complexOffset > record.PayloadLength - checked((int)value)) {
                        throw new InvalidDataException(
                            "The embedded binary PowerPoint FOPT complex property is truncated.");
                    }
                    complexData = new byte[checked((int)value)];
                    for (int byteIndex = 0; byteIndex < complexData.Length; byteIndex++) {
                        complexData[byteIndex] = record.ReadByte(complexOffset + byteIndex);
                    }
                    complexOffset += complexData.Length;
                }
                result.Add(new LegacyPptWriterFoptProperty(operationId, value,
                    complexData));
            }
            return result;
        }

        private static byte[] BuildFoptRecord(
            IReadOnlyList<LegacyPptWriterFoptProperty> source,
            ushort recordType = OfficeArtFopt) {
            LegacyPptWriterFoptProperty[] properties = source
                .OrderBy(property => property.PropertyId)
                .ToArray();
            int fixedLength = checked(properties.Length * 6);
            int complexLength = properties.Sum(property =>
                property.ComplexData?.Length ?? 0);
            var payload = new byte[checked(fixedLength + complexLength)];
            int complexOffset = fixedLength;
            for (int index = 0; index < properties.Length; index++) {
                LegacyPptWriterFoptProperty property = properties[index];
                WriteUInt16(payload, index * 6, property.OperationId);
                WriteUInt32(payload, index * 6 + 2, property.ComplexData == null
                    ? property.Value
                    : checked((uint)property.ComplexData.Length));
                if (property.ComplexData != null) {
                    Buffer.BlockCopy(property.ComplexData, 0, payload, complexOffset,
                        property.ComplexData.Length);
                    complexOffset += property.ComplexData.Length;
                }
            }
            return BuildRecord(version: 3, checked((ushort)properties.Length),
                recordType, payload);
        }

        private static byte[] BuildGradientStopArray(
            IReadOnlyList<LegacyPptWriterGradientStop> stops) {
            var data = new byte[checked(6 + stops.Count * 8)];
            WriteUInt16(data, 0, checked((ushort)stops.Count));
            WriteUInt16(data, 2, checked((ushort)stops.Count));
            WriteUInt16(data, 4, 8);
            for (int index = 0; index < stops.Count; index++) {
                WriteUInt32(data, 6 + index * 8,
                    PackOfficeArtColor(stops[index].Color));
                WriteUInt32(data, 10 + index * 8, checked((uint)Math.Round(
                    stops[index].Position * 65536D,
                    MidpointRounding.AwayFromZero)));
            }
            return data;
        }

        private static uint PackOfficeArtColor(OfficeColor color) =>
            unchecked((uint)(color.R | color.G << 8 | color.B << 16));

        private static uint PackOfficeArtOpacity(byte alpha) => checked((uint)Math.Round(
            alpha / 255D * 65536D, MidpointRounding.AwayFromZero));

        private static double NormalizeBackgroundAngle(double angle) {
            double normalized = angle % 360D;
            return normalized < 0D ? normalized + 360D : normalized;
        }

        internal sealed partial class LegacyPptWriterBackground {
            private LegacyPptWriterBackground(bool filled, uint fillType,
                double angleDegrees, byte foregroundAlpha, byte backgroundAlpha,
                IReadOnlyList<LegacyPptWriterGradientStop> stops,
                OfficeXmlNode pictureFill = default, byte[]? pictureBytes = null,
                string? pictureContentType = null,
                PowerPointDocument? pictureDocument = null,
                string? picturePartName = null) {
                Filled = filled;
                FillType = fillType;
                AngleDegrees = angleDegrees;
                ForegroundAlpha = foregroundAlpha;
                BackgroundAlpha = backgroundAlpha;
                Stops = new ReadOnlyCollection<LegacyPptWriterGradientStop>(
                    stops.ToArray());
                PictureFillNode = pictureFill;
                PictureDocument = pictureDocument;
                PicturePartName = picturePartName;
                PictureBytes = pictureBytes == null
                    ? Array.Empty<byte>()
                    : (byte[])pictureBytes.Clone();
                PictureContentType = pictureContentType;
            }

            internal bool Filled { get; }
            internal uint FillType { get; }
            internal double AngleDegrees { get; }
            internal byte ForegroundAlpha { get; }
            internal byte BackgroundAlpha { get; }
            internal IReadOnlyList<LegacyPptWriterGradientStop> Stops { get; }
            /// <summary>The picture fill (<c>a:blipFill</c>) in the captured view, or none.</summary>
            internal OfficeXmlNode PictureFillNode { get; }
            /// <summary>The live document and part the picture fill was read from.</summary>
            internal PowerPointDocument? PictureDocument { get; }
            internal string? PicturePartName { get; }
            internal byte[] PictureBytes { get; }
            internal string? PictureContentType { get; }
            internal bool RequiresPictureCatalog => !PictureFillNode.IsNone;

            internal static LegacyPptWriterBackground NoFill() =>
                new(false, 0U, 0D, byte.MaxValue, byte.MaxValue,
                    Array.Empty<LegacyPptWriterGradientStop>());

            internal static LegacyPptWriterBackground Solid(OfficeColor color) =>
                new(true, 0U, 0D, color.A, color.A,
                    new[] { new LegacyPptWriterGradientStop(color, 0D) });

            internal static LegacyPptWriterBackground Gradient(uint fillType,
                double angleDegrees, byte foregroundAlpha, byte backgroundAlpha,
                IReadOnlyList<LegacyPptWriterGradientStop> stops) =>
                new(true, fillType, angleDegrees, foregroundAlpha,
                    backgroundAlpha, stops);

            internal static LegacyPptWriterBackground Picture(
                PowerPointDocument presentation, string partName, OfficeXmlNode fill,
                byte[] imageBytes, string contentType) =>
                new(true, 3U, 0D, byte.MaxValue, byte.MaxValue,
                    Array.Empty<LegacyPptWriterGradientStop>(), fill,
                    imageBytes, contentType, presentation, partName);
        }

        internal readonly struct LegacyPptWriterGradientStop {
            internal LegacyPptWriterGradientStop(OfficeColor color,
                double position) {
                Color = color;
                Position = position;
            }

            internal OfficeColor Color { get; }
            internal double Position { get; }
        }

        internal sealed class LegacyPptWriterFoptProperty {
            internal LegacyPptWriterFoptProperty(ushort operationId, uint value,
                byte[]? complexData = null) {
                OperationId = operationId;
                Value = value;
                ComplexData = complexData == null ? null : (byte[])complexData.Clone();
            }

            internal ushort OperationId { get; }
            internal ushort PropertyId => checked((ushort)(OperationId & 0x3FFF));
            internal uint Value { get; }
            internal byte[]? ComplexData { get; }
        }
    }
}
