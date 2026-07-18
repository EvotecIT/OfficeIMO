using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort OfficeArtFopt = 0xF00B;
        private const ushort OfficeArtTertiaryFopt = 0xF122;
        private const uint OfficeArtBackgroundShapeFlag = 1U << 10;

        internal static bool TryReadBackground(PowerPointSlide slide,
            out LegacyPptWriterBackground? background, out string? reason) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            P.Background? source = slide.SlidePart.Slide?.CommonSlideData?.Background;
            OpenXmlPart? ownerPart = source == null ? null : slide.SlidePart;
            if (source == null) {
                SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
                source = layoutPart?.SlideLayout?.CommonSlideData?.Background;
                ownerPart = source == null ? null : layoutPart;
            }
            if (source == null || ownerPart == null) {
                background = null;
                reason = null;
                return true;
            }
            return TryReadBackground(ownerPart, source, "slide or layout",
                out background, out reason);
        }

        internal static bool TryReadBackground(SlideMasterPart masterPart,
            out LegacyPptWriterBackground? background, out string? reason) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            P.Background? source = masterPart.SlideMaster?.CommonSlideData?.Background;
            if (source == null) {
                background = null;
                reason = null;
                return true;
            }
            return TryReadBackground(masterPart, source, "slide master",
                out background, out reason);
        }

        internal static bool TryReadBackground(NotesMasterPart masterPart,
            out LegacyPptWriterBackground? background, out string? reason) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            P.Background? source = masterPart.NotesMaster?.CommonSlideData?.Background;
            if (source == null) {
                background = null;
                reason = null;
                return true;
            }
            return TryReadBackground(masterPart, source, "notes master",
                out background, out reason);
        }

        internal static bool TryReadBackground(HandoutMasterPart masterPart,
            out LegacyPptWriterBackground? background, out string? reason) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            P.Background? source = masterPart.HandoutMaster?.CommonSlideData?
                .Background;
            if (source == null) {
                background = null;
                reason = null;
                return true;
            }
            return TryReadBackground(masterPart, source, "handout master",
                out background, out reason);
        }

        internal static bool TryReadBackground(SlideLayoutPart masterPart,
            out LegacyPptWriterBackground? background, out string? reason) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            P.Background? source = masterPart.SlideLayout?.CommonSlideData?
                .Background;
            if (source == null) {
                background = null;
                reason = null;
                return true;
            }
            return TryReadBackground(masterPart, source, "title master layout",
                out background, out reason);
        }

        internal static bool TryReadBackground(NotesSlidePart notesPart,
            out LegacyPptWriterBackground? background, out string? reason) {
            if (notesPart == null) throw new ArgumentNullException(nameof(notesPart));
            P.Background? source = notesPart.NotesSlide?.CommonSlideData?
                .Background;
            if (source == null) {
                background = null;
                reason = null;
                return true;
            }
            return TryReadBackground(notesPart, source, "notes page",
                out background, out reason);
        }

        private static bool TryReadBackground(OpenXmlPart ownerPart,
            P.Background source, string ownerName,
            out LegacyPptWriterBackground? background, out string? reason) {
            A.ColorScheme? colorScheme = GetBackgroundColorScheme(ownerPart);
            A.SchemeColor? placeholderColor = source.BackgroundStyleReference?
                .GetFirstChild<A.SchemeColor>();
            OpenXmlElement? fill = GetBackgroundFill(ownerPart, source,
                out OpenXmlPart fillOwnerPart);
            if (fill == null) {
                if (source.BackgroundStyleReference != null) {
                    background = null;
                    reason = $"The {ownerName} background style reference cannot be resolved from its active theme.";
                    return false;
                }
                background = null;
                reason = null;
                return true;
            }
            if (fill is A.GroupFill) {
                background = null;
                reason = null;
                return true;
            }
            if (fill is A.NoFill) {
                background = LegacyPptWriterBackground.NoFill();
                reason = null;
                return true;
            }
            if (fill is A.SolidFill solid) {
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
                    solid, colorScheme, placeholderColor);
                if (!color.HasValue) {
                    background = null;
                    reason = $"The {ownerName} solid background color cannot be resolved for binary PowerPoint writing.";
                    return false;
                }
                background = LegacyPptWriterBackground.Solid(color.Value);
                reason = null;
                return true;
            }
            if (fill is A.GradientFill gradient) {
                return TryReadGradientBackground(gradient, colorScheme,
                    placeholderColor, ownerName, out background, out reason);
            }
            if (fill is A.BlipFill picture) {
                return TryReadPictureBackground(fillOwnerPart, picture,
                    ownerName, out background, out reason);
            }

            background = null;
            reason = fill switch {
                A.PatternFill => $"The {ownerName} pattern background is not yet encoded by the native binary writer.",
                _ => $"The {ownerName} background fill type '{fill.LocalName}' is not supported by the native binary writer."
            };
            return false;
        }

        private static bool TryReadPictureBackground(OpenXmlPart ownerPart,
            A.BlipFill fill, string ownerName,
            out LegacyPptWriterBackground? background, out string? reason) {
            background = null;
            A.Blip? blip = fill.Blip;
            string? relationshipId = blip?.Embed?.Value;
            if (blip == null || string.IsNullOrWhiteSpace(relationshipId)) {
                reason = $"The {ownerName} picture background has no embedded image relationship.";
                return false;
            }
            if (!string.IsNullOrWhiteSpace(blip.Link?.Value)) {
                reason = $"The {ownerName} picture background is linked; embedding it in binary PowerPoint would lose the link.";
                return false;
            }
            if (blip.HasChildren) {
                reason = $"The {ownerName} picture background uses image effects that have no lossless classic background-fill mapping.";
                return false;
            }
            if (fill.SourceRectangle != null) {
                reason = $"The {ownerName} picture background uses source cropping that is not yet encoded by the native binary writer.";
                return false;
            }
            if (fill.GetFirstChild<A.Tile>() != null) {
                reason = $"The {ownerName} picture background uses tiled placement and cannot be converted to one stretched classic picture fill without changing its appearance.";
                return false;
            }
            A.Stretch? stretch = fill.GetFirstChild<A.Stretch>();
            A.FillRectangle? fillRectangle = stretch?
                .GetFirstChild<A.FillRectangle>();
            if (stretch == null || fillRectangle == null
                || fillRectangle.HasAttributes || fillRectangle.HasChildren
                || fill.ChildElements.Any(child => child is not A.Blip
                    && child is not A.Stretch)) {
                reason = $"The {ownerName} picture background does not use the plain full-frame stretch mapping supported by classic binary PowerPoint.";
                return false;
            }

            ImagePart? imagePart;
            try {
                imagePart = ownerPart.GetPartById(relationshipId!) as ImagePart;
            } catch (ArgumentOutOfRangeException) {
                reason = $"The {ownerName} picture background relationship cannot be resolved.";
                return false;
            }
            if (imagePart == null) {
                reason = $"The {ownerName} picture background relationship does not target an image part.";
                return false;
            }
            if (imagePart.Parts.Any() || imagePart.ExternalRelationships.Any()
                || imagePart.HyperlinkRelationships.Any()) {
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
                using Stream stream = imagePart.GetStream(FileMode.Open,
                    FileAccess.Read);
                imageBytes = OfficeStreamReader.ReadAllBytes(stream,
                    64 * 1024 * 1024);
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
            background = LegacyPptWriterBackground.Picture(fill, imageBytes,
                contentType);
            reason = null;
            return true;
        }

        private static bool TryReadGradientBackground(A.GradientFill gradient,
            A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor,
            string ownerName, out LegacyPptWriterBackground? background,
            out string? reason) {
            A.LinearGradientFill? linear = gradient.GetFirstChild<A.LinearGradientFill>();
            A.PathGradientFill? path = gradient.GetFirstChild<A.PathGradientFill>();
            if ((linear == null) == (path == null)) {
                background = null;
                reason = $"The {ownerName} gradient background must contain exactly one linear or supported path geometry.";
                return false;
            }
            A.TileFlipValues? flip = gradient.Flip?.Value;
            if (flip.HasValue && flip.Value != A.TileFlipValues.None) {
                background = null;
                reason = $"The {ownerName} gradient background uses tile flipping that the binary PowerPoint writer cannot reproduce losslessly.";
                return false;
            }
            if (gradient.ChildElements.Any(child =>
                    child is not A.GradientStopList
                    && !ReferenceEquals(child, linear)
                    && !ReferenceEquals(child, path))) {
                background = null;
                reason = $"The {ownerName} gradient background uses a custom tile rectangle that the binary PowerPoint format cannot reproduce through this writer.";
                return false;
            }
            if (path?.HasChildren == true) {
                background = null;
                reason = $"The {ownerName} path gradient uses a custom fill rectangle that has no lossless binary PowerPoint mapping.";
                return false;
            }
            A.GradientStop[] sourceStops = gradient.GetFirstChild<A.GradientStopList>()?
                .Elements<A.GradientStop>()
                .OrderBy(stop => stop.Position?.Value ?? 0)
                .ToArray() ?? Array.Empty<A.GradientStop>();
            if (sourceStops.Length < 2 || sourceStops.Length > ushort.MaxValue) {
                background = null;
                reason = $"The {ownerName} gradient background needs at least two resolvable stops.";
                return false;
            }

            var stops = new List<LegacyPptWriterGradientStop>(sourceStops.Length);
            foreach (A.GradientStop sourceStop in sourceStops) {
                int position = sourceStop.Position?.Value ?? 0;
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
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
            if (linear != null) {
                double openXmlAngle = (linear.Angle?.Value ?? 0) / 60000D;
                legacyAngle = NormalizeBackgroundAngle(270D - openXmlAngle);
                fillType = linear.Scaled?.Value == true ? 7U : 4U;
            } else if (path!.Path?.Value == A.PathShadeValues.Circle) {
                fillType = 5U;
            } else if (path.Path?.Value == A.PathShadeValues.Shape) {
                fillType = 6U;
            } else {
                background = null;
                reason = $"The {ownerName} path gradient geometry '{path.Path?.Value}' has no lossless binary PowerPoint mapping.";
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

        private static OpenXmlElement? GetBackgroundFill(OpenXmlPart ownerPart,
            P.Background source, out OpenXmlPart fillOwnerPart) {
            fillOwnerPart = ownerPart;
            P.BackgroundProperties? properties = source.BackgroundProperties;
            if (properties != null && properties.HasChildren) {
                return properties.ChildElements.FirstOrDefault();
            }
            P.BackgroundStyleReference? reference = source.BackgroundStyleReference;
            if (reference?.Index?.Value == null) return null;
            A.FormatScheme? formatScheme = GetBackgroundFormatScheme(
                ownerPart, out OpenXmlPart? themePart);
            if (formatScheme == null || themePart == null) return null;
            fillOwnerPart = themePart;
            uint index = reference.Index.Value;
            if (index >= 1001U) {
                OpenXmlElementList fills = formatScheme
                    .GetFirstChild<A.BackgroundFillStyleList>()?
                    .ChildElements ?? default;
                uint zeroBased = index - 1001U;
                return zeroBased < unchecked((uint)fills.Count)
                    ? fills[unchecked((int)zeroBased)]
                    : null;
            }
            if (index < 1U) return null;
            OpenXmlElementList styles = formatScheme
                .GetFirstChild<A.FillStyleList>()?.ChildElements ?? default;
            uint styleIndex = index - 1U;
            return styleIndex < unchecked((uint)styles.Count)
                ? styles[unchecked((int)styleIndex)]
                : null;
        }

        private static A.ColorScheme? GetBackgroundColorScheme(
            OpenXmlPart ownerPart) => GetBackgroundThemeSources(ownerPart)
            .Select(GetColorScheme)
            .FirstOrDefault(scheme => scheme != null);

        private static A.ColorScheme? GetColorScheme(
            OpenXmlPart themePart) => themePart switch {
                ThemeOverridePart themeOverridePart => themeOverridePart
                    .ThemeOverride?.ColorScheme,
                ThemePart masterThemePart => masterThemePart.Theme?
                    .ThemeElements?.ColorScheme,
                _ => null
            };

        private static A.FormatScheme? GetBackgroundFormatScheme(
            OpenXmlPart ownerPart, out OpenXmlPart? themeSourcePart) {
            foreach (OpenXmlPart candidate in
                     GetBackgroundThemeSources(ownerPart)) {
                A.FormatScheme? scheme = GetFormatScheme(candidate);
                if (scheme == null) continue;
                themeSourcePart = candidate;
                return scheme;
            }
            themeSourcePart = null;
            return null;
        }

        private static A.FormatScheme? GetFormatScheme(
            OpenXmlPart themePart) => themePart switch {
                ThemeOverridePart themeOverridePart => themeOverridePart
                    .ThemeOverride?.FormatScheme,
                ThemePart masterThemePart => masterThemePart.Theme?
                    .ThemeElements?.FormatScheme,
                _ => null
            };

        private static IEnumerable<OpenXmlPart> GetBackgroundThemeSources(
            OpenXmlPart ownerPart) {
            switch (ownerPart) {
                case SlidePart slidePart:
                    if (slidePart.ThemeOverridePart != null) {
                        yield return slidePart.ThemeOverridePart;
                    }
                    if (slidePart.SlideLayoutPart?.ThemeOverridePart != null) {
                        yield return slidePart.SlideLayoutPart.ThemeOverridePart;
                    }
                    if (slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart
                        != null) {
                        yield return slidePart.SlideLayoutPart
                            .SlideMasterPart.ThemePart;
                    }
                    break;
                case SlideLayoutPart layoutPart:
                    if (layoutPart.ThemeOverridePart != null) {
                        yield return layoutPart.ThemeOverridePart;
                    }
                    if (layoutPart.SlideMasterPart?.ThemePart != null) {
                        yield return layoutPart.SlideMasterPart.ThemePart;
                    }
                    break;
                case SlideMasterPart masterPart
                    when masterPart.ThemePart != null:
                    yield return masterPart.ThemePart;
                    break;
                case NotesSlidePart notesPart:
                    if (notesPart.ThemeOverridePart != null) {
                        yield return notesPart.ThemeOverridePart;
                    }
                    if (notesPart.NotesMasterPart?.ThemePart != null) {
                        yield return notesPart.NotesMasterPart.ThemePart;
                    }
                    break;
                case NotesMasterPart notesMasterPart
                    when notesMasterPart.ThemePart != null:
                    yield return notesMasterPart.ThemePart;
                    break;
                case HandoutMasterPart handoutMasterPart
                    when handoutMasterPart.ThemePart != null:
                    yield return handoutMasterPart.ThemePart;
                    break;
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
            if (prototype == null) throw new ArgumentNullException(
                nameof(prototype));
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
            if (background.PictureFill != null) {
                LegacyPptWriterPicture picture = pictureCatalog?
                    .Get(background.PictureFill)
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

        internal sealed class LegacyPptWriterBackground {
            private LegacyPptWriterBackground(bool filled, uint fillType,
                double angleDegrees, byte foregroundAlpha, byte backgroundAlpha,
                IReadOnlyList<LegacyPptWriterGradientStop> stops,
                A.BlipFill? pictureFill = null, byte[]? pictureBytes = null,
                string? pictureContentType = null) {
                Filled = filled;
                FillType = fillType;
                AngleDegrees = angleDegrees;
                ForegroundAlpha = foregroundAlpha;
                BackgroundAlpha = backgroundAlpha;
                Stops = new ReadOnlyCollection<LegacyPptWriterGradientStop>(
                    stops.ToArray());
                PictureFill = pictureFill;
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
            internal A.BlipFill? PictureFill { get; }
            internal byte[] PictureBytes { get; }
            internal string? PictureContentType { get; }
            internal bool RequiresPictureCatalog => PictureFill != null;

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

            internal static LegacyPptWriterBackground Picture(A.BlipFill fill,
                byte[] imageBytes, string contentType) =>
                new(true, 3U, 0D, byte.MaxValue, byte.MaxValue,
                    Array.Empty<LegacyPptWriterGradientStop>(), fill,
                    imageBytes, contentType);
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
