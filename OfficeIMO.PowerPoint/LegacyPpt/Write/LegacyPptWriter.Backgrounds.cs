using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort OfficeArtFopt = 0xF00B;
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
            A.ColorScheme? colorScheme = GetBackgroundThemePart(ownerPart)?
                .Theme?.ThemeElements?.ColorScheme;
            A.SchemeColor? placeholderColor = source.BackgroundStyleReference?
                .GetFirstChild<A.SchemeColor>();
            OpenXmlElement? fill = GetBackgroundFill(ownerPart, source);
            if (fill == null || fill is A.GroupFill) {
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

            background = null;
            reason = fill switch {
                A.BlipFill => $"The {ownerName} image background is not yet encoded by the native binary writer.",
                A.PatternFill => $"The {ownerName} pattern background is not yet encoded by the native binary writer.",
                _ => $"The {ownerName} background fill type '{fill.LocalName}' is not supported by the native binary writer."
            };
            return false;
        }

        private static bool TryReadGradientBackground(A.GradientFill gradient,
            A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor,
            string ownerName, out LegacyPptWriterBackground? background,
            out string? reason) {
            A.LinearGradientFill? linear = gradient.GetFirstChild<A.LinearGradientFill>();
            if (linear == null) {
                background = null;
                reason = $"The {ownerName} path or radial gradient background is not yet encoded by the native binary writer.";
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

            if (stops.Any(stop => stop.Color.A != byte.MaxValue)) {
                background = null;
                reason = $"The {ownerName} gradient uses stop opacity that is not yet encoded by the native binary writer.";
                return false;
            }

            double openXmlAngle = (linear.Angle?.Value ?? 0) / 60000D;
            double legacyAngle = NormalizeBackgroundAngle(270D - openXmlAngle);
            uint fillType = linear.Scaled?.Value == true ? 7U : 4U;
            background = LegacyPptWriterBackground.Gradient(fillType,
                legacyAngle, stops);
            reason = null;
            return true;
        }

        private static OpenXmlElement? GetBackgroundFill(OpenXmlPart ownerPart,
            P.Background source) {
            P.BackgroundProperties? properties = source.BackgroundProperties;
            if (properties != null && properties.HasChildren) {
                return properties.ChildElements.FirstOrDefault();
            }
            P.BackgroundStyleReference? reference = source.BackgroundStyleReference;
            if (reference?.Index?.Value == null) return null;
            A.FormatScheme? formatScheme = GetBackgroundThemePart(ownerPart)?
                .Theme?.ThemeElements?.FormatScheme;
            if (formatScheme == null) return null;
            uint index = reference.Index.Value;
            if (index >= 1001U) {
                return formatScheme.GetFirstChild<A.BackgroundFillStyleList>()?
                    .ChildElements.ElementAtOrDefault(checked((int)(index - 1001U)));
            }
            return index >= 1U
                ? formatScheme.GetFirstChild<A.FillStyleList>()?
                    .ChildElements.ElementAtOrDefault(checked((int)index - 1))
                : null;
        }

        private static ThemePart? GetBackgroundThemePart(OpenXmlPart ownerPart) =>
            ownerPart switch {
                SlidePart slidePart => slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart,
                SlideLayoutPart layoutPart => layoutPart.SlideMasterPart?.ThemePart,
                SlideMasterPart masterPart => masterPart.ThemePart,
                NotesSlidePart notesPart => notesPart.NotesMasterPart?.ThemePart,
                NotesMasterPart notesMasterPart => notesMasterPart.ThemePart,
                HandoutMasterPart handoutMasterPart => handoutMasterPart.ThemePart,
                _ => null
            };

        private static byte[] BuildBackgroundDrawingRecord(LegacyPptRecord drawing,
            LegacyPptWriterBackground background) {
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
                        drawingChildren.Add(BuildBackgroundShapeRecord(drawingChild,
                            background));
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

        private static byte[] BuildBackgroundShapeRecord(LegacyPptRecord prototype,
            LegacyPptWriterBackground background) {
            var children = new List<byte[]>(prototype.Children.Count + 1);
            bool wroteFopt = false;
            foreach (LegacyPptRecord child in prototype.Children) {
                if (child.Type == OfficeArtFopt) {
                    children.Add(BuildBackgroundFoptRecord(child, background));
                    wroteFopt = true;
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (!wroteFopt) {
                children.Insert(Math.Min(1, children.Count),
                    BuildBackgroundFoptRecord(null, background));
            }
            return BuildContainer(OfficeArtSpContainer, prototype.Instance, children);
        }

        private static byte[] BuildBackgroundFoptRecord(LegacyPptRecord? prototype,
            LegacyPptWriterBackground background) {
            List<LegacyPptWriterFoptProperty> properties = prototype == null
                ? new List<LegacyPptWriterFoptProperty>()
                : ReadFoptProperties(prototype).Where(property =>
                    property.PropertyId < 0x0180 || property.PropertyId > 0x01BF)
                    .ToList();
            properties.Add(new LegacyPptWriterFoptProperty(0x0180,
                background.FillType));
            if (background.Filled) {
                LegacyPptWriterGradientStop first = background.Stops[0];
                LegacyPptWriterGradientStop last = background.Stops[background.Stops.Count - 1];
                properties.Add(new LegacyPptWriterFoptProperty(0x0181,
                    PackOfficeArtColor(first.Color)));
                properties.Add(new LegacyPptWriterFoptProperty(0x0182,
                    PackOfficeArtOpacity(first.Color.A)));
                properties.Add(new LegacyPptWriterFoptProperty(0x0183,
                    PackOfficeArtColor(last.Color)));
                properties.Add(new LegacyPptWriterFoptProperty(0x0184,
                    PackOfficeArtOpacity(last.Color.A)));
                if (background.FillType is 4U or 7U) {
                    properties.Add(new LegacyPptWriterFoptProperty(0x018B,
                        unchecked((uint)checked((int)Math.Round(
                            background.AngleDegrees * 65536D,
                            MidpointRounding.AwayFromZero)))));
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
            IReadOnlyList<LegacyPptWriterFoptProperty> source) {
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
                OfficeArtFopt, payload);
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
                double angleDegrees,
                IReadOnlyList<LegacyPptWriterGradientStop> stops) {
                Filled = filled;
                FillType = fillType;
                AngleDegrees = angleDegrees;
                Stops = new ReadOnlyCollection<LegacyPptWriterGradientStop>(
                    stops.ToArray());
            }

            internal bool Filled { get; }
            internal uint FillType { get; }
            internal double AngleDegrees { get; }
            internal IReadOnlyList<LegacyPptWriterGradientStop> Stops { get; }

            internal static LegacyPptWriterBackground NoFill() =>
                new(false, 0U, 0D, Array.Empty<LegacyPptWriterGradientStop>());

            internal static LegacyPptWriterBackground Solid(OfficeColor color) =>
                new(true, 0U, 0D,
                    new[] { new LegacyPptWriterGradientStop(color, 0D) });

            internal static LegacyPptWriterBackground Gradient(uint fillType,
                double angleDegrees,
                IReadOnlyList<LegacyPptWriterGradientStop> stops) =>
                new(true, fillType, angleDegrees, stops);
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

        private sealed class LegacyPptWriterFoptProperty {
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
