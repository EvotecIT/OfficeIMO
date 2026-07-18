using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using System.Globalization;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static bool TryReadShapeFormatting(PowerPointShape shape,
            out LegacyPptWriterShapeFormatting formatting,
            out string? reason) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (!TryReadShapeTransform(shape,
                    out LegacyPptWriterShapeFormatting transform,
                    out reason)
                || !TryReadShapeVisualStyle(shape,
                    out IReadOnlyList<LegacyPptWriterFoptProperty> visual,
                    out reason)
                || !TryReadShapeMetadataForWrite(shape,
                    out IReadOnlyList<LegacyPptWriterFoptProperty> metadata,
                    out reason)
                || !TryReadShapeVisibilityForWrite(shape,
                    out IReadOnlyList<LegacyPptWriterFoptProperty> visibility,
                    out reason)) {
                formatting = LegacyPptWriterShapeFormatting.Empty;
                return false;
            }
            var properties = transform.Properties.ToList();
            properties.AddRange(visual);
            properties.AddRange(metadata);
            properties.AddRange(visibility);
            if (shape is PowerPointTextBox textBox) {
                if (!TryReadTextFrameForWrite(textBox,
                        out IReadOnlyList<LegacyPptWriterFoptProperty> frame,
                        out reason)) {
                    formatting = LegacyPptWriterShapeFormatting.Empty;
                    return false;
                }
                properties.AddRange(frame);
            }
            formatting = new LegacyPptWriterShapeFormatting(properties,
                transform.FspFlags);
            reason = null;
            return true;
        }

        internal static bool TryReadShapeVisibilityForWrite(
            PowerPointShape shape,
            out IReadOnlyList<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            properties = shape.Hidden
                ? new[] {
                    new LegacyPptWriterFoptProperty(0x03BF,
                        (1U << 14) | (1U << 30))
                }
                : Array.Empty<LegacyPptWriterFoptProperty>();
            reason = null;
            return true;
        }

        internal static bool TryReadShapeMetadataForWrite(
            PowerPointShape shape,
            out IReadOnlyList<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (!string.IsNullOrWhiteSpace(shape.Title)) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "Accessibility titles have no native PowerPoint 97-2003 shape-property representation.";
                return false;
            }
            if (shape.Decorative) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "The modern decorative-shape flag has no native PowerPoint 97-2003 representation.";
                return false;
            }
            var result = new List<LegacyPptWriterFoptProperty>(2);
            AddShapeMetadataText(result, 0x0380, shape.Name);
            AddShapeMetadataText(result, 0x0381, shape.Description);
            properties = result;
            reason = null;
            return true;
        }

        private static void AddShapeMetadataText(
            ICollection<LegacyPptWriterFoptProperty> properties,
            ushort propertyId, string? value) {
            if (string.IsNullOrEmpty(value)) return;
            byte[] text = Encoding.Unicode.GetBytes(value + "\0");
            properties.Add(new LegacyPptWriterFoptProperty(
                unchecked((ushort)(0x8000 | propertyId)),
                unchecked((uint)text.Length), text));
        }

        internal static bool TryReadShapeTransform(PowerPointShape shape,
            out LegacyPptWriterShapeFormatting formatting,
            out string? reason) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            var properties = new List<LegacyPptWriterFoptProperty>(1);
            uint fspFlags = 0x00000A00U;
            if (shape.HorizontalFlip == true) fspFlags |= 1U << 6;
            if (shape.VerticalFlip == true) fspFlags |= 1U << 7;
            if (shape.Rotation.HasValue) {
                double normalized = shape.Rotation.Value % 360D;
                if (double.IsNaN(normalized)
                    || double.IsInfinity(normalized)) {
                    formatting = LegacyPptWriterShapeFormatting.Empty;
                    reason = "Shape rotation must be a finite angle for binary PowerPoint.";
                    return false;
                }
                if (normalized < 0D) normalized += 360D;
                properties.Add(new LegacyPptWriterFoptProperty(0x0004,
                    unchecked((uint)checked((int)Math.Round(
                        normalized * 65536D,
                        MidpointRounding.AwayFromZero)))));
            }
            formatting = new LegacyPptWriterShapeFormatting(properties,
                fspFlags);
            reason = null;
            return true;
        }

        internal static bool TryReadShapeGeometry(PowerPointShape shape,
            ushort shapeType,
            out IReadOnlyList<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            A.PresetGeometry? geometry = shape.Element
                .Descendants<A.PresetGeometry>().FirstOrDefault();
            A.AdjustValueList? values = geometry?.AdjustValueList;
            if (values == null || !values.HasChildren) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = null;
                return true;
            }
            if (values.HasAttributes
                || values.ChildElements.Any(child => child is not A.ShapeGuide)
                || values.Elements<A.ShapeGuide>().Skip(1).Any()
                || shapeType is not 2 and not 23) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "Only one exact round-rectangle or donut adjustment can currently be written to binary PowerPoint.";
                return false;
            }
            A.ShapeGuide guide = values.Elements<A.ShapeGuide>().Single();
            if (guide.HasChildren
                || guide.GetAttributes().Any(attribute =>
                    attribute.LocalName is not "name" and not "fmla")
                || !string.Equals(guide.Name?.Value, "adj",
                    StringComparison.Ordinal)
                || !TryReadConstantGuide(guide.Formula?.Value,
                    out long drawingValue)) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "The shape adjustment must be a single 'adj' guide with a constant 'val' formula.";
                return false;
            }
            double rawValue = drawingValue * (21600D / 100000D);
            if (drawingValue < 0L || drawingValue > 100000L
                || rawValue < int.MinValue || rawValue > int.MaxValue) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "The shape adjustment lies outside the exact classic OfficeArt range.";
                return false;
            }
            int officeArtValue = checked((int)Math.Round(rawValue,
                MidpointRounding.AwayFromZero));
            long projectedValue = checked((long)Math.Round(
                officeArtValue * (100000D / 21600D),
                MidpointRounding.AwayFromZero));
            if (projectedValue != drawingValue) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "The DrawingML adjustment cannot be represented exactly by the integer classic OfficeArt adjustment slot.";
                return false;
            }
            properties = new[] {
                new LegacyPptWriterFoptProperty(0x0147,
                    unchecked((uint)officeArtValue))
            };
            reason = null;
            return true;
        }

        private static bool TryReadConstantGuide(string? formula,
            out long value) {
            value = 0L;
            if (formula == null || string.IsNullOrWhiteSpace(formula)) return false;
            string[] parts = formula.Trim().Split(new[] { ' ' },
                StringSplitOptions.RemoveEmptyEntries);
            return parts.Length == 2
                && string.Equals(parts[0], "val",
                    StringComparison.Ordinal)
                && long.TryParse(parts[1], NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out value);
        }

        internal static bool TryReadShapeVisualStyle(PowerPointShape shape,
            out IReadOnlyList<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (shape.Element.GetFirstChild<P.ShapeStyle>() != null) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "Theme-referenced shape styles must be resolved before binary PowerPoint writing.";
                return false;
            }
            var result = new List<LegacyPptWriterFoptProperty>();
            P.ShapeProperties? shapeProperties = shape.Element switch {
                P.Shape value => value.ShapeProperties,
                P.ConnectionShape value => value.ShapeProperties,
                P.Picture value => value.ShapeProperties,
                _ => null
            };
            if (shapeProperties != null
                && !TryReadShapeProperties(shapeProperties, result,
                    out reason)) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                return false;
            }
            properties = result;
            reason = null;
            return true;
        }

        private static bool TryReadShapeProperties(
            P.ShapeProperties source,
            ICollection<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            foreach (OpenXmlElement child in source.ChildElements) {
                if (child is A.Transform2D or A.PresetGeometry
                    or A.NoFill or A.SolidFill or A.Outline
                    or A.EffectList) {
                    continue;
                }
                reason = $"The shape property '{child.LocalName}' has no lossless classic OfficeArt mapping.";
                return false;
            }

            OpenXmlElement[] fills = source.ChildElements.Where(child =>
                child is A.NoFill or A.SolidFill).ToArray();
            if (fills.Length > 1) {
                reason = "The shape contains multiple explicit fills.";
                return false;
            }
            if (fills.Length == 1) {
                if (fills[0] is A.NoFill noFill) {
                    if (noFill.HasAttributes || noFill.HasChildren) {
                        reason = "The no-fill shape property contains an unsupported extension.";
                        return false;
                    }
                    properties.Add(new LegacyPptWriterFoptProperty(0x01BF,
                        0x00100000U));
                } else if (!TryReadRgbFill((A.SolidFill)fills[0],
                               out OfficeColor fillColor,
                               out int? fillAlpha, out reason)) {
                    return false;
                } else {
                    properties.Add(new LegacyPptWriterFoptProperty(0x0180,
                        0U));
                    properties.Add(new LegacyPptWriterFoptProperty(0x0181,
                        PackOfficeArtColor(fillColor)));
                    if (fillAlpha.HasValue) {
                        properties.Add(new LegacyPptWriterFoptProperty(0x0182,
                            ToOfficeArtOpacity(fillAlpha.Value)));
                    }
                    properties.Add(new LegacyPptWriterFoptProperty(0x01BF,
                        0x00100010U));
                }
            }

            A.Outline? outline = source.GetFirstChild<A.Outline>();
            if (outline != null
                && !TryReadOutline(outline, properties, out reason)) {
                return false;
            }
            A.EffectList? effects = source.GetFirstChild<A.EffectList>();
            if (effects != null
                && !TryReadShapeEffects(effects, properties, out reason)) {
                return false;
            }
            reason = null;
            return true;
        }

        private static bool TryReadOutline(A.Outline outline,
            ICollection<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (outline.GetAttributes().Any(attribute =>
                    attribute.LocalName is not "w" and not "cap")) {
                reason = "Binary PowerPoint does not support the requested compound or aligned outline.";
                return false;
            }
            OpenXmlElement[] fills = outline.ChildElements.Where(child =>
                child is A.NoFill or A.SolidFill).ToArray();
            if (fills.Length > 1) {
                reason = "The shape outline contains multiple explicit fills.";
                return false;
            }
            if (fills.Length == 0 && (outline.HasAttributes
                    || outline.HasChildren)) {
                reason = "An explicitly formatted outline requires an RGB solid fill or no-fill for binary PowerPoint.";
                return false;
            }
            if (fills.Length == 1 && fills[0] is A.NoFill noFill) {
                if (noFill.HasAttributes || noFill.HasChildren) {
                    reason = "The no-fill outline contains an unsupported extension.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(0x01FF,
                    0x00080000U));
            } else if (fills.Length == 1) {
                if (!TryReadRgbFill((A.SolidFill)fills[0],
                        out OfficeColor lineColor, out int? lineAlpha,
                        out reason)) {
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(0x01C0,
                    PackOfficeArtColor(lineColor)));
                if (lineAlpha.HasValue) {
                    properties.Add(new LegacyPptWriterFoptProperty(0x01C1,
                        ToOfficeArtOpacity(lineAlpha.Value)));
                }
                properties.Add(new LegacyPptWriterFoptProperty(0x01FF,
                    0x00080008U));
            }
            if (outline.Width?.Value is int width) {
                if (width < 0) {
                    reason = "Shape outline width cannot be negative.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(0x01CB,
                    checked((uint)width)));
            }
            if (outline.CapType?.Value is A.LineCapValues cap) {
                uint value = cap == A.LineCapValues.Round ? 0U
                    : cap == A.LineCapValues.Square ? 1U
                    : cap == A.LineCapValues.Flat ? 2U
                    : uint.MaxValue;
                if (value == uint.MaxValue) {
                    reason = "The outline cap has no classic OfficeArt mapping.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(0x01D7,
                    value));
            }

            foreach (OpenXmlElement child in outline.ChildElements) {
                switch (child) {
                    case A.NoFill:
                    case A.SolidFill:
                        break;
                    case A.PresetDash dash:
                        if (dash.HasChildren
                            || dash.GetAttributes().Any(attribute =>
                                attribute.LocalName != "val")
                            || !TryMapLineDash(dash.Val?.Value,
                                out uint dashValue)) {
                            reason = "The outline dash pattern has no classic OfficeArt mapping.";
                            return false;
                        }
                        properties.Add(new LegacyPptWriterFoptProperty(0x01CE,
                            dashValue));
                        break;
                    case A.Bevel bevel:
                        if (bevel.HasAttributes || bevel.HasChildren) {
                            reason = "The bevel line join contains an unsupported extension.";
                            return false;
                        }
                        properties.Add(new LegacyPptWriterFoptProperty(0x01D6,
                            0U));
                        break;
                    case A.Miter miter:
                        if (miter.HasAttributes || miter.HasChildren) {
                            reason = "A custom miter limit has no exact classic OfficeArt mapping.";
                            return false;
                        }
                        properties.Add(new LegacyPptWriterFoptProperty(0x01D6,
                            1U));
                        break;
                    case A.Round round:
                        if (round.HasAttributes || round.HasChildren) {
                            reason = "The round line join contains an unsupported extension.";
                            return false;
                        }
                        properties.Add(new LegacyPptWriterFoptProperty(0x01D6,
                            2U));
                        break;
                    case A.HeadEnd head:
                        if (!TryAddLineEnd(head, isHead: true, properties,
                                out reason)) return false;
                        break;
                    case A.TailEnd tail:
                        if (!TryAddLineEnd(tail, isHead: false, properties,
                                out reason)) return false;
                        break;
                    default:
                        reason = $"The outline property '{child.LocalName}' has no lossless classic OfficeArt mapping.";
                        return false;
                }
            }
            reason = null;
            return true;
        }

        private static bool TryReadShapeEffects(A.EffectList effects,
            ICollection<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (effects.HasAttributes || effects.ChildElements.Count > 1
                || effects.FirstChild is not A.OuterShadow shadow) {
                if (!effects.HasChildren && !effects.HasAttributes) {
                    reason = null;
                    return true;
                }
                reason = "Only one classic outer shadow can be written as a native binary PowerPoint shape effect.";
                return false;
            }
            if (shadow.GetAttributes().Any(attribute => attribute.LocalName
                    is not "blurRad" and not "dist" and not "dir"
                    and not "rotWithShape")
                || shadow.RotateWithShape?.Value == true
                || shadow.ChildElements.Count != 1
                || shadow.GetFirstChild<A.RgbColorModelHex>() is not
                    A.RgbColorModelHex rgb
                || !TryReadRgbColor(rgb, out OfficeColor color,
                    out int? alpha)) {
                reason = "The outer shadow uses scaling, skew, alignment, rotation, or color transforms that have no exact classic OfficeArt mapping.";
                return false;
            }
            long distance = shadow.Distance?.Value ?? 0L;
            long blur = shadow.BlurRadius?.Value ?? 0L;
            if (distance < 0L || blur < 0L) {
                reason = "Outer-shadow distance and blur radius cannot be negative.";
                return false;
            }
            int direction = shadow.Direction?.Value ?? 0;
            double radians = direction / 60000D * Math.PI / 180D;
            long offsetX = checked((long)Math.Round(Math.Cos(radians)
                * distance, MidpointRounding.AwayFromZero));
            long offsetY = checked((long)Math.Round(Math.Sin(radians)
                * distance, MidpointRounding.AwayFromZero));
            if (offsetX < int.MinValue || offsetX > int.MaxValue
                || offsetY < int.MinValue || offsetY > int.MaxValue
                || blur > int.MaxValue) {
                reason = "The outer shadow exceeds the classic OfficeArt coordinate range.";
                return false;
            }
            properties.Add(new LegacyPptWriterFoptProperty(0x0200, 0U));
            properties.Add(new LegacyPptWriterFoptProperty(0x0201,
                PackOfficeArtColor(color)));
            if (alpha.HasValue) {
                properties.Add(new LegacyPptWriterFoptProperty(0x0204,
                    ToOfficeArtOpacity(alpha.Value)));
            }
            properties.Add(new LegacyPptWriterFoptProperty(0x0205,
                unchecked((uint)checked((int)offsetX))));
            properties.Add(new LegacyPptWriterFoptProperty(0x0206,
                unchecked((uint)checked((int)offsetY))));
            if (blur > 0L) {
                properties.Add(new LegacyPptWriterFoptProperty(0x021C,
                    checked((uint)blur)));
            }
            properties.Add(new LegacyPptWriterFoptProperty(0x023F,
                0x00020002U));
            reason = null;
            return true;
        }

        private static bool TryReadRgbFill(A.SolidFill fill,
            out OfficeColor color, out int? alpha, out string? reason) {
            if (fill.HasAttributes || fill.ChildElements.Count != 1
                || fill.GetFirstChild<A.RgbColorModelHex>() is not
                    A.RgbColorModelHex rgb
                || !TryReadRgbColor(rgb, out color, out alpha)) {
                color = default;
                alpha = null;
                reason = "Binary PowerPoint shape fills require one RGB color with an optional alpha value.";
                return false;
            }
            reason = null;
            return true;
        }

        private static bool TryReadRgbColor(A.RgbColorModelHex rgb,
            out OfficeColor color, out int? alpha) {
            color = default;
            alpha = null;
            if (rgb.Val?.Value == null
                || !OfficeColor.TryParse(rgb.Val.Value, out color)
                || rgb.ExtendedAttributes.Any()
                || rgb.ChildElements.Any(child => child is not A.Alpha)
                || rgb.Elements<A.Alpha>().Skip(1).Any()) {
                return false;
            }
            A.Alpha? sourceAlpha = rgb.GetFirstChild<A.Alpha>();
            if (sourceAlpha == null) return true;
            if (sourceAlpha.HasChildren || sourceAlpha.ExtendedAttributes.Any()
                || sourceAlpha.Val?.Value is not int value
                || value < 0 || value > 100000) {
                return false;
            }
            alpha = value;
            return true;
        }

        private static bool TryAddLineEnd(OpenXmlElement source,
            bool isHead,
            ICollection<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (source.HasChildren || source.GetAttributes().Any(attribute =>
                    attribute.LocalName is not "type" and not "w"
                        and not "len")) {
                reason = "The line end contains an unsupported extension.";
                return false;
            }
            A.LineEndValues? type;
            A.LineEndWidthValues? width;
            A.LineEndLengthValues? length;
            if (source is A.HeadEnd head) {
                type = head.Type?.Value;
                width = head.Width?.Value;
                length = head.Length?.Value;
            } else {
                A.TailEnd tail = (A.TailEnd)source;
                type = tail.Type?.Value;
                width = tail.Width?.Value;
                length = tail.Length?.Value;
            }
            ushort typeId = isHead ? (ushort)0x01D0 : (ushort)0x01D1;
            ushort widthId = isHead ? (ushort)0x01D2 : (ushort)0x01D4;
            ushort lengthId = isHead ? (ushort)0x01D3 : (ushort)0x01D5;
            if (type.HasValue) {
                uint value = type.Value == A.LineEndValues.None ? 0U
                    : type.Value == A.LineEndValues.Triangle ? 1U
                    : type.Value == A.LineEndValues.Stealth ? 2U
                    : type.Value == A.LineEndValues.Diamond ? 3U
                    : type.Value == A.LineEndValues.Oval ? 4U
                    : type.Value == A.LineEndValues.Arrow ? 5U
                    : uint.MaxValue;
                if (value == uint.MaxValue) {
                    reason = "The line-end type has no classic OfficeArt mapping.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(typeId, value));
            }
            if (width.HasValue) {
                uint value = width.Value == A.LineEndWidthValues.Small ? 0U
                    : width.Value == A.LineEndWidthValues.Medium ? 1U
                    : width.Value == A.LineEndWidthValues.Large ? 2U
                    : uint.MaxValue;
                if (value == uint.MaxValue) {
                    reason = "The line-end width has no classic OfficeArt mapping.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(widthId,
                    value));
            }
            if (length.HasValue) {
                uint value = length.Value == A.LineEndLengthValues.Small ? 0U
                    : length.Value == A.LineEndLengthValues.Medium ? 1U
                    : length.Value == A.LineEndLengthValues.Large ? 2U
                    : uint.MaxValue;
                if (value == uint.MaxValue) {
                    reason = "The line-end length has no classic OfficeArt mapping.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(lengthId,
                    value));
            }
            reason = null;
            return true;
        }

        private static bool TryMapLineDash(A.PresetLineDashValues? value,
            out uint mapped) {
            mapped = value == A.PresetLineDashValues.Solid ? 0U
                : value == A.PresetLineDashValues.SystemDash ? 1U
                : value == A.PresetLineDashValues.SystemDot ? 2U
                : value == A.PresetLineDashValues.SystemDashDot ? 3U
                : value == A.PresetLineDashValues.SystemDashDotDot ? 4U
                : value == A.PresetLineDashValues.Dot ? 5U
                : value == A.PresetLineDashValues.Dash ? 6U
                : value == A.PresetLineDashValues.LargeDash ? 7U
                : value == A.PresetLineDashValues.DashDot ? 8U
                : value == A.PresetLineDashValues.LargeDashDot ? 9U
                : value == A.PresetLineDashValues.LargeDashDotDot ? 10U
                : uint.MaxValue;
            return mapped != uint.MaxValue;
        }

        private static uint ToOfficeArtOpacity(int openXmlAlpha) =>
            checked((uint)Math.Round(openXmlAlpha / 100000D * 65536D,
                MidpointRounding.AwayFromZero));

        private static void AddShapeFormattingProperties(
            ICollection<LegacyPptWriterFoptProperty> properties,
            PowerPointShape shape) {
            if (!TryReadShapeFormatting(shape,
                    out LegacyPptWriterShapeFormatting formatting,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            foreach (LegacyPptWriterFoptProperty property in
                     formatting.Properties) {
                properties.Add(property);
            }
        }

        private static void AddShapeTransformProperties(
            ICollection<LegacyPptWriterFoptProperty> properties,
            PowerPointShape shape) {
            if (!TryReadShapeTransform(shape,
                    out LegacyPptWriterShapeFormatting formatting,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            foreach (LegacyPptWriterFoptProperty property in
                     formatting.Properties) {
                properties.Add(property);
            }
        }

        private static void AddShapeVisualStyleProperties(
            ICollection<LegacyPptWriterFoptProperty> properties,
            PowerPointShape shape) {
            if (!TryReadShapeVisualStyle(shape,
                    out IReadOnlyList<LegacyPptWriterFoptProperty> formatting,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            foreach (LegacyPptWriterFoptProperty property in formatting) {
                properties.Add(property);
            }
        }

        private static byte[]? BuildShapeFoptRecord(PowerPointShape shape,
            ushort? officeArtShapeType = null) {
            var properties = new List<LegacyPptWriterFoptProperty>();
            AddShapeFormattingProperties(properties, shape);
            if (officeArtShapeType.HasValue) {
                if (!TryReadShapeGeometry(shape, officeArtShapeType.Value,
                        out IReadOnlyList<LegacyPptWriterFoptProperty> geometry,
                        out string? reason)) {
                    throw new NotSupportedException(reason);
                }
                properties.AddRange(geometry);
            }
            return properties.Count == 0 ? null : BuildFoptRecord(properties);
        }

        private static uint GetShapeFspFlags(PowerPointShape shape) {
            if (!TryReadShapeTransform(shape,
                    out LegacyPptWriterShapeFormatting formatting,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            return formatting.FspFlags;
        }

        internal static byte[] BuildPreservedFspRecord(
            LegacyPptRecord prototype, PowerPointShape shape) {
            if (prototype == null) throw new ArgumentNullException(
                nameof(prototype));
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (prototype.Type != OfficeArtFsp
                || prototype.PayloadLength < 8) {
                throw new InvalidDataException(
                    "The preserved OfficeArt FSP atom is truncated.");
            }
            byte[] bytes = prototype.CopyRecordBytes();
            uint flags = prototype.ReadUInt32(4);
            const uint flipMask = (1U << 6) | (1U << 7);
            flags = (flags & ~flipMask) | (GetShapeFspFlags(shape)
                & flipMask);
            WriteUInt32(bytes, 12, flags);
            return bytes;
        }

        internal static byte[]? BuildPreservedShapeFoptRecord(
            LegacyPptRecord? prototype, PowerPointShape shape,
            bool rewriteShapeTransform,
            bool rewriteShapeGeometry,
            bool rewriteShapeVisualStyle,
            bool rewritePictureFormatting,
            bool rewriteTextFrame = false,
            bool rewriteShapeMetadata = false,
            bool rewriteShapeVisibility = false) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (!rewriteShapeTransform && !rewriteShapeGeometry
                && !rewriteShapeVisualStyle
                && !rewritePictureFormatting && !rewriteTextFrame
                && !rewriteShapeMetadata && !rewriteShapeVisibility) {
                throw new ArgumentException(
                    "At least one shape-property family must be rewritten.");
            }
            IReadOnlyList<LegacyPptWriterFoptProperty> sourceProperties =
                prototype == null
                    ? Array.Empty<LegacyPptWriterFoptProperty>()
                    : ReadFoptProperties(prototype);
            var properties = sourceProperties.ToList();
            if (rewriteShapeTransform) {
                properties = properties.Where(property =>
                        property.PropertyId != 0x0004
                        && property.PropertyId != 0x033F)
                    .ToList();
                AddShapeTransformProperties(properties, shape);
                PreserveBooleanPropertyBits(sourceProperties, properties,
                    0x033F, (1U << 8) | (1U << 9)
                        | (1U << 24) | (1U << 25));
            }
            if (rewriteShapeGeometry) {
                if (!TryReadOfficeArtShapeType(shape,
                        requireConnector: false, out ushort shapeType,
                        out string? shapeTypeReason)) {
                    throw new NotSupportedException(shapeTypeReason);
                }
                if (!TryReadShapeGeometry(shape, shapeType,
                        out IReadOnlyList<LegacyPptWriterFoptProperty> geometry,
                        out string? geometryReason)) {
                    throw new NotSupportedException(geometryReason);
                }
                properties = properties.Where(property =>
                        property.PropertyId != 0x0147)
                    .ToList();
                properties.AddRange(geometry);
            }
            if (rewriteShapeVisualStyle) {
                properties = properties.Where(property =>
                        property.PropertyId is not (>= 0x0180 and <= 0x023F))
                    .ToList();
                AddShapeVisualStyleProperties(properties, shape);
                PreserveBooleanPropertyBits(sourceProperties, properties,
                    0x01BF, 0x00100010U);
                PreserveBooleanPropertyBits(sourceProperties, properties,
                    0x01FF, 0x00080008U);
                PreserveBooleanPropertyBits(sourceProperties, properties,
                    0x023F, 0x00020002U);
            }
            if (rewritePictureFormatting) {
                if (shape is not PowerPointPicture picture
                    || shape is PowerPointMedia) {
                    throw new ArgumentException(
                        "Picture formatting requires a picture shape.",
                        nameof(shape));
                }
                const uint pictureBooleanMask = (1U << 18)
                    | (1U << 17) | (1U << 2) | (1U << 1);
                IReadOnlyList<LegacyPptWriterFoptProperty>
                    currentShapeProperties = properties.ToArray();
                properties = properties.Where(property =>
                        property.PropertyId != 0x007F
                        && property.PropertyId != 0x033F
                        && property.PropertyId is not (>= 0x0100 and <= 0x0103)
                        && property.PropertyId != 0x0107
                        && property.PropertyId != 0x0108
                        && property.PropertyId != 0x0109
                        && property.PropertyId != 0x013F)
                    .ToList();
                AddPictureFormatProperties(properties, picture);
                PreserveBooleanPropertyBits(sourceProperties, properties,
                    0x013F, pictureBooleanMask);
                PreserveBooleanPropertyBits(sourceProperties, properties,
                    0x007F, PictureProtectionRewriteMask);
                PreserveBooleanPropertyBits(currentShapeProperties,
                    properties, 0x033F,
                    PictureShapeBooleanRewriteMask);
            }
            if (rewriteTextFrame) {
                if (shape is not PowerPointTextBox textBox) {
                    throw new NotSupportedException(
                        "Text-frame rewriting requires a text shape.");
                }
                if (!TryReadTextFrameForWrite(textBox,
                        out IReadOnlyList<LegacyPptWriterFoptProperty> frame,
                        out string? frameReason)) {
                    throw new NotSupportedException(frameReason);
                }
                properties = properties.Where(property =>
                        property.PropertyId is not (>= 0x0081 and <= 0x0085)
                        && property.PropertyId != 0x0087
                        && property.PropertyId != 0x0088
                        && property.PropertyId != 0x00BF)
                    .ToList();
                LegacyPptWriterFoptProperty? currentFlags = frame
                    .LastOrDefault(property => property.PropertyId == 0x00BF);
                properties.AddRange(frame.Where(property =>
                    property.PropertyId != 0x00BF));
                uint rewriteMask = TextFitShapeMasks;
                A.BodyProperties body = ((P.Shape)textBox.Element)
                    .TextBody!.BodyProperties!;
                if (HasExplicitTextInsets(body)) {
                    rewriteMask |= TextAutoMarginMasks;
                }
                uint preservedFlags = (sourceProperties.LastOrDefault(
                        property => property.PropertyId == 0x00BF)?.Value
                    ?? 0U) & ~rewriteMask;
                uint flags = (currentFlags?.Value ?? 0U) | preservedFlags;
                if (flags != 0U) {
                    properties.Add(new LegacyPptWriterFoptProperty(0x00BF,
                        flags));
                }
            }
            if (rewriteShapeMetadata) {
                if (!TryReadShapeMetadataForWrite(shape,
                        out IReadOnlyList<LegacyPptWriterFoptProperty>
                            metadata, out string? metadataReason)) {
                    throw new NotSupportedException(metadataReason);
                }
                properties = properties.Where(property =>
                        property.PropertyId is not 0x0380 and not 0x0381)
                    .ToList();
                properties.AddRange(metadata);
            }
            if (rewriteShapeVisibility) {
                if (!TryReadShapeVisibilityForWrite(shape,
                        out IReadOnlyList<LegacyPptWriterFoptProperty>
                            visibility, out string? visibilityReason)) {
                    throw new NotSupportedException(visibilityReason);
                }
                IReadOnlyList<LegacyPptWriterFoptProperty>
                    currentProperties = properties.ToArray();
                properties = properties.Where(property =>
                        property.PropertyId != 0x03BF)
                    .ToList();
                properties.AddRange(visibility);
                const uint hiddenMask = (1U << 14) | (1U << 30);
                PreserveBooleanPropertyBits(currentProperties, properties,
                    0x03BF, hiddenMask);
            }
            return properties.Count == 0
                ? null
                : BuildFoptRecord(properties);
        }

        private static void PreserveBooleanPropertyBits(
            IReadOnlyList<LegacyPptWriterFoptProperty> source,
            IList<LegacyPptWriterFoptProperty> target,
            ushort propertyId, uint rewrittenMask) {
            uint preserved = (source.LastOrDefault(property =>
                    property.PropertyId == propertyId)?.Value ?? 0U)
                & ~rewrittenMask;
            int targetIndex = -1;
            for (int index = target.Count - 1; index >= 0; index--) {
                if (target[index].PropertyId != propertyId) continue;
                targetIndex = index;
                break;
            }
            if (targetIndex >= 0) {
                LegacyPptWriterFoptProperty current = target[targetIndex];
                target[targetIndex] = new LegacyPptWriterFoptProperty(
                    current.OperationId, current.Value | preserved,
                    current.ComplexData);
            } else if (preserved != 0U) {
                target.Add(new LegacyPptWriterFoptProperty(propertyId,
                    preserved));
            }
        }

        internal static bool TryReadOfficeArtShapeType(PowerPointShape shape,
            bool requireConnector, out ushort shapeType,
            out string? reason) {
            shapeType = 0;
            A.ShapeTypeValues? preset = shape switch {
                PowerPointConnectionShape connector => connector.ShapeType,
                PowerPointAutoShape autoShape => autoShape.ShapeType,
                _ when shape.Element is P.Shape source => source
                    .ShapeProperties?.GetFirstChild<A.PresetGeometry>()?
                    .Preset?.Value,
                _ => null
            };
            if (!preset.HasValue
                || !LegacyPptShapeGeometryMapper.TryGetShapeType(
                    preset.Value, out shapeType)
                || requireConnector
                    != LegacyPptShapeGeometryMapper.IsConnector(shapeType)) {
                reason = requireConnector
                    ? "The connector geometry has no classic OfficeArt connector type."
                    : "The DrawingML preset geometry has no classic OfficeArt shape type.";
                return false;
            }
            if (!TryReadShapeGeometry(shape, shapeType, out _,
                    out reason)) {
                return false;
            }
            if (requireConnector
                && (shape.Element.Descendants<A.StartConnection>().Any()
                    || shape.Element.Descendants<A.EndConnection>().Any())) {
                reason = "Fresh connector attachment rules are not yet encoded by the binary PowerPoint writer.";
                return false;
            }
            reason = null;
            return true;
        }

        internal sealed class LegacyPptWriterShapeFormatting {
            internal static LegacyPptWriterShapeFormatting Empty { get; } =
                new(Array.Empty<LegacyPptWriterFoptProperty>(),
                    0x00000A00U);

            internal LegacyPptWriterShapeFormatting(
                IReadOnlyList<LegacyPptWriterFoptProperty> properties,
                uint fspFlags) {
                Properties = properties;
                FspFlags = fspFlags;
            }

            internal IReadOnlyList<LegacyPptWriterFoptProperty> Properties {
                get;
            }

            internal uint FspFlags { get; }
        }
    }
}
