using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using System.Globalization;
using System.Text;
using static OfficeIMO.PowerPoint.LegacyPpt.Write.LegacyPptWriterXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static bool TryReadShapeFormatting(PowerPointShape shape,
            out LegacyPptWriterShapeFormatting formatting,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(shape);
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
            ArgumentNullException.ThrowIfNull(shape);
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
            ArgumentNullException.ThrowIfNull(shape);
            return TryReadShapeMetadataForWrite(shape.Title, shape.Decorative,
                shape.Name, shape.AltText, out properties, out reason);
        }

        private static bool TryReadShapeMetadataForWrite(string? title,
            bool decorative, string? name, string? altText,
            out IReadOnlyList<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (!string.IsNullOrWhiteSpace(title)) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "Accessibility titles have no native PowerPoint 97-2003 shape-property representation.";
                return false;
            }
            if (decorative) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "The modern decorative-shape flag has no native PowerPoint 97-2003 representation.";
                return false;
            }
            var result = new List<LegacyPptWriterFoptProperty>(2);
            AddShapeMetadataText(result, 0x0380, name);
            AddShapeMetadataText(result, 0x0381, altText);
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
            ArgumentNullException.ThrowIfNull(shape);
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
            ArgumentNullException.ThrowIfNull(shape);
            OfficeXmlNode geometry = FirstDescendant(ShapeXml(shape),
                LegacyPptXml.A, "prstGeom");
            OfficeXmlNode values = geometry.Child(LegacyPptXml.A, "avLst");
            if (values.IsNone || !HasChildren(values)) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = null;
                return true;
            }
            if (HasAttributes(values)
                || AnyChildOtherThan(values, LegacyPptXml.A, "gd")
                || values.Children(LegacyPptXml.A, "gd").Count > 1
                || shapeType is not 2 and not 23) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "Only one exact round-rectangle or donut adjustment can currently be written to binary PowerPoint.";
                return false;
            }
            OfficeXmlNode guide = values.Child(LegacyPptXml.A, "gd");
            if (HasChildren(guide)
                || HasAttributeOtherThan(guide, "name", "fmla")
                || !string.Equals(guide.Attribute("name"), "adj",
                    StringComparison.Ordinal)
                || !TryReadConstantGuide(guide.Attribute("fmla"),
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
            ArgumentNullException.ThrowIfNull(shape);
            return TryReadShapeVisualStyle(ShapeXml(shape), out properties,
                out reason);
        }

        internal static bool TryReadShapeVisualStyle(OfficeXmlNode element,
            out IReadOnlyList<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (!element.Child(LegacyPptXml.P, "style").IsNone) {
                properties = Array.Empty<LegacyPptWriterFoptProperty>();
                reason = "Theme-referenced shape styles must be resolved before binary PowerPoint writing.";
                return false;
            }
            var result = new List<LegacyPptWriterFoptProperty>();
            OfficeXmlNode shapeProperties = element.Is(LegacyPptXml.P, "sp")
                || element.Is(LegacyPptXml.P, "cxnSp")
                || element.Is(LegacyPptXml.P, "pic")
                ? element.Child(LegacyPptXml.P, "spPr")
                : default;
            if (!shapeProperties.IsNone
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
            OfficeXmlNode source,
            ICollection<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            foreach (OfficeXmlNode child in source.Children()) {
                if (child.Is(LegacyPptXml.A, "xfrm")
                    || child.Is(LegacyPptXml.A, "prstGeom")
                    || child.Is(LegacyPptXml.A, "noFill")
                    || child.Is(LegacyPptXml.A, "solidFill")
                    || child.Is(LegacyPptXml.A, "ln")
                    || child.Is(LegacyPptXml.A, "effectLst")) {
                    continue;
                }
                reason = $"The shape property '{child.LocalName}' has no lossless classic OfficeArt mapping.";
                return false;
            }

            List<OfficeXmlNode> fills = FillChildren(source);
            if (fills.Count > 1) {
                reason = "The shape contains multiple explicit fills.";
                return false;
            }
            if (fills.Count == 1) {
                if (fills[0].Is(LegacyPptXml.A, "noFill")) {
                    if (HasAttributes(fills[0]) || HasChildren(fills[0])) {
                        reason = "The no-fill shape property contains an unsupported extension.";
                        return false;
                    }
                    properties.Add(new LegacyPptWriterFoptProperty(0x01BF,
                        0x00100000U));
                } else if (!TryReadRgbFill(fills[0],
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

            OfficeXmlNode outline = source.Child(LegacyPptXml.A, "ln");
            if (!outline.IsNone
                && !TryReadOutline(outline, properties, out reason)) {
                return false;
            }
            OfficeXmlNode effects = source.Child(LegacyPptXml.A, "effectLst");
            if (!effects.IsNone
                && !TryReadShapeEffects(effects, properties, out reason)) {
                return false;
            }
            reason = null;
            return true;
        }

        private static List<OfficeXmlNode> FillChildren(OfficeXmlNode parent) {
            var fills = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode child in parent.Children()) {
                if (child.Is(LegacyPptXml.A, "noFill")
                    || child.Is(LegacyPptXml.A, "solidFill")) {
                    fills.Add(child);
                }
            }
            return fills;
        }

        private static bool TryReadOutline(OfficeXmlNode outline,
            ICollection<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (HasAttributeOtherThan(outline, "w", "cap")) {
                reason = "Binary PowerPoint does not support the requested compound or aligned outline.";
                return false;
            }
            List<OfficeXmlNode> fills = FillChildren(outline);
            if (fills.Count > 1) {
                reason = "The shape outline contains multiple explicit fills.";
                return false;
            }
            if (fills.Count == 0 && (HasAttributes(outline)
                    || HasChildren(outline))) {
                reason = "An explicitly formatted outline requires an RGB solid fill or no-fill for binary PowerPoint.";
                return false;
            }
            if (fills.Count == 1 && fills[0].Is(LegacyPptXml.A, "noFill")) {
                if (HasAttributes(fills[0]) || HasChildren(fills[0])) {
                    reason = "The no-fill outline contains an unsupported extension.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(0x01FF,
                    0x00080000U));
            } else if (fills.Count == 1) {
                if (!TryReadRgbFill(fills[0],
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
            if (outline.Int("w") is int width) {
                if (width < 0) {
                    reason = "Shape outline width cannot be negative.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(0x01CB,
                    checked((uint)width)));
            }
            if (outline.Attribute("cap") is string cap) {
                uint value = cap switch {
                    "rnd" => 0U,
                    "sq" => 1U,
                    "flat" => 2U,
                    _ => uint.MaxValue
                };
                if (value == uint.MaxValue) {
                    reason = "The outline cap has no classic OfficeArt mapping.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(0x01D7,
                    value));
            }

            foreach (OfficeXmlNode child in outline.Children()) {
                string name = child.NamespaceId == LegacyPptXml.A
                    ? child.LocalName : string.Empty;
                switch (name) {
                    case "noFill":
                    case "solidFill":
                        break;
                    case "prstDash":
                        if (HasChildren(child)
                            || HasAttributeOtherThan(child, "val")
                            || !TryMapLineDash(child.Attribute("val"),
                                out uint dashValue)) {
                            reason = "The outline dash pattern has no classic OfficeArt mapping.";
                            return false;
                        }
                        properties.Add(new LegacyPptWriterFoptProperty(0x01CE,
                            dashValue));
                        break;
                    case "bevel":
                        if (HasAttributes(child) || HasChildren(child)) {
                            reason = "The bevel line join contains an unsupported extension.";
                            return false;
                        }
                        properties.Add(new LegacyPptWriterFoptProperty(0x01D6,
                            0U));
                        break;
                    case "miter":
                        if (HasAttributes(child) || HasChildren(child)) {
                            reason = "A custom miter limit has no exact classic OfficeArt mapping.";
                            return false;
                        }
                        properties.Add(new LegacyPptWriterFoptProperty(0x01D6,
                            1U));
                        break;
                    case "round":
                        if (HasAttributes(child) || HasChildren(child)) {
                            reason = "The round line join contains an unsupported extension.";
                            return false;
                        }
                        properties.Add(new LegacyPptWriterFoptProperty(0x01D6,
                            2U));
                        break;
                    case "headEnd":
                        if (!TryAddLineEnd(child, isHead: true, properties,
                                out reason)) return false;
                        break;
                    case "tailEnd":
                        if (!TryAddLineEnd(child, isHead: false, properties,
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

        private static bool TryReadShapeEffects(OfficeXmlNode effects,
            ICollection<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            OfficeXmlNode shadow = effects.FirstChild;
            if (HasAttributes(effects) || ChildCount(effects) > 1
                || !shadow.Is(LegacyPptXml.A, "outerShdw")) {
                if (!HasChildren(effects) && !HasAttributes(effects)) {
                    reason = null;
                    return true;
                }
                reason = "Only one classic outer shadow can be written as a native binary PowerPoint shape effect.";
                return false;
            }
            OfficeXmlNode rgb = shadow.Child(LegacyPptXml.A, "srgbClr");
            if (HasAttributeOtherThan(shadow, "blurRad", "dist", "dir",
                    "rotWithShape")
                || shadow.Bool("rotWithShape") == true
                || ChildCount(shadow) != 1
                || rgb.IsNone
                || !TryReadRgbColor(rgb, out OfficeColor color,
                    out int? alpha)) {
                reason = "The outer shadow uses scaling, skew, alignment, rotation, or color transforms that have no exact classic OfficeArt mapping.";
                return false;
            }
            long distance = shadow.Long("dist") ?? 0L;
            long blur = shadow.Long("blurRad") ?? 0L;
            if (distance < 0L || blur < 0L) {
                reason = "Outer-shadow distance and blur radius cannot be negative.";
                return false;
            }
            int direction = shadow.Int("dir") ?? 0;
            double radians = direction / 60000D * Math.PI / 180D;
            double roundedOffsetX = Math.Round(Math.Cos(radians)
                * distance, MidpointRounding.AwayFromZero);
            double roundedOffsetY = Math.Round(Math.Sin(radians)
                * distance, MidpointRounding.AwayFromZero);
            if (double.IsNaN(roundedOffsetX) || double.IsInfinity(roundedOffsetX)
                || double.IsNaN(roundedOffsetY) || double.IsInfinity(roundedOffsetY)
                || roundedOffsetX < int.MinValue || roundedOffsetX > int.MaxValue
                || roundedOffsetY < int.MinValue || roundedOffsetY > int.MaxValue
                || blur > int.MaxValue) {
                reason = "The outer shadow exceeds the classic OfficeArt coordinate range.";
                return false;
            }
            int offsetX = (int)roundedOffsetX;
            int offsetY = (int)roundedOffsetY;
            properties.Add(new LegacyPptWriterFoptProperty(0x0200, 0U));
            properties.Add(new LegacyPptWriterFoptProperty(0x0201,
                PackOfficeArtColor(color)));
            if (alpha.HasValue) {
                properties.Add(new LegacyPptWriterFoptProperty(0x0204,
                    ToOfficeArtOpacity(alpha.Value)));
            }
            properties.Add(new LegacyPptWriterFoptProperty(0x0205,
                unchecked((uint)offsetX)));
            properties.Add(new LegacyPptWriterFoptProperty(0x0206,
                unchecked((uint)offsetY)));
            if (blur > 0L) {
                properties.Add(new LegacyPptWriterFoptProperty(0x021C,
                    checked((uint)blur)));
            }
            properties.Add(new LegacyPptWriterFoptProperty(0x023F,
                0x00020002U));
            reason = null;
            return true;
        }

        private static bool TryReadRgbFill(OfficeXmlNode fill,
            out OfficeColor color, out int? alpha, out string? reason) {
            OfficeXmlNode rgb = fill.Child(LegacyPptXml.A, "srgbClr");
            if (HasAttributes(fill) || ChildCount(fill) != 1
                || rgb.IsNone
                || !TryReadRgbColor(rgb, out color, out alpha)) {
                color = default;
                alpha = null;
                reason = "Binary PowerPoint shape fills require one RGB color with an optional alpha value.";
                return false;
            }
            reason = null;
            return true;
        }

        private static bool TryReadRgbColor(OfficeXmlNode rgb,
            out OfficeColor color, out int? alpha) {
            color = default;
            alpha = null;
            string? value = rgb.Attribute("val");
            if (value == null
                || !OfficeColor.TryParse(value, out color)
                || HasExtendedAttributes(rgb, "val")
                || AnyChildOtherThan(rgb, LegacyPptXml.A, "alpha")
                || rgb.Children(LegacyPptXml.A, "alpha").Count > 1) {
                return false;
            }
            OfficeXmlNode sourceAlpha = rgb.Child(LegacyPptXml.A, "alpha");
            if (sourceAlpha.IsNone) return true;
            if (HasChildren(sourceAlpha)
                || HasExtendedAttributes(sourceAlpha, "val")
                || sourceAlpha.Int("val") is not int alphaValue
                || alphaValue < 0 || alphaValue > 100000) {
                return false;
            }
            alpha = alphaValue;
            return true;
        }

        private static bool TryAddLineEnd(OfficeXmlNode source,
            bool isHead,
            ICollection<LegacyPptWriterFoptProperty> properties,
            out string? reason) {
            if (HasChildren(source)
                || HasAttributeOtherThan(source, "type", "w", "len")) {
                reason = "The line end contains an unsupported extension.";
                return false;
            }
            string? type = source.Attribute("type");
            string? width = source.Attribute("w");
            string? length = source.Attribute("len");
            ushort typeId = isHead ? (ushort)0x01D0 : (ushort)0x01D1;
            ushort widthId = isHead ? (ushort)0x01D2 : (ushort)0x01D4;
            ushort lengthId = isHead ? (ushort)0x01D3 : (ushort)0x01D5;
            if (type != null) {
                uint value = type switch {
                    "none" => 0U,
                    "triangle" => 1U,
                    "stealth" => 2U,
                    "diamond" => 3U,
                    "oval" => 4U,
                    "arrow" => 5U,
                    _ => uint.MaxValue
                };
                if (value == uint.MaxValue) {
                    reason = "The line-end type has no classic OfficeArt mapping.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(typeId, value));
            }
            if (width != null) {
                uint value = MapLineEndSize(width);
                if (value == uint.MaxValue) {
                    reason = "The line-end width has no classic OfficeArt mapping.";
                    return false;
                }
                properties.Add(new LegacyPptWriterFoptProperty(widthId,
                    value));
            }
            if (length != null) {
                uint value = MapLineEndSize(length);
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

        private static uint MapLineEndSize(string value) => value switch {
            "sm" => 0U,
            "med" => 1U,
            "lg" => 2U,
            _ => uint.MaxValue
        };

        private static bool TryMapLineDash(string? value,
            out uint mapped) {
            mapped = value switch {
                "solid" => 0U,
                "sysDash" => 1U,
                "sysDot" => 2U,
                "sysDashDot" => 3U,
                "sysDashDotDot" => 4U,
                "dot" => 5U,
                "dash" => 6U,
                "lgDash" => 7U,
                "dashDot" => 8U,
                "lgDashDot" => 9U,
                "lgDashDotDot" => 10U,
                _ => uint.MaxValue
            };
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
            ArgumentNullException.ThrowIfNull(prototype);
            ArgumentNullException.ThrowIfNull(shape);
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
            ArgumentNullException.ThrowIfNull(shape);
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
                if (shape is not PowerPointImage picture
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
                AddPictureFormatProperties(properties, ShapeXml(picture));
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
                OfficeXmlNode body = ShapeXml(textBox)
                    .Child(LegacyPptXml.P, "txBody")
                    .Child(LegacyPptXml.A, "bodyPr");
                if (HasExplicitTextInsetAttributes(body)) {
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
            OfficePresetShapeType? preset = shape switch {
                PowerPointConnectionShape connector => connector.ShapeType,
                PowerPointAutoShape autoShape => autoShape.ShapeType,
                _ => ReadPresetGeometry(ShapeXml(shape))
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
                && (!FirstDescendant(ShapeXml(shape), LegacyPptXml.A,
                        "stCxn").IsNone
                    || !FirstDescendant(ShapeXml(shape), LegacyPptXml.A,
                        "endCxn").IsNone)) {
                reason = "Fresh connector attachment rules are not yet encoded by the binary PowerPoint writer.";
                return false;
            }
            reason = null;
            return true;
        }

        private static OfficePresetShapeType? ReadPresetGeometry(
            OfficeXmlNode element) {
            if (!element.Is(LegacyPptXml.P, "sp")) return null;
            string? token = element.Child(LegacyPptXml.P, "spPr")
                .Child(LegacyPptXml.A, "prstGeom").Attribute("prst");
            return token != null
                && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(
                    token, out OfficePresetShapeType value)
                ? value
                : null;
        }

        // The same test as LegacyPptWriter.TextFrames' HasExplicitTextInsets: an inset attribute with a valid value.
        private static bool HasExplicitTextInsetAttributes(OfficeXmlNode body) =>
            body.Int("lIns").HasValue || body.Int("tIns").HasValue
            || body.Int("rIns").HasValue || body.Int("bIns").HasValue;

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

    /// <summary>
    /// Reading helpers the binary writer uses on drawing markup: element and attribute tests that match how the markup
    /// is written, shape markup that can be enumerated in either document model, and stable shape identities for the
    /// write catalogs.
    /// </summary>
    internal static class LegacyPptWriterXml {
        /// <summary>
        /// The shape's markup as an element of an Open XML layer tree: the package element itself, or a detached copy of
        /// the SDK element while the SDK holds the document. Read it; don't compare it by identity.
        /// </summary>
        internal static OfficeXmlNode ShapeXml(PowerPointShape shape) {
            ArgumentNullException.ThrowIfNull(shape);
            return Readable(shape.XmlElement);
        }

        /// <summary>The element itself when it lives in a layer tree, otherwise a detached copy parsed from its markup.</summary>
        internal static OfficeXmlNode Readable(OfficeXmlNode element) {
            if (element.IsNone || element.Tree != null) return element;
            byte[] markup = element.OuterXml();
            OfficeXmlTree tree = OfficeXmlTree.Parse(markup, markup.Length,
                new OfficeXmlNamespaces.Table());
            return new OfficeXmlNode(tree, tree.Root);
        }

        internal static OfficeXmlNode FirstDescendant(OfficeXmlNode element,
            int namespaceId, string localName) {
            foreach (OfficeXmlNode descendant in element.AllDescendants()) {
                if (descendant.Is(namespaceId, localName)) return descendant;
            }
            return default;
        }

        internal static bool HasDescendant(OfficeXmlNode element,
            int namespaceId, string localName) =>
            !FirstDescendant(element, namespaceId, localName).IsNone;

        /// <summary>Whether the element has child elements.</summary>
        internal static bool HasChildren(OfficeXmlNode element) =>
            !element.FirstChild.IsNone;

        internal static int ChildCount(OfficeXmlNode element) =>
            element.Children().Count;

        internal static bool AnyChildOtherThan(OfficeXmlNode element,
            int namespaceId, string localName) {
            foreach (OfficeXmlNode child in element.Children()) {
                if (!child.Is(namespaceId, localName)) return true;
            }
            return false;
        }

        /// <summary>
        /// The element's attributes as markup states them, namespace declarations and markup-compatibility attributes
        /// left out.
        /// </summary>
        internal static List<(int NamespaceId, string LocalName, string Value)> Attributes(
            OfficeXmlNode element) {
            var attributes = new List<(int, string, string)>();
            OfficeXmlNode readable = Readable(element);
            if (readable.Tree is not { } tree) return attributes;
            int count = tree.AttributeCountOf(readable.Node);
            for (int index = 0; index < count; index++) {
                tree.GetAttributeAt(readable.Node, index, out int namespaceId,
                    out string localName, out string value);
                if (namespaceId == OfficeXmlNamespaces.Xmlns
                    || namespaceId == OfficeXmlNamespaces.MarkupCompatibility
                    || namespaceId == OfficeXmlNamespaces.None
                        && string.Equals(localName, "xmlns", StringComparison.Ordinal)) {
                    continue;
                }
                attributes.Add((namespaceId, localName, value));
            }
            return attributes;
        }

        /// <summary>The value of the attribute with this namespace URI and local name, or null.</summary>
        internal static string? AttributeIn(OfficeXmlNode element,
            string namespaceUri, string localName) {
            OfficeXmlNode readable = Readable(element);
            if (readable.Tree is not { } tree) return null;
            int count = tree.AttributeCountOf(readable.Node);
            for (int index = 0; index < count; index++) {
                tree.GetAttributeAt(readable.Node, index, out int namespaceId,
                    out string name, out string value);
                if (string.Equals(name, localName, StringComparison.Ordinal)
                    && string.Equals(tree.Namespaces.GetUri(namespaceId),
                        namespaceUri, StringComparison.Ordinal)) return value;
            }
            return null;
        }

        internal static bool HasAttributes(OfficeXmlNode element) =>
            Attributes(element).Count > 0;

        /// <summary>Whether an attribute has a local name outside <paramref name="localNames"/>, in any namespace.</summary>
        internal static bool HasAttributeOtherThan(OfficeXmlNode element,
            params string[] localNames) {
            foreach ((_, string localName, _) in Attributes(element)) {
                if (Array.IndexOf(localNames, localName) < 0) return true;
            }
            return false;
        }

        /// <summary>
        /// Whether the element carries an attribute its schema does not define: anything other than the unqualified
        /// <paramref name="knownLocalNames"/>.
        /// </summary>
        internal static bool HasExtendedAttributes(OfficeXmlNode element,
            params string[] knownLocalNames) {
            foreach ((int namespaceId, string localName, _) in Attributes(element)) {
                if (namespaceId != OfficeXmlNamespaces.None
                    || Array.IndexOf(knownLocalNames, localName) < 0) return true;
            }
            return false;
        }

        /// <summary>
        /// A key that identifies a shape's element for the length of a write, whichever model holds the document: its
        /// part and position there, or the element itself when it is not attached to a part.
        /// </summary>
        internal static object ShapeKey(PowerPointShape shape) {
            ArgumentNullException.ThrowIfNull(shape);
            OfficeXmlNode element = shape.XmlElement;
            if (element.IsNone) return shape;
            if (!element.IsAttached) return element;
            return ElementPath(shape.OwnerPartRef.Name, element);
        }

        /// <summary>The part name and the element's child positions from the part root.</summary>
        internal static string ElementPath(string partName, OfficeXmlNode element) {
            var positions = new Stack<int>();
            for (OfficeXmlNode current = element; !current.Parent.IsNone;
                 current = current.Parent) {
                int index = 0;
                for (OfficeXmlNode sibling = current.PreviousSibling;
                     !sibling.IsNone; sibling = sibling.PreviousSibling) {
                    index++;
                }
                positions.Push(index);
            }
            var path = new StringBuilder(partName);
            while (positions.Count > 0) {
                path.Append('/').Append(positions.Pop().ToString(
                    CultureInfo.InvariantCulture));
            }
            return path.ToString();
        }
    }
}
