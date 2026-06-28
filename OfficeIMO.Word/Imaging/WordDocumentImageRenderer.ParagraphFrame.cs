using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static void AddParagraphFrame(WordParagraph paragraph, WordImageFlowContext context, WordImageTextLayout textLayout, double height, A.ColorScheme? colorScheme) {
            OfficeColor? fill = ResolveParagraphFillColor(paragraph, colorScheme);
            OfficeBorderBox borders = ResolveParagraphBorders(paragraph, colorScheme);
            if (!fill.HasValue && !borders.HasVisibleSide) {
                return;
            }

            context.Drawing.AddBorderBox(textLayout.TextLeft, context.Y, textLayout.TextWidth, height, fill, borders);
        }

        private static OfficeColor? ResolveParagraphFillColor(WordParagraph paragraph, A.ColorScheme? colorScheme) {
            Shading? shading = ResolveParagraphShading(paragraph);
            string? resolvedThemeColor = ResolveThemeColor(
                GetWordAttribute(shading, "themeFill"),
                GetWordAttribute(shading, "themeFillTint"),
                GetWordAttribute(shading, "themeFillShade"),
                colorScheme);
            if (TryParseOfficeColor(resolvedThemeColor, out OfficeColor themeFill)) {
                return themeFill;
            }

            string? fill = GetWordAttribute(shading, "fill");
            if (!string.IsNullOrWhiteSpace(fill) && !string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase) && TryParseOfficeColor(fill, out OfficeColor fillColor)) {
                return fillColor;
            }

            return shading == paragraph._paragraphProperties?.Shading ? paragraph.ShadingFillColor : null;
        }

        private static OfficeBorderBox ResolveParagraphBorders(WordParagraph paragraph, A.ColorScheme? colorScheme) {
            ParagraphBorders? directBorders = paragraph._paragraphProperties?.GetFirstChild<ParagraphBorders>();
            List<ParagraphBorders> inheritedBorders = ResolveParagraphStyleBorders(paragraph).ToList();
            return new OfficeBorderBox(
                ResolveParagraphBorderSide(directBorders?.LeftBorder ?? FirstInheritedBorderSide(inheritedBorders, borders => borders.LeftBorder), colorScheme),
                ResolveParagraphBorderSide(directBorders?.TopBorder ?? FirstInheritedBorderSide(inheritedBorders, borders => borders.TopBorder), colorScheme),
                ResolveParagraphBorderSide(directBorders?.RightBorder ?? FirstInheritedBorderSide(inheritedBorders, borders => borders.RightBorder), colorScheme),
                ResolveParagraphBorderSide(directBorders?.BottomBorder ?? FirstInheritedBorderSide(inheritedBorders, borders => borders.BottomBorder), colorScheme));
        }

        private static Shading? ResolveParagraphShading(WordParagraph paragraph) =>
            paragraph._paragraphProperties?.Shading
            ?? EnumerateParagraphStyleProperties(paragraph).Select(properties => properties.Shading).FirstOrDefault(shading => shading != null);

        private static IEnumerable<ParagraphBorders> ResolveParagraphStyleBorders(WordParagraph paragraph) {
            foreach (StyleParagraphProperties properties in EnumerateParagraphStyleProperties(paragraph)) {
                ParagraphBorders? borders = properties.GetFirstChild<ParagraphBorders>();
                if (borders != null) {
                    yield return borders;
                }
            }
        }

        private static OpenXmlElement? FirstInheritedBorderSide(IEnumerable<ParagraphBorders> inheritedBorders, Func<ParagraphBorders, OpenXmlElement?> selector) =>
            inheritedBorders.Select(selector).FirstOrDefault(side => side != null);

        private static OfficeBorderSide? ResolveParagraphBorderSide(OpenXmlElement? source, A.ColorScheme? colorScheme) {
            if (source == null) {
                return null;
            }

            BorderValues? style = ResolveParagraphBorderStyle(source);
            if (!style.HasValue || style == BorderValues.Nil || style == BorderValues.None) {
                return null;
            }

            double width = ResolveParagraphBorderWidth(GetWordAttribute(source, "sz"));
            return new OfficeBorderSide(
                ResolveParagraphBorderColor(source, colorScheme),
                width,
                MapBorderDashStyle(style),
                style == BorderValues.Double ? OfficeBorderLineKind.Double : OfficeBorderLineKind.Single,
                style == BorderValues.Double ? Math.Max(1.5D, width * 3D) : 0D);
        }

        private static BorderValues? ResolveParagraphBorderStyle(OpenXmlElement source) {
            string? value = GetWordAttribute(source, "val");
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            return value switch {
                "single" => BorderValues.Single,
                "double" => BorderValues.Double,
                "dashed" => BorderValues.Dashed,
                "dashSmallGap" => BorderValues.DashSmallGap,
                "dotted" => BorderValues.Dotted,
                "dotDash" => BorderValues.DotDash,
                "dotDotDash" => BorderValues.DotDotDash,
                "nil" => BorderValues.Nil,
                "none" => BorderValues.None,
                _ => null
            };
        }

        private static OfficeColor ResolveParagraphBorderColor(OpenXmlElement source, A.ColorScheme? colorScheme) {
            string? resolvedThemeColor = ResolveThemeColor(
                GetWordAttribute(source, "themeColor"),
                GetWordAttribute(source, "themeTint"),
                GetWordAttribute(source, "themeShade"),
                colorScheme);
            if (TryParseOfficeColor(resolvedThemeColor, out OfficeColor themeColor)) {
                return themeColor;
            }

            string? color = GetWordAttribute(source, "color");
            if (!string.IsNullOrWhiteSpace(color) && !string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase)) {
                try {
                    return Helpers.ParseColor(color!);
                } catch (ArgumentException) {
                    return OfficeColor.LightGray;
                }
            }

            return OfficeColor.LightGray;
        }

        private static double ResolveParagraphBorderWidth(string? borderSize) {
            if (!uint.TryParse(borderSize, out uint size) || size == 0U) {
                return 0.75D;
            }

            return Math.Max(0.5D, size / 8D);
        }
    }
}
