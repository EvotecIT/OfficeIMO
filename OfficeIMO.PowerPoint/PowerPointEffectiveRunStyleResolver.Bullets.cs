using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointEffectiveRunStyleResolver {
        /// <summary>Resolves bullet paint separately from the body, honoring an explicit return to text color.</summary>
        internal static OfficeColor ResolveBulletColorForExport(PowerPointParagraph paragraph,
            PowerPointTextBox textBox, OfficeColor textColor) {
            int level = paragraph.Paragraph.ParagraphProperties?.Level?.Value ?? 0;
            var sources = new List<A.TextParagraphPropertiesType?> { paragraph.Paragraph.ParagraphProperties };
            sources.AddRange(FindParagraphProperties(textBox.TextBody?.ListStyle, level));
            foreach (P.Shape shape in EnumerateTextColorShapes(textBox).Skip(1)) {
                sources.Add(shape.TextBody?.Elements<A.Paragraph>()
                    .FirstOrDefault(p => (p.ParagraphProperties?.Level?.Value ?? 0) == level)?.ParagraphProperties);
                sources.AddRange(FindParagraphProperties(shape.TextBody?.ListStyle, level));
            }
            sources.AddRange(FindParagraphProperties(ResolveTextMasterStyle(textBox), level));
            SlidePart? slidePart = textBox.OwnerSlide?.SlidePart ?? textBox.SlidePart;
            var defaults = slidePart?.GetParentParts().OfType<PresentationPart>()
                .FirstOrDefault()?.Presentation?.DefaultTextStyle;
            sources.AddRange(FindParagraphProperties(defaults, level));
            return ResolveBulletColor(sources, color => ResolveTextColorElementForExport(color, textBox), textColor);
        }

        /// <summary>Resolves table-cell paragraph and list defaults in the consuming slide's color map.</summary>
        internal static OfficeColor ResolveBulletColorForExport(PowerPointParagraph paragraph,
            PowerPointTableCell cell, A.ColorScheme? colorScheme, OfficeColor textColor) {
            int level = paragraph.Paragraph.ParagraphProperties?.Level?.Value ?? 0;
            var sources = new List<A.TextParagraphPropertiesType?> { paragraph.Paragraph.ParagraphProperties };
            sources.AddRange(FindParagraphProperties(cell.Cell.TextBody?.ListStyle, level));
            return ResolveBulletColor(sources, color => OfficeOpenXmlThemeColorResolver.ResolveColor(color,
                colorScheme, colorMap: PowerPointSlide.ResolveExportColorMap(cell.OwnerPart as SlidePart)), textColor);
        }

        private static OfficeColor ResolveBulletColor(IEnumerable<A.TextParagraphPropertiesType?> sources,
            Func<OpenXmlElement, OfficeColor?> resolve, OfficeColor textColor) {
            foreach (A.TextParagraphPropertiesType? source in sources) {
                if (source?.GetFirstChild<A.BulletColorText>() != null) return textColor;
                if (source?.GetFirstChild<A.BulletColor>() is { } bullet && resolve(bullet) is { } color) return color;
            }
            return textColor;
        }
    }
}
