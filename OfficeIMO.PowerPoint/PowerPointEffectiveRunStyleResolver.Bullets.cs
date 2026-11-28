using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointEffectiveRunStyleResolver {
        /// <summary>Resolves bullet paint separately from the body, honoring an explicit return to text color.</summary>
        internal static OfficeColor ResolveBulletColorForExport(PowerPointParagraph paragraph,
            PowerPointTextBox textBox, OfficeColor textColor) {
            OfficeXmlNode properties = paragraph.Element.Child(Drawing, "pPr");
            int level = ParagraphLevel(properties);
            var sources = new List<OfficeXmlNode> { properties };
            sources.AddRange(FindParagraphProperties(TextBoxListStyle(textBox), level));
            foreach (OfficeXmlNode shape in EnumerateTextColorShapes(textBox).Skip(1)) {
                OfficeXmlNode body = shape.Child(Presentation, "txBody");
                sources.Add(PlaceholderParagraphProperties(body, level));
                sources.AddRange(FindParagraphProperties(body.Child(Drawing, "lstStyle"), level));
            }
            sources.AddRange(FindParagraphProperties(TextBoxMasterStyle(textBox), level));
            sources.AddRange(FindParagraphProperties(PresentationTextDefaults(TextBoxSlide(textBox)), level));
            return ResolveBulletColor(sources, color => ResolveTextColorElementForExport(color, textBox), textColor);
        }

        private static OfficeColor ResolveBulletColor(IEnumerable<OfficeXmlNode> sources,
            Func<OfficeXmlNode, OfficeColor?> resolve, OfficeColor textColor) {
            foreach (OfficeXmlNode source in sources) {
                if (source.Child(Drawing, "buClrTx").Exists) return textColor;
                if (source.Child(Drawing, "buClr") is { Exists: true } bullet && resolve(bullet) is { } color) return color;
            }

            return textColor;
        }
    }
}
