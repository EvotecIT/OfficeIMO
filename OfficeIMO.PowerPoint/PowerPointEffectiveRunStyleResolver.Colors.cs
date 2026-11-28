using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointEffectiveRunStyleResolver {
        /// <summary>Resolves ordinary text in the consuming slide, including placeholder defaults and font-reference colors.</summary>
        internal static OfficeColor? ResolveTextColorForExport(PowerPointTextRun? run,
            PowerPointParagraph? paragraph, PowerPointTextBox textBox) {
            OfficeXmlNode[] shapes = EnumerateTextColorShapes(textBox).ToArray();
            OfficeXmlNode fontReference = FirstFontReference(shapes);
            OfficeXmlNode paragraphProperties = paragraph?.Element.Child(Drawing, "pPr") ?? default;
            int level = ParagraphLevel(paragraphProperties);
            var sources = new List<OfficeXmlNode>(4);
            if (run?.PropertiesElement is { Exists: true } runProperties) sources.Add(runProperties);
            if (paragraphProperties.Child(Drawing, "defRPr") is { Exists: true } defaults) sources.Add(defaults);
            sources.AddRange(FindDefaultRunProperties(TextBoxListStyle(textBox), level));
            for (int index = 1; index < shapes.Length; index++) AddPlaceholderDefaults(sources, shapes[index], level);

            PowerPointThemeScope scope = PowerPointThemeScope.For(TextBoxSlide(textBox));
            OfficeColor? color = ResolveTextSourceColors(sources, scope, fontReference);
            if (color.HasValue) return color;

            // A shape's style supplies its text color before the presentation-wide text defaults.
            color = OfficeOpenXmlThemeColorResolver.ResolveColor(fontReference, scope.ColorScheme, colorMap: scope.ColorMap);
            if (color.HasValue) return color;
            sources.Clear();
            sources.AddRange(FindDefaultRunProperties(TextBoxMasterStyle(textBox), level));
            sources.AddRange(FindDefaultRunProperties(PresentationTextDefaults(TextBoxSlide(textBox)), level));
            return ResolveTextSourceColors(sources, scope, fontReference);
        }

        /// <summary>Resolves a direct text color choice, such as a highlight, with the same theme context.</summary>
        internal static OfficeColor? ResolveTextColorElementForExport(OfficeXmlNode source, PowerPointTextBox textBox) {
            if (source.IsNone) return null;
            PowerPointThemeScope scope = PowerPointThemeScope.For(TextBoxSlide(textBox));
            return OfficeOpenXmlThemeColorResolver.ResolveColor(source, scope.ColorScheme,
                FirstFontReference(EnumerateTextColorShapes(textBox)), scope.ColorMap);
        }

        private static OfficeColor? ResolveTextSourceColors(IReadOnlyList<OfficeXmlNode> sources,
            PowerPointThemeScope scope, OfficeXmlNode fontReference) {
            for (int index = 0; index < sources.Count; index++) {
                OfficeXmlNode fill = sources[index].Child(Drawing, "solidFill");
                if (fill.IsNone) continue;
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(fill, scope.ColorScheme, fontReference, scope.ColorMap);
                if (color.HasValue) return color;
            }

            return null;
        }

        private static OfficeXmlNode FirstFontReference(IEnumerable<OfficeXmlNode> shapes) {
            foreach (OfficeXmlNode shape in shapes) {
                OfficeXmlNode reference = shape.Child(Presentation, "style").Child(Drawing, "fontRef");
                if (reference.Exists) return reference;
            }

            return default;
        }

        internal static OfficeXmlNode TextBoxListStyle(PowerPointTextBox textBox) =>
            textBox.XmlElement.Child(Presentation, "txBody").Child(Drawing, "lstStyle");

        internal static OfficeXmlNode TextBoxMasterStyle(PowerPointTextBox textBox) =>
            MasterTextStyles(TextBoxSlide(textBox)).Child(Presentation, ShapePlaceholder(textBox.XmlElement).Attribute("type") switch {
                "title" or "ctrTitle" => "titleStyle",
                "body" or "obj" or "subTitle" => "bodyStyle",
                _ => "otherStyle"
            });

        internal static OfficeXmlNode MasterTextStyles(PowerPointPartRef slide) =>
            slide.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType)
                .RelatedPart(PowerPointPartRef.SlideMasterRelationshipType).Root.Child(Presentation, "txStyles");

        private static OfficeXmlNode PresentationTextDefaults(PowerPointPartRef slide) {
            PowerPointPartRef presentation = slide.Document?.PresentationPartRef ?? default;
            return presentation.RelationshipIdOf(slide, PowerPointPartRef.SlideRelationshipType) == null
                ? default : presentation.Root.Child(Presentation, "defaultTextStyle");
        }

        private static PowerPointPartRef TextBoxSlide(PowerPointTextBox textBox) => textBox.OwnerSlide?.PartRef ?? textBox.SlidePartRef;

        private static IEnumerable<OfficeXmlNode> EnumerateTextColorShapes(PowerPointTextBox textBox) =>
            EnumeratePlaceholderShapes(textBox.XmlElement, TextBoxSlide(textBox).RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType));
    }
}
