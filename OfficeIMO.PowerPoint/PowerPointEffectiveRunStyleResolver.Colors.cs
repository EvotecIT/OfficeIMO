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
        /// <summary>Resolves ordinary text in the consuming slide, including placeholder defaults and font-reference colors.</summary>
        internal static OfficeColor? ResolveTextColorForExport(PowerPointTextRun? run,
            PowerPointParagraph? paragraph, PowerPointTextBox textBox) {
            P.Shape[] shapes = EnumerateTextColorShapes(textBox).ToArray();
            A.FontReference? fontReference = shapes.Select(shape => shape.ShapeStyle?.FontReference)
                .FirstOrDefault(reference => reference != null);
            int level = paragraph?.Paragraph.ParagraphProperties?.Level?.Value ?? 0;
            var sources = new List<A.TextCharacterPropertiesType>();
            if (run?.RunProperties != null) sources.Add(run.RunProperties);
            if (paragraph?.Paragraph.ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>() is { } defaults)
                sources.Add(defaults);
            sources.AddRange(FindDefaultRunProperties(textBox.TextBody?.ListStyle, level));
            foreach (P.Shape shape in shapes.Skip(1)) AddPlaceholderDefaults(sources, shape, level);
            OfficeColor? color = ResolveTextSourceColors(sources, textBox, fontReference);
            if (color.HasValue) return color;

            // A shape's style supplies its text color before the presentation-wide text defaults.
            color = ResolveExportTextColor(fontReference, textBox, placeholder: null);
            if (color.HasValue) return color;
            sources.Clear();
            sources.AddRange(FindDefaultRunProperties(ResolveTextMasterStyle(textBox), level));
            SlidePart? slidePart = textBox.OwnerSlide?.SlidePart ?? textBox.SlidePart;
            P.DefaultTextStyle? presentationDefaults = slidePart?.GetParentParts()
                .OfType<PresentationPart>().FirstOrDefault()?.Presentation?.DefaultTextStyle;
            sources.AddRange(FindDefaultRunProperties(presentationDefaults, level));
            return ResolveTextSourceColors(sources, textBox, fontReference);
        }

        /// <summary>Resolves a direct text color choice, such as a highlight, with the same theme context.</summary>
        internal static OfficeColor? ResolveTextColorElementForExport(OpenXmlElement? source,
            PowerPointTextBox textBox) => source == null ? null : ResolveExportTextColor(source, textBox,
                EnumerateTextColorShapes(textBox).Select(shape => shape.ShapeStyle?.FontReference)
                    .FirstOrDefault(reference => reference != null));

        private static OfficeColor? ResolveTextSourceColors(IEnumerable<A.TextCharacterPropertiesType> sources,
            PowerPointTextBox textBox, A.FontReference? fontReference) {
            foreach (A.TextCharacterPropertiesType source in sources) {
                A.SolidFill? fill = source.GetFirstChild<A.SolidFill>();
                if (fill == null) continue;
                OfficeColor? color = ResolveExportTextColor(fill, textBox, fontReference);
                if (color.HasValue) return color;
            }
            return null;
        }

        private static OfficeColor? ResolveExportTextColor(OpenXmlElement? source,
            PowerPointTextBox textBox, OpenXmlElement? placeholder) {
            SlidePart? slidePart = textBox.OwnerSlide?.SlidePart ?? textBox.SlidePart;
            return OfficeOpenXmlThemeColorResolver.ResolveColor(source,
                textBox.OwnerSlide?.ResolvedColorScheme ?? ResolveColorScheme(slidePart), placeholder,
                PowerPointSlide.ResolveExportColorMap(slidePart));
        }

        private static OpenXmlCompositeElement? ResolveTextMasterStyle(PowerPointTextBox textBox) {
            if (textBox.MasterTextStyle != null) return textBox.MasterTextStyle;
            P.TextStyles? styles = textBox.OwnerSlide?.SlidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
            P.PlaceholderValues? type = textBox.Element.GetFirstChild<P.NonVisualShapeProperties>()?
                .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>()?.Type?.Value;
            if (type == P.PlaceholderValues.Title || type == P.PlaceholderValues.CenteredTitle) return styles?.TitleStyle;
            if (type == P.PlaceholderValues.Body || type == P.PlaceholderValues.Object || type == P.PlaceholderValues.SubTitle) return styles?.BodyStyle;
            return styles?.OtherStyle;
        }

        private static IEnumerable<P.Shape> EnumerateTextColorShapes(PowerPointTextBox textBox) =>
            textBox.Element is P.Shape shape
                ? EnumeratePlaceholderShapes(shape, (textBox.OwnerSlide?.SlidePart ?? textBox.SlidePart)?.SlideLayoutPart)
                : Enumerable.Empty<P.Shape>();
    }
}
