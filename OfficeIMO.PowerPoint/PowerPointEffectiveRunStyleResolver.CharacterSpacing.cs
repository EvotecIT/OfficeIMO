using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointEffectiveRunStyleResolver {
        // Zero is an authored reset and must stop inheritance just like a nonzero value.
        internal static double? ResolveCharacterSpacingPoints(PowerPointTextRun run,
            PowerPointParagraph paragraph, PowerPointTextBox textBox) {
            var sources = new List<A.TextCharacterPropertiesType>(
                ResolveDirectSources(run, paragraph, textBox.TextBody?.ListStyle, includeParagraphMark: false));
            int level = paragraph.Paragraph.ParagraphProperties?.Level?.Value ?? 0;
            P.PlaceholderShape? placeholder = textBox.Element.GetFirstChild<P.NonVisualShapeProperties>()?
                .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
            if (placeholder != null) {
                var layout = textBox.SlidePart?.SlideLayoutPart;
                P.Shape? layoutShape = FindSpacingPlaceholder(layout?.SlideLayout?.CommonSlideData?.ShapeTree,
                    placeholder, matchIndex: true);
                AddPlaceholderDefaults(sources, layoutShape, level);
                P.PlaceholderShape effectivePlaceholder = layoutShape?.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() ?? placeholder;
                AddPlaceholderDefaults(sources, FindSpacingPlaceholder(
                    layout?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree,
                    effectivePlaceholder, matchIndex: false), level);
            }
            sources.AddRange(FindDefaultRunProperties(textBox.MasterTextStyle, level));
            OpenXmlCompositeElement? presentationDefaults = textBox.SlidePart?.GetParentParts()
                .OfType<DocumentFormat.OpenXml.Packaging.PresentationPart>().FirstOrDefault()?.Presentation?.DefaultTextStyle;
            sources.AddRange(FindDefaultRunProperties(presentationDefaults, level));
            int? spacing = ResolveValue(sources, static source => source.Spacing?.Value);
            return spacing.HasValue ? spacing.Value / 100D : (double?)null;
        }

        private static void AddPlaceholderDefaults(List<A.TextCharacterPropertiesType> sources, P.Shape? shape, int level) {
            A.DefaultRunProperties? paragraphDefaults = shape?.TextBody?.Elements<A.Paragraph>()
                .FirstOrDefault(paragraph => (paragraph.ParagraphProperties?.Level?.Value ?? 0) == level)?
                .ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>();
            if (paragraphDefaults != null) sources.Add(paragraphDefaults);
            sources.AddRange(FindDefaultRunProperties(shape?.TextBody?.ListStyle, level));
        }

        private static P.Shape? FindSpacingPlaceholder(P.ShapeTree? tree, P.PlaceholderShape source, bool matchIndex) {
            IEnumerable<P.Shape> matches = tree?.Elements<P.Shape>() ?? Enumerable.Empty<P.Shape>();
            matches = matches.Where(shape => {
                P.PlaceholderShape? candidate = shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
                return candidate != null && (matchIndex
                    ? (candidate.Index?.Value ?? 0U) == (source.Index?.Value ?? 0U)
                    : NormalizePlaceholderType(candidate.Type?.Value) == NormalizePlaceholderType(source.Type?.Value));
            });
            // Ambiguous source declarations cannot establish an effective inherited value.
            P.Shape[] selected = matches.Take(2).ToArray();
            return selected.Length == 1 ? selected[0] : null;
        }

        private static P.PlaceholderValues NormalizePlaceholderType(P.PlaceholderValues? value) =>
            value == P.PlaceholderValues.CenteredTitle ? P.PlaceholderValues.Title : value ?? P.PlaceholderValues.Object;
    }
}
