using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointEffectiveRunStyleResolver {
        // Zero is an authored reset and must stop inheritance just like a nonzero value.
        internal static double? ResolveCharacterSpacingPoints(PowerPointTextRun run,
            PowerPointParagraph paragraph, PowerPointTextBox textBox,
            IDictionary<int, int?>? inheritedSpacingByLevel = null) {
            int? spacing = ResolveValue(ResolveDirectSources(run, paragraph, textBox.TextBody?.ListStyle,
                includeParagraphMark: false), static source => source.Spacing?.Value);
            if (spacing.HasValue) return spacing.Value / 100D;

            int level = paragraph.Paragraph.ParagraphProperties?.Level?.Value ?? 0;
            if (inheritedSpacingByLevel != null && inheritedSpacingByLevel.TryGetValue(level, out spacing))
                return spacing.HasValue ? spacing.Value / 100D : (double?)null;

            var sources = new List<A.TextCharacterPropertiesType>();
            foreach (P.Shape inherited in EnumerateTextColorShapes(textBox).Skip(1))
                AddPlaceholderDefaults(sources, inherited, level);
            sources.AddRange(FindDefaultRunProperties(textBox.MasterTextStyle, level));
            OpenXmlCompositeElement? presentationDefaults = textBox.SlidePart?.GetParentParts()
                .OfType<DocumentFormat.OpenXml.Packaging.PresentationPart>().FirstOrDefault()?.Presentation?.DefaultTextStyle;
            sources.AddRange(FindDefaultRunProperties(presentationDefaults, level));
            spacing = ResolveValue(sources, static source => source.Spacing?.Value);
            if (inheritedSpacingByLevel != null) inheritedSpacingByLevel[level] = spacing;
            return spacing.HasValue ? spacing.Value / 100D : (double?)null;
        }

    }
}
