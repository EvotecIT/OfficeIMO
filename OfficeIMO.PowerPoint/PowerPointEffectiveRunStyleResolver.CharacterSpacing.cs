using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointEffectiveRunStyleResolver {
        // Zero is an authored reset and must stop inheritance just like a nonzero value.
        internal static double? ResolveCharacterSpacingPoints(PowerPointTextRun run,
            PowerPointParagraph paragraph, PowerPointTextBox textBox,
            IDictionary<int, int?>? inheritedSpacingByLevel = null) {
            int? spacing = ResolveValue(ResolveDirectSources(run, paragraph, TextBoxListStyle(textBox),
                includeParagraphMark: false), static source => source.Int32Attribute("spc"));
            if (spacing.HasValue) return spacing.Value / 100D;

            int level = ParagraphLevel(paragraph.Element.Child(Drawing, "pPr"));
            if (inheritedSpacingByLevel != null && inheritedSpacingByLevel.TryGetValue(level, out spacing))
                return spacing.HasValue ? spacing.Value / 100D : (double?)null;

            var sources = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode inherited in EnumerateTextColorShapes(textBox).Skip(1))
                AddPlaceholderDefaults(sources, inherited, level);
            sources.AddRange(FindDefaultRunProperties(TextBoxMasterStyle(textBox), level));
            sources.AddRange(FindDefaultRunProperties(PresentationTextDefaults(TextBoxSlide(textBox)), level));
            spacing = ResolveValue(sources, static source => source.Int32Attribute("spc"));
            if (inheritedSpacingByLevel != null) inheritedSpacingByLevel[level] = spacing;
            return spacing.HasValue ? spacing.Value / 100D : (double?)null;
        }
    }
}
