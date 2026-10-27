using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace OfficeIMO.Word {
    /// <summary>
    /// Word's built-in table-of-contents paragraph styles. Generated TOCs reference
    /// <c>TOCHeading</c> and <c>TOC1</c>–<c>TOC9</c>; without definitions their
    /// title and entries fall back to body text in every viewer.
    /// </summary>
    internal static class WordTableOfContentStyles {
        private const int EntryIndentTwips = 220;
        private const int EntryLevels = 9;

        internal static IEnumerable<Style> Create() {
            yield return CreateHeading();
            for (int level = 1; level <= EntryLevels; level++) {
                yield return CreateEntry(level);
            }
        }

        /// <summary>Adds any missing TOC style to a style part, leaving existing definitions untouched.</summary>
        internal static void EnsureDefined(StyleDefinitionsPart? part) {
            Styles? styles = part?.Styles;
            if (styles == null) return;
            var existing = new HashSet<string>(
                styles.Elements<Style>().Select(style => style.StyleId?.Value).Where(id => id != null).Cast<string>(),
                StringComparer.OrdinalIgnoreCase);
            bool changed = false;
            foreach (Style style in Create()) {
                if (existing.Add(style.StyleId!.Value!)) {
                    styles.Append(style);
                    changed = true;
                }
            }
            if (changed) styles.Save();
        }

        private static Style CreateHeading() {
            return new Style(
                new StyleName { Val = "TOC Heading" },
                new BasedOn { Val = "Heading1" },
                new NextParagraphStyle { Val = "Normal" },
                new UIPriority { Val = 39 },
                new UnhideWhenUsed(),
                new PrimaryStyle(),
                new StyleParagraphProperties(
                    // Outline level 9 is body text: the TOC title never lists itself.
                    new OutlineLevel { Val = 9 })) {
                Type = StyleValues.Paragraph,
                StyleId = "TOCHeading"
            };
        }

        private static Style CreateEntry(int level) {
            var paragraphProperties = new StyleParagraphProperties(
                new SpacingBetweenLines { After = "100" });
            if (level > 1) {
                paragraphProperties.Append(new Indentation { Left = ((level - 1) * EntryIndentTwips).ToString(System.Globalization.CultureInfo.InvariantCulture) });
            }

            return new Style(
                new StyleName { Val = "toc " + level.ToString(System.Globalization.CultureInfo.InvariantCulture) },
                new BasedOn { Val = "Normal" },
                new NextParagraphStyle { Val = "Normal" },
                new AutoRedefine(),
                new UIPriority { Val = 39 },
                new UnhideWhenUsed(),
                paragraphProperties) {
                Type = StyleValues.Paragraph,
                StyleId = "TOC" + level.ToString(System.Globalization.CultureInfo.InvariantCulture)
            };
        }
    }
}
