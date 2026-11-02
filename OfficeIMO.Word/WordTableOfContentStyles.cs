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

        // Letter text width (8.5in less two 1in margins), just inside the right margin.
        private const int DefaultRightTabTwips = 9350;

        internal static IEnumerable<Style> Create(int rightTabTwips = DefaultRightTabTwips) {
            yield return CreateHeading();
            for (int level = 1; level <= EntryLevels; level++) {
                yield return CreateEntry(level, rightTabTwips);
            }
        }

        /// <summary>Right edge of the first section's text column, so page numbers align with the margin.</summary>
        private static int ResolveRightTabTwips(StyleDefinitionsPart? part) {
            var sectionProperties = (part?.OpenXmlPackage as WordprocessingDocument)?.MainDocumentPart?.Document?.Body?
                .Descendants<SectionProperties>().FirstOrDefault();
            uint? width = sectionProperties?.GetFirstChild<PageSize>()?.Width?.Value;
            PageMargin? margin = sectionProperties?.GetFirstChild<PageMargin>();
            if (width == null || margin == null) return DefaultRightTabTwips;
            long text = (long)width.Value - (margin.Left?.Value ?? 1440U) - (margin.Right?.Value ?? 1440U) - 10;
            return text > 1440 && text < 31680 ? (int)text : DefaultRightTabTwips;
        }

        /// <summary>Adds any missing TOC style to a style part, leaving existing definitions untouched.</summary>
        internal static void EnsureDefined(StyleDefinitionsPart? part) {
            Styles? styles = part?.Styles;
            if (styles == null) return;
            var existing = new HashSet<string>(
                styles.Elements<Style>().Select(style => style.StyleId?.Value).Where(id => id != null).Cast<string>(),
                StringComparer.OrdinalIgnoreCase);
            bool changed = false;
            foreach (Style style in Create(ResolveRightTabTwips(part))) {
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

        private static Style CreateEntry(int level, int rightTabTwips) {
            // Entries end in a dot-leader right tab, so stored entries read as a finished TOC
            // even in viewers that do not regenerate the field.
            var paragraphProperties = new StyleParagraphProperties(
                new Tabs(new TabStop {
                    Val = TabStopValues.Right,
                    Leader = TabStopLeaderCharValues.Dot,
                    Position = rightTabTwips
                }),
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
