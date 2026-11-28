using System.Globalization;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private sealed class NativeTableOfContentsEntry {
            public NativeTableOfContentsEntry(string text, int level, string? destinationName) {
                Text = text;
                Level = level;
                DestinationName = destinationName;
            }

            public string Text { get; }
            public int Level { get; }
            public string? DestinationName { get; }
        }

        private static Dictionary<OfficeXmlNode, string> BuildNativeHeadingDestinations(WordDocument document) {
            var destinations = new Dictionary<OfficeXmlNode, string>();
            var used = new HashSet<string>(StringComparer.Ordinal);
            var nextSuffixByBaseName = new Dictionary<string, int>(StringComparer.Ordinal);
            int headingIndex = 0;

            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in EnumerateNativeTableOfContentsElements(section)) {
                    if (element is not WordParagraph paragraph ||
                        paragraph.Node().IsNone ||
                        GetNativeTableOfContentsHeadingLevel(paragraph) <= 0) {
                        continue;
                    }

                    string headingText = GetNativeParagraphDisplayText(paragraph);
                    if (string.IsNullOrWhiteSpace(headingText)) {
                        continue;
                    }

                    string? bookmarkName = GetNativeParagraphBookmarkNames(paragraph).FirstOrDefault();
                    string destinationName = bookmarkName ?? CreateNativeHeadingDestinationName(headingText, ++headingIndex, used, nextSuffixByBaseName);
                    destinations[paragraph.Node()] = destinationName;
                    used.Add(destinationName);
                }
            }

            return destinations;
        }

        private static string CreateNativeHeadingDestinationName(string text, int headingIndex, HashSet<string> used, Dictionary<string, int> nextSuffixByBaseName) {
            var builder = new StringBuilder("officeimo-heading-");
            foreach (char ch in text) {
                if (char.IsLetterOrDigit(ch)) {
                    builder.Append(char.ToLowerInvariant(ch));
                } else if (builder[builder.Length - 1] != '-') {
                    builder.Append('-');
                }

                if (builder.Length >= 80) {
                    break;
                }
            }

            string baseName = builder.ToString().TrimEnd('-');
            if (baseName.Length <= "officeimo-heading".Length) {
                baseName = "officeimo-heading-" + headingIndex.ToString(CultureInfo.InvariantCulture);
            }

            if (!used.Contains(baseName)) {
                nextSuffixByBaseName[baseName] = 2;
                return baseName;
            }

            int suffix = nextSuffixByBaseName.TryGetValue(baseName, out int nextSuffix) ? nextSuffix : 2;
            string name;
            do {
                name = baseName + "-" + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
            } while (used.Contains(name));

            nextSuffixByBaseName[baseName] = suffix;
            return name;
        }

        private static IReadOnlyList<NativeTableOfContentsEntry> BuildNativeTableOfContentsEntries(WordDocument document, IReadOnlyDictionary<OfficeXmlNode, string> headingDestinations) {
            var entries = new List<NativeTableOfContentsEntry>();
            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in EnumerateNativeTableOfContentsElements(section)) {
                    if (element is not WordParagraph paragraph || paragraph.Node().IsNone ||
                        !headingDestinations.TryGetValue(paragraph.Node(), out string? destination)) continue;
                    string text = GetNativeParagraphDisplayText(paragraph);
                    if (!string.IsNullOrWhiteSpace(text))
                        entries.Add(new NativeTableOfContentsEntry(text, GetNativeTableOfContentsHeadingLevel(paragraph), destination));
                }
            }
            return entries;
        }

        private static IEnumerable<WordElement> EnumerateNativeTableOfContentsElements(WordSection section) {
            foreach (WordElement element in CollapseNativeParagraphElements(section.Elements)) {
                if (element is WordStructuredDocumentTag structuredDocumentTag) {
                    foreach (WordElement structuredElement in CollapseNativeParagraphElements(GetNativeStructuredBlockElements(structuredDocumentTag.Document, structuredDocumentTag.SdtBlockNode()))) {
                        yield return structuredElement;
                    }

                    continue;
                }

                yield return element;
            }
        }

        private static string GetNativeParagraphDisplayText(WordParagraph paragraph) {
            if (paragraph._document.EquationOccurrences(paragraph.Node()).Count > 0) {
                return AppendNativeTextWithEquation(paragraph.Text, paragraph);
            }
            if (paragraph.IsHyperLink && paragraph.Hyperlink != null) {
                return paragraph.Hyperlink.Text;
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            string text = runs.Count > 0
                ? string.Concat(runs.Where(run => !run.IsImage).Select(run => run.Text))
                : paragraph.Text;
            return AppendNativeTextWithEquation(text, paragraph);
        }

        private static int GetNativeTableOfContentsHeadingLevel(WordParagraph paragraph) {
            if (!paragraph.Style.HasValue) {
                return 0;
            }

            return paragraph.Style.Value switch {
                WordParagraphStyles.Heading1 => 1,
                WordParagraphStyles.Heading2 => 2,
                WordParagraphStyles.Heading3 => 3,
                WordParagraphStyles.Heading4 => 4,
                WordParagraphStyles.Heading5 => 5,
                WordParagraphStyles.Heading6 => 6,
                WordParagraphStyles.Heading7 => 7,
                WordParagraphStyles.Heading8 => 8,
                WordParagraphStyles.Heading9 => 9,
                _ => 0
            };
        }

        /// <summary>
        /// The paragraph formatting of the document's TOC 1-9 style, so a table generated from the headings is laid out as
        /// Word lays out the entries it writes, and whether the style sets contextual spacing. The style is null when the
        /// document does not define it.
        /// </summary>
        private static (PdfCore.PdfParagraphStyle? Style, bool ContextualSpacing) CreateNativeTableOfContentsEntryStyle(WordDocument document, int level,
            NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap) {
            WordParagraph paragraph = document.CreateDetachedStyledParagraph("TOC" + level.ToString(CultureInfo.InvariantCulture), "0");
            string? styleId = paragraph.Node().WChild("pPr").ChildVal("pStyle");
            if (styleId == null || !WordStyleCatalog.For(document).Contains(styleId)) return (null, false);
            return (CreateNativeParagraphStyle(paragraph, nativeDefaults, nativeFontMap), UsesNativeContextualSpacing(paragraph));
        }

        private static void RenderNativeTableOfContents(INativePdfFlow pdf, WordTableOfContent tableOfContent, IReadOnlyList<NativeTableOfContentsEntry> entries, double? contentWidth,
            NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap) {
            var levelStyles = new Dictionary<int, (PdfCore.PdfParagraphStyle? Style, bool ContextualSpacing)>();
            PdfCore.PdfTableOfContentsEntry[] navigation = entries
                .Where(entry => entry.Level >= tableOfContent.MinLevel && entry.Level <= tableOfContent.MaxLevel && !string.IsNullOrEmpty(entry.DestinationName))
                .Select(entry => {
                    if (!levelStyles.TryGetValue(entry.Level, out var levelStyle)) {
                        levelStyle = CreateNativeTableOfContentsEntryStyle(tableOfContent.Document, entry.Level, nativeDefaults, nativeFontMap);
                        levelStyles.Add(entry.Level, levelStyle);
                    }
                    return new PdfCore.PdfTableOfContentsEntry(NormalizeNativeDirectText(entry.Text), entry.DestinationName!, entry.Level,
                        "Table of contents: " + NormalizeNativeDirectText(entry.Text)) { ParagraphStyle = levelStyle.Style };
                }).ToArray();
            // Word shares the space between paragraphs: the next entry's space before is reduced by the declared space after.
            // Each level has its own TOC style, so contextual spacing drops the space between consecutive entries of one level.
            for (int index = 0; index + 1 < navigation.Length; index++) {
                PdfCore.PdfTableOfContentsEntry current = navigation[index], next = navigation[index + 1];
                if (current.ParagraphStyle == null || next.ParagraphStyle == null) continue;
                double declaredAfter = current.ParagraphStyle.SpacingAfter?.ToPoints() ?? 0D;
                double before = next.ParagraphStyle.SpacingBefore.ToPoints();
                bool contextual = current.Level == next.Level && levelStyles[current.Level].ContextualSpacing;
                if (contextual) {
                    current.ParagraphStyle = current.ParagraphStyle.Clone();
                    current.ParagraphStyle.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D);
                }
                double sharedBefore = contextual ? 0D : Math.Max(0D, before - declaredAfter);
                if (sharedBefore != before) {
                    next.ParagraphStyle = next.ParagraphStyle.Clone();
                    next.ParagraphStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(sharedBefore);
                }
            }
            if (navigation.Length > 0) {
                pdf.TableOfContents(new PdfCore.PdfTableOfContentsOptions {
                    Title = null,
                    MinimumLevel = tableOfContent.MinLevel,
                    MaximumLevel = tableOfContent.MaxLevel,
                    IndentPerLevel = global::OfficeIMO.OfficeLength.Points(22D),
                    Entries = navigation
                });
            }

            if (navigation.Length == 0) {
                string fallback = string.IsNullOrWhiteSpace(tableOfContent.TextNoContent)
                    ? "No table of contents entries found."
                    : tableOfContent.TextNoContent;
                pdf.Paragraph(builder => builder.AddText(NormalizeNativeDirectText(fallback)).SetFontSize(global::OfficeIMO.OfficeLength.Points(10.5D)));
            }
        }

    }
}
