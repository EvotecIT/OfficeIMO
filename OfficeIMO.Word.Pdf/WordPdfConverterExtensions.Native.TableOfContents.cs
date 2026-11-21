using System.Globalization;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
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

        private static Dictionary<W.Paragraph, string> BuildNativeHeadingDestinations(WordDocument document) {
            var destinations = new Dictionary<W.Paragraph, string>();
            var used = new HashSet<string>(StringComparer.Ordinal);
            var nextSuffixByBaseName = new Dictionary<string, int>(StringComparer.Ordinal);
            int headingIndex = 0;

            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in EnumerateNativeTableOfContentsElements(section)) {
                    if (element is not WordParagraph paragraph ||
                        paragraph._paragraph == null ||
                        GetNativeTableOfContentsHeadingLevel(paragraph) <= 0) {
                        continue;
                    }

                    string headingText = GetNativeParagraphDisplayText(paragraph);
                    if (string.IsNullOrWhiteSpace(headingText)) {
                        continue;
                    }

                    string? bookmarkName = GetNativeParagraphBookmarkNames(paragraph).FirstOrDefault();
                    string destinationName = bookmarkName ?? CreateNativeHeadingDestinationName(headingText, ++headingIndex, used, nextSuffixByBaseName);
                    destinations[paragraph._paragraph] = destinationName;
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

        private static IReadOnlyList<NativeTableOfContentsEntry> BuildNativeTableOfContentsEntries(WordDocument document, IReadOnlyDictionary<W.Paragraph, string> headingDestinations) {
            var entries = new List<NativeTableOfContentsEntry>();
            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in EnumerateNativeTableOfContentsElements(section)) {
                    if (element is not WordParagraph paragraph || paragraph._paragraph == null ||
                        !headingDestinations.TryGetValue(paragraph._paragraph, out string? destination)) continue;
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
                    foreach (WordElement structuredElement in CollapseNativeParagraphElements(GetNativeStructuredBlockElements(structuredDocumentTag.Document, structuredDocumentTag.SdtBlock))) {
                        yield return structuredElement;
                    }

                    continue;
                }

                yield return element;
            }
        }

        private static string GetNativeParagraphDisplayText(WordParagraph paragraph) {
            if (WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph).Count > 0) {
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
        /// Word lays out the entries it writes. Returns null when the document does not define the style.
        /// </summary>
        private static PdfCore.PdfParagraphStyle? CreateNativeTableOfContentsEntryStyle(WordDocument document, int level,
            NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap) {
            var entry = new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = "TOC" + level.ToString(CultureInfo.InvariantCulture) }),
                new W.Run(new W.Text("0")));
            document.LocalizeBuiltInStyleReferences(entry);
            string? styleId = entry.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (styleId == null || !WordStyleCatalog.For(document).Contains(styleId)) return null;
            var paragraph = new WordParagraph(document, entry, entry.GetFirstChild<W.Run>()!);
            return CreateNativeParagraphStyle(paragraph, nativeDefaults, nativeFontMap);
        }

        private static void RenderNativeTableOfContents(INativePdfFlow pdf, WordTableOfContent tableOfContent, IReadOnlyList<NativeTableOfContentsEntry> entries, double? contentWidth, bool renderTitle,
            NativeDocumentDefaults nativeDefaults, NativeFontMap? nativeFontMap) {
            if (renderTitle) {
                string title = string.IsNullOrWhiteSpace(tableOfContent.Text) ? "Table of Contents" : tableOfContent.Text;
                pdf.Paragraph(builder => builder.AddText(NormalizeNativeDirectText(title)).SetFontSize(global::OfficeIMO.OfficeLength.Points(11D)), global::OfficeIMO.OfficeAlignment.Left, null, new PdfCore.PdfParagraphStyle {
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(5D),
                    KeepWithNext = true
                });
            }

            var levelStyles = new Dictionary<int, PdfCore.PdfParagraphStyle?>();
            PdfCore.PdfTableOfContentsEntry[] navigation = entries
                .Where(entry => entry.Level >= tableOfContent.MinLevel && entry.Level <= tableOfContent.MaxLevel && !string.IsNullOrEmpty(entry.DestinationName))
                .Select(entry => {
                    if (!levelStyles.TryGetValue(entry.Level, out PdfCore.PdfParagraphStyle? levelStyle)) {
                        levelStyle = CreateNativeTableOfContentsEntryStyle(tableOfContent.Document, entry.Level, nativeDefaults, nativeFontMap);
                        levelStyles.Add(entry.Level, levelStyle);
                    }
                    return new PdfCore.PdfTableOfContentsEntry(NormalizeNativeDirectText(entry.Text), entry.DestinationName!, entry.Level,
                        "Table of contents: " + NormalizeNativeDirectText(entry.Text)) { ParagraphStyle = levelStyle };
                }).ToArray();
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
