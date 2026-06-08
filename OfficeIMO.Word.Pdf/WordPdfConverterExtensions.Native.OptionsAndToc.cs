using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static PdfCore.PdfOptions CreateNativeOptions(WordDocument document, PdfSaveOptions? options) {
            WordSection? firstSection = document.Sections.FirstOrDefault();
            PdfCore.PdfOptions pdfOptions = options?.PdfOptions?.Clone() ?? new PdfCore.PdfOptions();
            pdfOptions.PageSize = firstSection == null ? PdfCore.PageSizes.A4 : GetNativePageSize(firstSection, options);
            pdfOptions.Margins = firstSection == null ? PdfCore.PageMargins.Uniform(72) : GetNativeMargins(firstSection, options);
            bool preserveConfiguredFontSlots = ApplyNativeDefaultFont(document, options, pdfOptions) ||
                                                options?.PdfOptions != null;
            RegisterNativeDocumentFonts(document, pdfOptions, preserveConfiguredFontSlots);
            ApplyNativeFallbackFont(options, pdfOptions, preserveConfiguredFontSlots);
            pdfOptions.BackgroundColor = ParseNativeColor(document.Background?.Color);
            pdfOptions.CreateOutlineFromHeadings = true;
            return pdfOptions;
        }

        private static bool ApplyNativeDefaultFont(WordDocument document, PdfSaveOptions? options, PdfCore.PdfOptions pdfOptions) {
            string? optionFontFamily = options?.FontFamily;
            if (!string.IsNullOrWhiteSpace(optionFontFamily) &&
                TryApplyNativeDefaultFontCandidate(optionFontFamily, pdfOptions, embedSystemFont: true)) {
                return true;
            }

            foreach (string? family in new[] {
                document.Settings.FontFamily,
                document.Settings.FontFamilyHighAnsi,
                document.Settings.FontFamilyEastAsia,
                document.Settings.FontFamilyComplexScript
            }) {
                if (TryApplyNativeDefaultFontCandidate(family, pdfOptions, embedSystemFont: true)) {
                    return true;
                }
            }

            return false;
        }

        private static void ApplyNativeFallbackFont(PdfSaveOptions? options, PdfCore.PdfOptions pdfOptions, bool preserveConfiguredFontSlots) {
            if (options?.PdfOptions == null &&
                !preserveConfiguredFontSlots &&
                !pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont)) {
                pdfOptions.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true);
            }
        }

        private static bool TryApplyNativeDefaultFontCandidate(string? familyName, PdfCore.PdfOptions pdfOptions, bool embedSystemFont, bool requireEmbeddedFont = false) {
            return pdfOptions.TryUseOfficeFontFamily(familyName, embedSystemFont, requireEmbeddedFont);
        }

        private static void RegisterNativeDocumentFonts(WordDocument document, PdfCore.PdfOptions pdfOptions, bool preserveConfiguredFontSlots) {
            var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots = CreateNativeRegisteredFontSlots(pdfOptions, preserveConfiguredFontSlots);
            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in CollapseNativeParagraphElements(section.Elements)) {
                    if (element is WordParagraph paragraph) {
                        RegisterNativeParagraphFonts(paragraph, pdfOptions, registeredFamilies, registeredFontSlots);
                        foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                            RegisterNativeParagraphFonts(run, pdfOptions, registeredFamilies, registeredFontSlots);
                        }
                    } else if (element is WordTable table) {
                        RegisterNativeTableFonts(table, pdfOptions, registeredFamilies, registeredFontSlots);
                    }
                }
            }
        }

        private static HashSet<PdfCore.PdfStandardFont> CreateNativeRegisteredFontSlots(PdfCore.PdfOptions pdfOptions, bool preserveConfiguredFontSlots) {
            var registeredFontSlots = new HashSet<PdfCore.PdfStandardFont>();
            if (preserveConfiguredFontSlots) {
                AddNativeRegisteredFontSlot(registeredFontSlots, pdfOptions.DefaultFont);
                AddNativeRegisteredFontSlot(registeredFontSlots, pdfOptions.HeaderFont);
                AddNativeRegisteredFontSlot(registeredFontSlots, pdfOptions.FooterFont);
            }

            foreach (PdfCore.PdfStandardFont embeddedFont in pdfOptions.EmbeddedFonts.Keys) {
                AddNativeRegisteredFontSlot(registeredFontSlots, embeddedFont);
            }

            return registeredFontSlots;
        }

        private static void AddNativeRegisteredFontSlot(HashSet<PdfCore.PdfStandardFont> registeredFontSlots, PdfCore.PdfStandardFont font) {
            registeredFontSlots.Add(PdfCore.PdfStandardFontMapper.GetFontFamily(font));
        }

        private static void RegisterNativeTableFonts(WordTable table, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots) {
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    foreach (WordParagraph paragraph in cell.Paragraphs) {
                        RegisterNativeParagraphFonts(paragraph, pdfOptions, registeredFamilies, registeredFontSlots);
                        foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                            RegisterNativeParagraphFonts(run, pdfOptions, registeredFamilies, registeredFontSlots);
                        }
                    }

                    foreach (WordTable nestedTable in cell.NestedTables) {
                        RegisterNativeTableFonts(nestedTable, pdfOptions, registeredFamilies, registeredFontSlots);
                    }
                }
            }
        }

        private static void RegisterNativeParagraphFonts(WordParagraph paragraph, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots) {
            RegisterNativeFontCandidate(paragraph.FontFamily, pdfOptions, registeredFamilies, registeredFontSlots);
            RegisterNativeFontCandidate(paragraph.FontFamilyHighAnsi, pdfOptions, registeredFamilies, registeredFontSlots);
            RegisterNativeFontCandidate(paragraph.FontFamilyEastAsia, pdfOptions, registeredFamilies, registeredFontSlots);
            RegisterNativeFontCandidate(paragraph.FontFamilyComplexScript, pdfOptions, registeredFamilies, registeredFontSlots);
        }

        private static void RegisterNativeFontCandidate(string? familyName, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots) {
            if (string.IsNullOrWhiteSpace(familyName)) {
                return;
            }

            string trimmedFamilyName = familyName!.Trim();
            if (!registeredFamilies.Add(trimmedFamilyName)) {
                return;
            }

            if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(trimmedFamilyName, out PdfCore.PdfStandardFont standardFont)) {
                PdfCore.PdfStandardFont fontFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(standardFont);
                if (registeredFontSlots.Add(fontFamily)) {
                    pdfOptions.RegisterOfficeFontFamily(trimmedFamilyName, fontFamily);
                }
            }
        }

        private sealed class NativeTableOfContentsEntry {
            public NativeTableOfContentsEntry(string text, int level, int pageNumber, string? destinationName) {
                Text = text;
                Level = level;
                PageNumber = pageNumber;
                DestinationName = destinationName;
            }

            public string Text { get; }
            public int Level { get; }
            public int PageNumber { get; }
            public string? DestinationName { get; }
        }

        private static Dictionary<W.Paragraph, string> BuildNativeHeadingDestinations(WordDocument document) {
            var destinations = new Dictionary<W.Paragraph, string>();
            var used = new HashSet<string>(StringComparer.Ordinal);
            int headingIndex = 0;

            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in CollapseNativeParagraphElements(section.Elements)) {
                    if (element is not WordParagraph paragraph ||
                        paragraph._paragraph == null ||
                        GetNativeTableOfContentsHeadingLevel(paragraph) <= 0) {
                        continue;
                    }

                    string headingText = GetNativeParagraphDisplayText(paragraph);
                    if (string.IsNullOrWhiteSpace(headingText)) {
                        continue;
                    }

                    string? bookmarkName = string.IsNullOrWhiteSpace(paragraph.Bookmark?.Name)
                        ? null
                        : paragraph.Bookmark!.Name;
                    string destinationName = bookmarkName ?? CreateNativeHeadingDestinationName(headingText, ++headingIndex, used);
                    destinations[paragraph._paragraph] = destinationName;
                    used.Add(destinationName);
                }
            }

            return destinations;
        }

        private static string CreateNativeHeadingDestinationName(string text, int headingIndex, HashSet<string> used) {
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

            string name = baseName;
            int suffix = 2;
            while (used.Contains(name)) {
                name = baseName + "-" + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
            }

            return name;
        }

        private static IReadOnlyList<NativeTableOfContentsEntry> BuildNativeTableOfContentsEntries(WordDocument document, PdfSaveOptions? options, IReadOnlyDictionary<W.Paragraph, string> headingDestinations) {
            var entries = new List<NativeTableOfContentsEntry>();
            int headingCount = CountNativeDocumentHeadings(document);
            int currentPage = 1;
            double consumedOnPage = 0D;
            bool firstSection = true;

            foreach (WordSection section in document.Sections) {
                if (!firstSection) {
                    currentPage++;
                    consumedOnPage = 0D;
                }

                firstSection = false;
                PdfCore.PageSize pageSize = GetNativePageSize(section, options);
                PdfCore.PageMargins margins = GetNativeMargins(section, options);
                double contentHeight = Math.Max(72D, pageSize.Height - margins.Top - margins.Bottom);
                double contentWidth = Math.Max(72D, pageSize.Width - margins.Left - margins.Right);

                foreach (WordElement element in CollapseNativeParagraphElements(section.Elements)) {
                    if (element is WordParagraph paragraph && paragraph.PageBreakBefore) {
                        currentPage++;
                        consumedOnPage = 0D;
                    }

                    if (element is WordParagraph pageBreakParagraph && pageBreakParagraph.IsPageBreak) {
                        currentPage++;
                        consumedOnPage = 0D;
                        continue;
                    }

                    if (element is WordBreak wordBreak && wordBreak.BreakType == W.BreakValues.Page) {
                        currentPage++;
                        consumedOnPage = 0D;
                        continue;
                    }

                    double estimatedHeight = EstimateNativeElementHeight(element, contentWidth, headingCount);
                    if (estimatedHeight <= 0D) {
                        continue;
                    }

                    if (consumedOnPage > 0D && consumedOnPage + estimatedHeight > contentHeight) {
                        currentPage++;
                        consumedOnPage = 0D;
                    }

                    if (element is WordParagraph headingParagraph) {
                        int headingLevel = GetNativeTableOfContentsHeadingLevel(headingParagraph);
                        if (headingLevel > 0) {
                            string headingText = GetNativeParagraphDisplayText(headingParagraph);
                            if (!string.IsNullOrWhiteSpace(headingText)) {
                                string? destinationName = headingParagraph._paragraph != null &&
                                    headingDestinations.TryGetValue(headingParagraph._paragraph, out string? foundDestination)
                                        ? foundDestination
                                        : null;
                                entries.Add(new NativeTableOfContentsEntry(headingText, headingLevel, currentPage, destinationName));
                            }
                        }
                    }

                    consumedOnPage += estimatedHeight;
                    while (consumedOnPage > contentHeight) {
                        currentPage++;
                        consumedOnPage -= contentHeight;
                    }
                }
            }

            return entries;
        }

        private static int CountNativeDocumentHeadings(WordDocument document) {
            int count = 0;
            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in CollapseNativeParagraphElements(section.Elements)) {
                    if (element is WordParagraph paragraph &&
                        GetNativeTableOfContentsHeadingLevel(paragraph) > 0 &&
                        !string.IsNullOrWhiteSpace(GetNativeParagraphDisplayText(paragraph))) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static double EstimateNativeElementHeight(WordElement element, double contentWidth, int headingCount) {
            switch (element) {
                case WordTableOfContent:
                    return 18D + Math.Max(1, headingCount) * 15D + 10D;
                case WordTable table:
                    return EstimateNativeTableHeight(table, contentWidth);
                case WordImage image:
                    return image.Height.HasValue ? image.Height.Value * 72D / 96D + 6D : 150D;
                case WordParagraph paragraph:
                    return EstimateNativeParagraphHeight(paragraph, contentWidth);
                default:
                    return 0D;
            }
        }

        private static double EstimateNativeTableHeight(WordTable table, double contentWidth) {
            int rowCount = Math.Max(1, table.Rows.Count);
            int columnCount = Math.Max(1, table.Rows.Select(row => row.Cells.Count).DefaultIfEmpty(1).Max());
            double cellWidth = Math.Max(48D, contentWidth / columnCount);
            double height = 0D;
            foreach (WordTableRow row in table.Rows) {
                int rowLines = 1;
                foreach (WordTableCell cell in row.Cells) {
                    string cellText = GetNativeCellText(cell);
                    rowLines = Math.Max(rowLines, EstimateNativeLineCount(cellText, cellWidth, 10D));
                }

                height += rowLines * 14D + 12D;
            }

            return Math.Max(rowCount * 22D, height) + 6D;
        }

        private static double EstimateNativeParagraphHeight(WordParagraph paragraph, double contentWidth) {
            if (paragraph.IsPageBreak) {
                return 0D;
            }

            string text = GetNativeParagraphDisplayText(paragraph);
            if (string.IsNullOrWhiteSpace(text) &&
                paragraph.Image == null &&
                paragraph.Shape == null &&
                paragraph.PictureControl?.Image == null) {
                return 0D;
            }

            int headingLevel = GetNativeTableOfContentsHeadingLevel(paragraph);
            if (headingLevel > 0) {
                double headingSize = headingLevel == 1 ? 18D : headingLevel == 2 ? 15D : 13D;
                return EstimateNativeLineCount(text, contentWidth, headingSize) * headingSize * 1.25D + 8D;
            }

            double fontSize = paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0 ? paragraph.FontSize.Value : 11D;
            double height = EstimateNativeLineCount(text, contentWidth, fontSize) * fontSize * NativeDefaultParagraphLineHeight + NativeDefaultParagraphSpacingAfter;
            if (!string.IsNullOrWhiteSpace(paragraph.ShadingFillColorHex) ||
                HasNativeBorder(paragraph.Borders.TopStyle) ||
                HasNativeBorder(paragraph.Borders.BottomStyle) ||
                HasNativeBorder(paragraph.Borders.LeftStyle) ||
                HasNativeBorder(paragraph.Borders.RightStyle)) {
                height += 8D;
            }

            return height;
        }

        private static int EstimateNativeLineCount(string? text, double contentWidth, double fontSize) {
            if (string.IsNullOrEmpty(text)) {
                return 1;
            }

            double averageCharacterWidth = Math.Max(3D, fontSize * 0.48D);
            int charactersPerLine = Math.Max(12, (int)Math.Floor(contentWidth / averageCharacterWidth));
            int lines = 0;
            foreach (string part in text!.Replace("\r\n", "\n").Split('\n')) {
                lines += Math.Max(1, (int)Math.Ceiling(part.Length / (double)charactersPerLine));
            }

            return Math.Max(1, lines);
        }

        private static string GetNativeParagraphDisplayText(WordParagraph paragraph) {
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

        private static void RenderNativeTableOfContents(INativePdfFlow pdf, WordTableOfContent tableOfContent, IReadOnlyList<NativeTableOfContentsEntry> entries) {
            string title = string.IsNullOrWhiteSpace(tableOfContent.Text) ? "Table of Contents" : tableOfContent.Text;
            pdf.Paragraph(builder => builder.FontSize(11D).Text(title), PdfCore.PdfAlign.Left, null, new PdfCore.PdfParagraphStyle {
                SpacingAfter = 5D,
                KeepWithNext = true
            });

            int minLevel = tableOfContent.MinLevel;
            int maxLevel = tableOfContent.MaxLevel;
            int rendered = 0;
            foreach (NativeTableOfContentsEntry entry in entries) {
                if (entry.Level < minLevel || entry.Level > maxLevel) {
                    continue;
                }

                int relativeLevel = Math.Max(0, entry.Level - minLevel);
                var style = new PdfCore.PdfParagraphStyle {
                    LeftIndent = relativeLevel * 14D,
                    SpacingAfter = 1D,
                    DefaultTabStopWidth = 432D,
                    KeepWithNext = true
                };
                pdf.Paragraph(
                    builder => {
                        builder.FontSize(10.5D);
                        if (string.IsNullOrEmpty(entry.DestinationName)) {
                            builder.Text(entry.Text);
                        } else {
                            builder.LinkToBookmark(entry.Text, entry.DestinationName!, underline: false, contents: "Table of contents: " + entry.Text);
                        }

                        builder
                            .Tab(PdfCore.PdfTabLeaderStyle.Dots, PdfCore.PdfTabAlignment.Right)
                            .Text(entry.PageNumber.ToString(CultureInfo.InvariantCulture));
                    },
                    PdfCore.PdfAlign.Left,
                    null,
                    style);
                rendered++;
            }

            if (rendered == 0) {
                string fallback = string.IsNullOrWhiteSpace(tableOfContent.TextNoContent)
                    ? "No table of contents entries found."
                    : tableOfContent.TextNoContent;
                pdf.Paragraph(builder => builder.FontSize(10.5D).Text(fallback));
            }
        }

    }
}
