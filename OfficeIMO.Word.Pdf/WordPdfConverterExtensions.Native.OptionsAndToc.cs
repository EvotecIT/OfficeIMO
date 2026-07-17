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
        private static PdfCore.PdfOptions CreateNativeOptions(WordDocument document, PdfSaveOptions? options, NativeFontMap nativeFontMap) {
            WordSection? firstSection = document.Sections.FirstOrDefault();
            PdfCore.PdfOptions pdfOptions = options?.PdfOptions?.Clone() ?? new PdfCore.PdfOptions();
            if (options != null) {
                pdfOptions.ReportDiagnosticsTo(options.Report, "OfficeIMO.Word.Pdf");
            }

            NativeDocumentDefaults defaults = GetNativeDocumentDefaults(document);
            if (options?.PdfOptions == null) {
                pdfOptions.DefaultFontSize = defaults.FontSize;
            }

            pdfOptions.PageSize = firstSection == null ? PdfCore.PageSizes.A4 : GetNativePageSize(firstSection, options);
            pdfOptions.Margins = firstSection == null ? PdfCore.PageMargins.Uniform(72) : GetNativeMargins(firstSection, options);
            bool allowSystemFontEmbedding = options?.ResourcePolicy.AllowSystemFontEmbedding == true;
            bool preserveConfiguredFontSlots = ApplyNativeDefaultFont(document, options, pdfOptions, allowSystemFontEmbedding, nativeFontMap) ||
                                                options?.PdfOptions != null;
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots = RegisterNativeDocumentFonts(document, pdfOptions, preserveConfiguredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            RegisterNativeThemeStyleFonts(document, pdfOptions, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            ApplyNativeTextFallbacks(options, pdfOptions, registeredFontSlots, preserveConfiguredFontSlots, allowSystemFontEmbedding);
            pdfOptions.BackgroundColor = ParseNativeColor(document.Background?.Color);
            pdfOptions.CreateOutlineFromHeadings = true;
            ApplyNativeBiDiViewerPreferences(document, pdfOptions);
            return pdfOptions;
        }

        private static void ApplyNativeBiDiViewerPreferences(WordDocument document, PdfCore.PdfOptions pdfOptions) {
            if (!HasNativeBiDiParagraph(document)) {
                return;
            }

            pdfOptions.ConfigureViewerPreferences(viewerPreferences => {
                if (!viewerPreferences.Direction.HasValue) {
                    viewerPreferences.Direction = PdfCore.PdfViewerDirection.RightToLeft;
                }
            });
        }

        private static bool HasNativeBiDiParagraph(WordDocument document) {
            foreach (WordSection section in document.Sections) {
                if (section.Elements.Any(HasNativeBiDiParagraph) ||
                    HasNativeBiDiParagraph(section.Header?.Default) ||
                    HasNativeBiDiParagraph(section.Header?.First) ||
                    HasNativeBiDiParagraph(section.Header?.Even) ||
                    HasNativeBiDiParagraph(section.Footer?.Default) ||
                    HasNativeBiDiParagraph(section.Footer?.First) ||
                    HasNativeBiDiParagraph(section.Footer?.Even)) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasNativeBiDiParagraph(WordHeaderFooter? headerFooter) =>
            headerFooter?.Elements.Any(HasNativeBiDiParagraph) == true;

        private static bool HasNativeBiDiParagraph(WordElement element) {
            if (element is WordParagraph paragraph) {
                return IsNativeBiDiParagraph(paragraph);
            }

            if (element is WordTable table) {
                foreach (WordTableRow row in table.Rows) {
                    foreach (WordTableCell cell in row.Cells) {
                        foreach (WordElement cellElement in cell.Elements) {
                            if (HasNativeBiDiParagraph(cellElement)) {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        private static bool ApplyNativeDefaultFont(WordDocument document, PdfSaveOptions? options, PdfCore.PdfOptions pdfOptions, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            string? optionFontFamily = options?.FontFamily;
            if (!string.IsNullOrWhiteSpace(optionFontFamily) &&
                TryApplyNativeDefaultFontCandidate(optionFontFamily, pdfOptions, embedSystemFont: allowSystemFontEmbedding)) {
                nativeFontMap.Register(optionFontFamily!, pdfOptions.DefaultFont);
                nativeFontMap.PreferPdfDefaultForDocumentDefaultFont();
                return true;
            }

            foreach (string? family in new[] {
                document.Settings.FontFamily,
                document.Settings.FontFamilyHighAnsi,
                document.Settings.FontFamilyEastAsia,
                document.Settings.FontFamilyComplexScript
            }) {
                if (TryApplyNativeDefaultFontCandidate(family, pdfOptions, embedSystemFont: allowSystemFontEmbedding)) {
                    nativeFontMap.Register(family!, pdfOptions.DefaultFont);
                    return true;
                }
            }

            return false;
        }

        private static void ApplyNativeTextFallbacks(
            PdfSaveOptions? options,
            PdfCore.PdfOptions pdfOptions,
            HashSet<PdfCore.PdfStandardFont> reservedFontSlots,
            bool preserveConfiguredFontSlots,
            bool allowSystemFontEmbedding) {
            if (options == null ||
                !allowSystemFontEmbedding ||
                options.TextFallbacks == PdfCore.PdfTextFallbackFeatures.None) {
                return;
            }

            PdfCore.PdfTextFallbackFeatures fallbackFeatures = options.TextFallbacks;
            if (preserveConfiguredFontSlots || pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont)) {
                fallbackFeatures &= ~PdfCore.PdfTextFallbackFeatures.DocumentFont;
            }

            if (fallbackFeatures != PdfCore.PdfTextFallbackFeatures.None) {
                pdfOptions.UseTextFallbacks(fallbackFeatures, reservedFontSlots, allowSystemFontEmbedding);
                foreach (PdfCore.PdfStandardFont slot in pdfOptions.EmbeddedFontFallbacks?.FontSlots ?? Array.Empty<PdfCore.PdfStandardFont>()) {
                    PdfCore.PdfOptions.AddRegisteredFontFamilySlot(reservedFontSlots, slot);
                }
            }
        }

        private static bool TryApplyNativeDefaultFontCandidate(string? familyName, PdfCore.PdfOptions pdfOptions, bool embedSystemFont, bool requireEmbeddedFont = false) {
            return pdfOptions.TryUseOfficeFontFamily(familyName, embedSystemFont, requireEmbeddedFont);
        }

        private static HashSet<PdfCore.PdfStandardFont> RegisterNativeDocumentFonts(WordDocument document, PdfCore.PdfOptions pdfOptions, bool preserveConfiguredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots = pdfOptions.CreateRegisteredFontFamilySlots(preserveConfiguredFontSlots);
            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in CollapseNativeParagraphElements(section.Elements)) {
                    if (element is WordCoverPage coverPage) {
                        foreach (WordElement coverElement in GetNativeStructuredBlockElements(coverPage.Document, coverPage.SdtBlock)) {
                            RegisterNativeElementFonts(coverElement, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                        }

                        continue;
                    }

                    if (element is WordStructuredDocumentTag structuredDocumentTag) {
                        foreach (WordElement structuredElement in GetNativeStructuredBlockElements(structuredDocumentTag.Document, structuredDocumentTag.SdtBlock)) {
                            RegisterNativeElementFonts(structuredElement, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                        }

                        continue;
                    }

                    RegisterNativeElementFonts(element, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                }

                RegisterNativeHeaderFooterFonts(section.Header?.Default, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                RegisterNativeHeaderFooterFonts(section.Header?.First, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                RegisterNativeHeaderFooterFonts(section.Header?.Even, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                RegisterNativeHeaderFooterFonts(section.Footer?.Default, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                RegisterNativeHeaderFooterFonts(section.Footer?.First, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                RegisterNativeHeaderFooterFonts(section.Footer?.Even, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);

                foreach (WordWatermark watermark in section.Watermarks) {
                    RegisterNativeFontCandidate(watermark.FontFamily, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                }
            }

            return registeredFontSlots;
        }

        private static void RegisterNativeHeaderFooterFonts(WordHeaderFooter? headerFooter, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            if (headerFooter == null) {
                return;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                RegisterNativeElementFonts(element, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            }
        }

        private static void RegisterNativeElementFonts(WordElement element, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            if (element is WordParagraph paragraph) {
                RegisterNativeParagraphFonts(paragraph, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                    RegisterNativeParagraphFonts(run, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                }
            } else if (element is WordTable table) {
                RegisterNativeTableFonts(table, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            }
        }

        private static void RegisterNativeTableFonts(WordTable table, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            NativeTableStyleDefaults tableStyleDefaults = GetNativeTableStyleDefaults(
                table,
                GetNativeDocumentDefaults(table.Document),
                ignoreFallbackTableStyle: pdfOptions.HasExplicitDefaultTableStyle);
            RegisterNativeFontCandidate(tableStyleDefaults.RunStyle.FontFamily, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);

            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    foreach (WordParagraph paragraph in cell.Paragraphs) {
                        RegisterNativeParagraphFonts(paragraph, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                        foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                            RegisterNativeParagraphFonts(run, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                        }
                    }

                    foreach (WordTable nestedTable in cell.NestedTables) {
                        RegisterNativeTableFonts(nestedTable, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                    }
                }
            }
        }

        private static void RegisterNativeParagraphFonts(WordParagraph paragraph, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            RegisterNativeFontCandidate(paragraph.FontFamily, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            RegisterNativeFontCandidate(paragraph.FontFamilyHighAnsi, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            RegisterNativeFontCandidate(paragraph.FontFamilyEastAsia, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            RegisterNativeFontCandidate(paragraph.FontFamilyComplexScript, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            RegisterNativeFontCandidate(GetNativeParagraphStyleDefaults(paragraph).FontFamily, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            RegisterNativeFontCandidate(GetNativeCharacterStyleDefaults(paragraph._document, GetNativeRunProperties(paragraph)).FontFamily, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
        }

        private static void RegisterNativeFontCandidate(string? familyName, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            if (!PdfCore.PdfOptions.TryAddOfficeFontFamilyKey(familyName, registeredFamilies, NormalizeNativeFontFamily, out string trimmedFamilyName)) {
                return;
            }

            if (allowSystemFontEmbedding &&
                PdfCore.PdfStandardFontMapper.TryMapFontFamily(trimmedFamilyName, out PdfCore.PdfStandardFont mappedFont) &&
                registeredFontSlots.Contains(PdfCore.PdfStandardFontMapper.GetFontFamily(mappedFont)) &&
                !EmbeddedFontSlotMatchesFamily(pdfOptions, mappedFont, trimmedFamilyName) &&
                PdfCore.PdfOptions.TrySelectAvailableFontFamilySlot(trimmedFamilyName, registeredFontSlots, out PdfCore.PdfStandardFont distinctFontSlot) &&
                PdfCore.PdfEmbeddedFontFamily.TryFromSystem(trimmedFamilyName, out PdfCore.PdfEmbeddedFontFamily? distinctEmbeddedFamily) &&
                distinctEmbeddedFamily != null) {
                registeredFontSlots.Add(distinctFontSlot);
                pdfOptions.RegisterFontFamily(distinctFontSlot, distinctEmbeddedFamily);
                nativeFontMap.Register(trimmedFamilyName, distinctFontSlot);
                return;
            }

            if (pdfOptions.TryRegisterMappedOfficeFontFamily(trimmedFamilyName, registeredFontSlots, allowSystemFontEmbedding, out PdfCore.PdfStandardFont fontFamily)) {
                nativeFontMap.Register(trimmedFamilyName, fontFamily);
                return;
            }

            if (allowSystemFontEmbedding &&
                PdfCore.PdfOptions.TrySelectAvailableFontFamilySlot(trimmedFamilyName, registeredFontSlots, out PdfCore.PdfStandardFont fontSlot) &&
                PdfCore.PdfEmbeddedFontFamily.TryFromSystem(trimmedFamilyName, out PdfCore.PdfEmbeddedFontFamily? embeddedFamily) &&
                embeddedFamily != null) {
                registeredFontSlots.Add(fontSlot);
                pdfOptions.RegisterFontFamily(fontSlot, embeddedFamily);
                nativeFontMap.Register(trimmedFamilyName, fontSlot);
            }
        }

        private static bool EmbeddedFontSlotMatchesFamily(PdfCore.PdfOptions options, PdfCore.PdfStandardFont slot, string familyName) {
            PdfCore.PdfStandardFont normalizedSlot = PdfCore.PdfStandardFontMapper.GetFontFamily(slot);
            return !options.EmbeddedFonts.TryGetValue(normalizedSlot, out PdfCore.PdfEmbeddedFont? embedded) ||
                string.Equals(
                    NormalizeNativeFontFamily(embedded.FontName ?? string.Empty),
                    NormalizeNativeFontFamily(familyName),
                    StringComparison.OrdinalIgnoreCase);
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

                    string? bookmarkName = string.IsNullOrWhiteSpace(paragraph.Bookmark?.Name)
                        ? null
                        : paragraph.Bookmark!.Name;
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

                List<WordElement> elements = EnumerateNativeTableOfContentsElements(section).ToList();
                for (int index = 0; index < elements.Count; index++) {
                    WordElement element = elements[index];
                    if (element is WordCoverPage) {
                        if (index + 1 >= elements.Count || !IsNativeTableOfContentsExplicitPageBreak(elements[index + 1])) {
                            currentPage++;
                            consumedOnPage = 0D;
                        }

                        continue;
                    }

                    if (element is WordParagraph paragraph && HasNativePageBreakBefore(paragraph)) {
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
                foreach (WordElement element in EnumerateNativeTableOfContentsElements(section)) {
                    if (element is WordParagraph paragraph &&
                        GetNativeTableOfContentsHeadingLevel(paragraph) > 0 &&
                        !string.IsNullOrWhiteSpace(GetNativeParagraphDisplayText(paragraph))) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static bool IsNativeTableOfContentsExplicitPageBreak(WordElement element) {
            if (element is WordParagraph paragraph) {
                return HasNativePageBreakBefore(paragraph) || paragraph.IsPageBreak;
            }

            return element is WordBreak wordBreak && wordBreak.BreakType == W.BreakValues.Page;
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
                paragraph.Chart == null &&
                paragraph.PictureControl?.Image == null) {
                return 0D;
            }

            if (paragraph.Chart != null) {
                (double _, double chartHeight) = GetNativeWordChartSizePoints(paragraph.Chart);
                return chartHeight + 8D;
            }

            int headingLevel = GetNativeTableOfContentsHeadingLevel(paragraph);
            if (headingLevel > 0) {
                double headingSize = headingLevel == 1 ? 18D : headingLevel == 2 ? 15D : 13D;
                return EstimateNativeLineCount(text, contentWidth, headingSize) * headingSize * 1.25D + 8D;
            }

            double fontSize = paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0 ? paragraph.FontSize.Value : 11D;
            double height = EstimateNativeLineCount(text, contentWidth, fontSize) * fontSize * NativeDefaultParagraphLineHeight + NativeDefaultParagraphSpacingAfter;
            NativeParagraphBorders borders = GetNativeEffectiveParagraphBorders(paragraph);
            if (!string.IsNullOrWhiteSpace(GetNativeEffectiveParagraphShadingFill(paragraph)) ||
                HasNativeParagraphBorder(borders)) {
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

        private static void RenderNativeTableOfContents(INativePdfFlow pdf, WordTableOfContent tableOfContent, IReadOnlyList<NativeTableOfContentsEntry> entries, double? contentWidth) {
            string title = string.IsNullOrWhiteSpace(tableOfContent.Text) ? "Table of Contents" : tableOfContent.Text;
            pdf.Paragraph(builder => builder.FontSize(11D).Text(NormalizeNativeDirectText(title)), PdfCore.PdfAlign.Left, null, new PdfCore.PdfParagraphStyle {
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
                PdfCore.PdfParagraphStyle style = CreateNativeTableOfContentsEntryStyle(relativeLevel, contentWidth);
                pdf.Paragraph(
                    builder => {
                        builder.FontSize(10.5D);
                        if (string.IsNullOrEmpty(entry.DestinationName)) {
                            builder.Text(NormalizeNativeDirectText(entry.Text));
                        } else {
                            string entryText = NormalizeNativeDirectText(entry.Text);
                            builder.LinkToBookmark(entryText, entry.DestinationName!, underline: false, contents: "Table of contents: " + entryText);
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
                pdf.Paragraph(builder => builder.FontSize(10.5D).Text(NormalizeNativeDirectText(fallback)));
            }
        }

        private static PdfCore.PdfParagraphStyle CreateNativeTableOfContentsEntryStyle(int relativeLevel, double? contentWidth) {
            double leftIndent = GetNativeTableOfContentsLevelIndent(relativeLevel);
            double effectiveContentWidth = contentWidth.HasValue && contentWidth.Value > 0D ? contentWidth.Value : 432D;
            double textFrameWidth = Math.Max(36D, effectiveContentWidth - leftIndent);

            return new PdfCore.PdfParagraphStyle {
                LeftIndent = leftIndent,
                SpacingAfter = 1D,
                DefaultTabStopWidth = textFrameWidth,
                KeepWithNext = true
            };
        }

        private static double GetNativeTableOfContentsLevelIndent(int relativeLevel) {
            if (relativeLevel <= 0) {
                return 0D;
            }

            return Math.Min(relativeLevel, 8) * 22D;
        }

    }
}
