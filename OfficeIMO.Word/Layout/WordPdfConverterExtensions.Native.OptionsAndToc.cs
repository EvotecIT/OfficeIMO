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
        private static PdfCore.PdfOptions CreateNativeOptions(
            WordDocument document,
            WordToPdfOptions? options,
            NativeFontMap nativeFontMap,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            WordSection? firstSection = document.Sections.FirstOrDefault();
            PdfCore.PdfOptions pdfOptions = options?.PdfOptions?.Clone() ?? new PdfCore.PdfOptions();
            pdfOptions.UseFontUnitTextLayout = true;
            // Flat running content follows the same source defaults as styled runs.
            // Retain the configured shaper so explicit ligatures use its Latin feature path.
            pdfOptions.DefaultTextFeatureSettings = default(NativeTextSpacing).ResolveFeatureSettings(pdfOptions.DefaultFontSize);
            if (!pdfOptions.HasExplicitTextWhitespaceConfiguration) {
                pdfOptions.TextWhitespaceMode = PdfCore.PdfTextWhitespaceMode.Preserve;
            }
            pdfOptions.UseContentStreamCompressionByDefault();
            if (options != null) {
                pdfOptions.ReportDiagnosticsTo(options.Report, "OfficeIMO.Word.Pdf");
            }

            NativeDocumentDefaults defaults = GetNativeDocumentDefaults(document);
            bool hasConfiguredPdfOptions =
                options?.HasExplicitPdfFontConfiguration == true;
            bool hasConfiguredDefaultFontSize =
                options?.HasExplicitPdfDefaultFontSizeConfiguration == true;
            if (!hasConfiguredDefaultFontSize) {
                pdfOptions.DefaultFontSize = defaults.FontSize;
            }
            if ((string.IsNullOrWhiteSpace(pdfOptions.Language) ||
                 string.Equals(pdfOptions.Language, "und", StringComparison.OrdinalIgnoreCase)) &&
                !string.IsNullOrWhiteSpace(defaults.Language)) {
                pdfOptions.Language = defaults.Language;
            }

            pdfOptions.PageSize = firstSection == null ? PdfCore.PageSizes.A4 : GetNativePageSize(firstSection, options);
            pdfOptions.Margins = firstSection == null
                ? PdfCore.PageMargins.Uniform(72)
                : GetNativeMargins(firstSection, options, GetNativeHeaderFooterMarginExpansion(firstSection, options, listMarkers));
            pdfOptions.MirrorMargins = firstSection != null && ShouldMirrorNativeMargins(firstSection, options);
            bool allowSystemFontEmbedding = options?.ResourcePolicy.AllowSystemFontEmbedding == true;
            bool allowDocumentFontEmbedding = allowSystemFontEmbedding &&
                                              options?.ResourcePolicy.AllowDocumentFontEmbedding == true;
            bool appliedNativeDefaultFont = false;
            if (!hasConfiguredPdfOptions || !string.IsNullOrWhiteSpace(options?.FontFamily)) {
                appliedNativeDefaultFont = ApplyNativeDefaultFont(
                    document,
                    defaults,
                    options,
                    pdfOptions,
                    allowSystemFontEmbedding,
                    allowDocumentFontEmbedding,
                    nativeFontMap);
            }
            bool hasResolvedDocumentDefaultSubstitution =
                string.IsNullOrWhiteSpace(options?.FontFamily) &&
                EnumerateNativeFontFamilies(defaults.FontFamily, defaults.FontFamilies).Any(family =>
                    pdfOptions.TryResolveFontFamilySubstitution(OfficeFontFamilyParser.Quote(family), out PdfCore.PdfFontFamilySubstitution? substitution) &&
                    substitution != null);
            // Fallback candidates select no body font, so with no caller body font a registered
            // document default family stays on the run-resolution path, as it does without them.
            bool hasRegisteredDocumentDefaultFamily =
                string.IsNullOrWhiteSpace(options?.FontFamily) &&
                !pdfOptions.HasExplicitDefaultFont &&
                !pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont) &&
                EnumerateNativeFontFamilies(defaults.FontFamily, defaults.FontFamilies).Any(family =>
                    nativeFontMap.TryGetNamedFontFamily(family, out _) || nativeFontMap.TryGetFontSlot(family, out _) ||
                    pdfOptions.TryResolveNamedOfficeFontFamily(OfficeFontFamilyParser.Quote(family), out _) ||
                    pdfOptions.TryResolveEmbeddedFontFamilySlot(family, out _) ||
                    TryRegisterNativeDeclaredFontCandidate(family, pdfOptions, allowDocumentFontEmbedding, nativeFontMap));
            if (hasConfiguredPdfOptions &&
                !appliedNativeDefaultFont &&
                !hasResolvedDocumentDefaultSubstitution &&
                !hasRegisteredDocumentDefaultFamily) {
                nativeFontMap.PreferPdfDefaultForDocumentDefaultFont();
            }

            bool preserveConfiguredFontSlots = appliedNativeDefaultFont || hasConfiguredPdfOptions;
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots = RegisterNativeDocumentFonts(document, pdfOptions, preserveConfiguredFontSlots, allowDocumentFontEmbedding, nativeFontMap);
            ApplyNativeTextFallbacks(document, options, pdfOptions, registeredFontSlots, allowSystemFontEmbedding, listMarkers);
            pdfOptions.BackgroundColor = ParseNativeColor(document.Background?.Color);
            pdfOptions.CreateOutlineFromHeadings = true;
            ApplyNativeBiDiViewerPreferences(document, pdfOptions);
            nativeFontMap.AttachResolvedOptions(pdfOptions);
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
                    HasNativeBiDiParagraph(section.ResolveDefaultHeader()) ||
                    HasNativeBiDiParagraph(section.ResolveFirstHeader()) ||
                    HasNativeBiDiParagraph(section.ResolveEvenHeader()) ||
                    HasNativeBiDiParagraph(section.ResolveDefaultFooter()) ||
                    HasNativeBiDiParagraph(section.ResolveFirstFooter()) ||
                    HasNativeBiDiParagraph(section.ResolveEvenFooter())) {
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

        private static bool ApplyNativeDefaultFont(
            WordDocument document,
            NativeDocumentDefaults defaults,
            WordToPdfOptions? options,
            PdfCore.PdfOptions pdfOptions,
            bool allowSystemFontEmbedding,
            bool allowDocumentFontEmbedding,
            NativeFontMap nativeFontMap) {
            string? optionFontFamily = options?.FontFamily;
            if (!string.IsNullOrWhiteSpace(optionFontFamily) &&
                TryApplyNativeDefaultFontCandidate(optionFontFamily, pdfOptions, embedSystemFont: allowSystemFontEmbedding,
                    replaceBodyResources: true)) {
                RegisterAppliedNativeDefaultFont(optionFontFamily!, pdfOptions, nativeFontMap, allowSystemFontEmbedding);
                return true;
            }
            if (!string.IsNullOrWhiteSpace(optionFontFamily) && options?.PdfOptions != null) {
                return false;
            }

            foreach (string? family in EnumerateNativeFontFamilies(defaults.FontFamily, defaults.FontFamilies).Concat(new[] {
                document.Settings.FontFamily,
                document.Settings.FontFamilyHighAnsi,
                document.Settings.FontFamilyEastAsia,
                document.Settings.FontFamilyComplexScript
            })) {
                // Keep an available named source family on the run-resolution path.
                // Mapping it to a standard slot would suppress inherited named text.
                if (!string.IsNullOrWhiteSpace(family) &&
                    (pdfOptions.TryResolveNamedOfficeFontFamily(OfficeFontFamilyParser.Quote(family!), out _) ||
                     pdfOptions.TryResolveEmbeddedFontFamilySlot(family!, out _))) {
                    return false;
                }
                if (!string.IsNullOrWhiteSpace(family) && TryRegisterNativeDeclaredFontCandidate(
                        family!, pdfOptions, allowDocumentFontEmbedding, nativeFontMap)) {
                    return false;
                }
                if (!string.IsNullOrWhiteSpace(family) &&
                    TryApplyNativeDefaultFontCandidate(OfficeFontFamilyParser.Quote(family!), pdfOptions, embedSystemFont: allowDocumentFontEmbedding)) {
                    RegisterAppliedNativeDefaultFont(family!, pdfOptions, nativeFontMap, allowDocumentFontEmbedding);
                    return true;
                }
            }

            return false;
        }

        private static void RegisterAppliedNativeDefaultFont(string familyName, PdfCore.PdfOptions pdfOptions, NativeFontMap nativeFontMap, bool sourceFontEmbeddingAllowed) {
            nativeFontMap.Register(familyName, pdfOptions.DefaultFont);
            nativeFontMap.PreferPdfDefaultForDocumentDefaultFont();

            bool representedExactly =
                EmbeddedFontSlotMatchesFamily(pdfOptions, pdfOptions.DefaultFont, familyName) ||
                (!pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont) &&
                 PdfCore.PdfStandardFontMapper.IsStandardPdfFamilyEquivalent(familyName, pdfOptions.DefaultFont));
            if (!representedExactly) {
                nativeFontMap.ReportFontSubstitution(
                    pdfOptions,
                    familyName,
                    pdfOptions.DefaultFont,
                    GetEmbeddedFontFamilyName(pdfOptions, pdfOptions.DefaultFont),
                    sourceFontEmbeddingAllowed);
            }
        }

        private static void ApplyNativeTextFallbacks(
            WordDocument document,
            WordToPdfOptions? options,
            PdfCore.PdfOptions pdfOptions,
            HashSet<PdfCore.PdfStandardFont> reservedFontSlots,
            bool allowSystemFontEmbedding,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            if (options == null ||
                !allowSystemFontEmbedding ||
                options.TextFallbacks == PdfCore.PdfTextFallbackFeatures.None) {
                return;
            }

            PdfCore.PdfTextFallbackFeatures fallbackFeatures = options.TextFallbacks;
            if (document.Charts.Count == 0 && document.SmartArts.Count == 0) {
                fallbackFeatures = PdfCore.PdfTextDiagnostics.ResolveRequiredFallbackFeatures(
                    fallbackFeatures,
                    EnumerateNativeDocumentFallbackText(document, listMarkers));
            }
            if (fallbackFeatures == PdfCore.PdfTextFallbackFeatures.None) {
                return;
            }

            bool preserveCallerFontSlots =
                options.HasExplicitPdfFontConfiguration ||
                !string.IsNullOrWhiteSpace(options.FontFamily);
            if (preserveCallerFontSlots) {
                fallbackFeatures &= ~PdfCore.PdfTextFallbackFeatures.DocumentFont;
            }

            if (fallbackFeatures != PdfCore.PdfTextFallbackFeatures.None) {
                pdfOptions.UseTextFallbacks(
                    fallbackFeatures,
                    reservedFontSlots,
                    allowSystemFontEmbedding,
                    preserveCallerFontSlots,
                    requiredText: string.Join("\n", EnumerateNativeDocumentFallbackText(document, listMarkers)));
                foreach (PdfCore.PdfStandardFont slot in pdfOptions.EmbeddedFontFallbacks?.FontSlots ?? Array.Empty<PdfCore.PdfStandardFont>()) {
                    PdfCore.PdfOptions.AddRegisteredFontFamilySlot(reservedFontSlots, slot);
                }
            }
        }

        private static IEnumerable<string?> EnumerateNativeDocumentFallbackText(
            WordDocument document,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            yield return document._document.InnerText;
            yield return document._wordprocessingDocument.MainDocumentPart?.FootnotesPart?.Footnotes?.InnerText;
            yield return document._wordprocessingDocument.MainDocumentPart?.EndnotesPart?.Endnotes?.InnerText;

            foreach ((int Level, string Marker) marker in listMarkers.Values) {
                yield return marker.Marker;
            }

            foreach (WordSection section in document.Sections) {
                foreach (WordHeaderFooter? headerFooter in new WordHeaderFooter?[] {
                             section.ResolveDefaultHeader(),
                             section.ResolveFirstHeader(),
                             section.ResolveEvenHeader(),
                             section.ResolveDefaultFooter(),
                             section.ResolveFirstFooter(),
                             section.ResolveEvenFooter()
                         }) {
                    NativeHeaderFooterText? text = GetNativeHeaderFooterText(headerFooter, listMarkers);
                    yield return text?.Left;
                    yield return text?.Center;
                    yield return text?.Right;
                }

                foreach (WordWatermark watermark in section.Watermarks) {
                    yield return watermark.Text;
                }
            }
        }

        private static bool TryApplyNativeDefaultFontCandidate(string? familyName, PdfCore.PdfOptions pdfOptions, bool embedSystemFont,
            bool requireEmbeddedFont = false, bool replaceBodyResources = false) {
            if (string.IsNullOrWhiteSpace(familyName)) {
                return false;
            }

            return pdfOptions.TryUseInheritedOfficeFontFamily(familyName, embedSystemFont, requireEmbeddedFont, replaceBodyResources);
        }

        private static HashSet<PdfCore.PdfStandardFont> RegisterNativeDocumentFonts(WordDocument document, PdfCore.PdfOptions pdfOptions, bool preserveConfiguredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots = pdfOptions.CreateRegisteredFontFamilySlots(preserveConfiguredFontSlots);
            if (!nativeFontMap.UsePdfDefaultForDocumentDefaultFont) {
                NativeDocumentDefaults defaults = GetNativeDocumentDefaults(document);
                RegisterNativeFontCandidates(EnumerateNativeFontFamilies(defaults.FontFamily, defaults.FontFamilies),
                    pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            }

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

                RegisterNativeHeaderFooterFonts(section.ResolveDefaultHeader(), pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                RegisterNativeHeaderFooterFonts(section.ResolveFirstHeader(), pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                RegisterNativeHeaderFooterFonts(section.ResolveEvenHeader(), pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                RegisterNativeHeaderFooterFonts(section.ResolveDefaultFooter(), pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                RegisterNativeHeaderFooterFonts(section.ResolveFirstFooter(), pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                RegisterNativeHeaderFooterFonts(section.ResolveEvenFooter(), pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);

                foreach (WordWatermark watermark in section.Watermarks) {
                    RegisterNativeFontCandidate(watermark.FontFamily, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                }
            }

            RegisterNativePageFootnoteFonts(document, pdfOptions, registeredFamilies, registeredFontSlots,
                allowSystemFontEmbedding, nativeFontMap);

            return registeredFontSlots;
        }

        private static void RegisterNativeHeaderFooterFonts(WordHeaderFooter? headerFooter, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            if (headerFooter == null) {
                return;
            }

            if (!headerFooter.ChildElements.Any()) {
                RegisterNativeElementFonts(CreateNativeImplicitRunningParagraph(headerFooter), pdfOptions,
                    registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                return;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                RegisterNativeElementFonts(element, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            }
        }

        private static void RegisterNativeElementFonts(WordElement element, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            if (element is WordParagraph paragraph) {
                RegisterNativeParagraphContentFonts(
                    paragraph,
                    NativeTableRunStyleDefaults.Empty,
                    pdfOptions,
                    registeredFamilies,
                    registeredFontSlots,
                    allowSystemFontEmbedding,
                    nativeFontMap);
            } else if (element is WordTable table) {
                RegisterNativeTableFonts(table, pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            }
        }

        private static void RegisterNativeTableFonts(WordTable table, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            foreach (WordTable currentTable in EnumerateNativeTableTree(table)) {
                NativeTableStyleDefaults tableStyleDefaults = GetNativeTableStyleDefaults(
                    currentTable,
                    GetNativeDocumentDefaults(currentTable.Document),
                    ignoreFallbackTableStyle: pdfOptions.HasExplicitDefaultTableStyle);

                TableLayout layout = TableLayoutCache.GetLayout(currentTable);
                int columnCount = GetNativeTableColumnCount(layout);
                int headerCount = GetNativeTableVisualHeaderRowCount(currentTable, layout.Rows.Count,
                    GetNativeTableRepeatedHeaderRowCount(currentTable, layout.Rows.Count));
                int footerStart = currentTable.ConditionalFormattingLastRow == true && layout.Rows.Count > headerCount
                    ? layout.Rows.Count - 1 : layout.Rows.Count;
                for (int rowIndex = 0; rowIndex < layout.Rows.Count; rowIndex++) {
                    int columnIndex = GetNativeTableRowStartColumn(layout, rowIndex);
                    foreach (WordTableCell cell in layout.Rows[rowIndex]) {
                        if (IsNativeHorizontalMergeContinuation(cell)) continue;
                        int span = GetNativeCellColumnSpan(cell);
                        if (!IsNativeVerticalMergeContinuation(cell)) {
                            NativeTableStyleDefaults cellStyle = GetNativeTableCellStyleDefaults(currentTable,
                                tableStyleDefaults, rowIndex, columnIndex, span, columnCount, headerCount, footerStart);
                            foreach (WordParagraph paragraph in cell.Paragraphs) {
                                RegisterNativeParagraphContentFonts(paragraph, cellStyle.RunStyle, pdfOptions,
                                    registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
                            }
                        }
                        columnIndex += span;
                    }
                }
            }
        }

        private static void RegisterNativeParagraphContentFonts(
            WordParagraph paragraph,
            NativeTableRunStyleDefaults tableRunStyleDefaults,
            PdfCore.PdfOptions pdfOptions,
            HashSet<string> registeredFamilies,
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
            bool allowSystemFontEmbedding,
            NativeFontMap nativeFontMap) {
            List<WordParagraph> runs = GetNativeRuns(paragraph, includeInlineContentControls: true);
            W.RunFonts? markFonts = GetNativeEmptyParagraphMarkFonts(paragraph, runs);
            if (markFonts != null) {
                RegisterNativeFontCandidates(EnumerateNativeLatinFontFamilies(paragraph._document, markFonts),
                    pdfOptions, registeredFamilies, registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            }
            if (runs.Count == 0) {
                RegisterNativeEffectiveParagraphFont(
                    paragraph,
                    tableRunStyleDefaults,
                    pdfOptions,
                    registeredFamilies,
                    registeredFontSlots,
                    allowSystemFontEmbedding,
                    nativeFontMap);
            } else {
                foreach (WordParagraph run in runs) {
                    RegisterNativeEffectiveParagraphFont(
                        run,
                        tableRunStyleDefaults,
                        pdfOptions,
                        registeredFamilies,
                        registeredFontSlots,
                        allowSystemFontEmbedding,
                        nativeFontMap);
                }
            }

            // Both authored states belong to the document even though only one
            // has a cached visible run. Filling the PDF selects its other AP.
            foreach (string family in GetNativeCheckBoxAppearanceFontFamilies(paragraph)) {
                RegisterNativeFontCandidate(family, pdfOptions, registeredFamilies,
                    registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            }

            WordDocumentTraversal.ListInfo? listInfo = WordDocumentTraversal.GetListInfo(paragraph);
            if (listInfo?.LevelSuffix == WordListLevelSuffix.Space) {
                RegisterNativeFontCandidate("Arial", pdfOptions, registeredFamilies,
                    registeredFontSlots, allowSystemFontEmbedding, nativeFontMap);
            }
            if (listInfo.HasValue &&
                !ShouldUseNativeListTextFontForNormalizedMarker(
                    listInfo.Value,
                    NormalizeNativeBulletMarker(listInfo.Value.LevelText ?? string.Empty))) {
                RegisterNativeFontCandidate(
                    listInfo.Value.MarkerFontFamily,
                    pdfOptions,
                    registeredFamilies,
                    registeredFontSlots,
                    allowSystemFontEmbedding,
                    nativeFontMap);
            }
        }

        private static void RegisterNativeEffectiveParagraphFont(
            WordParagraph paragraph,
            NativeTableRunStyleDefaults tableRunStyleDefaults,
            PdfCore.PdfOptions pdfOptions,
            HashSet<string> registeredFamilies,
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
            bool allowSystemFontEmbedding,
            NativeFontMap nativeFontMap) {
            NativeCharacterStyleDefaults characterStyleDefaults =
                GetNativeCharacterStyleDefaults(paragraph._document, GetNativeRunProperties(paragraph));
            NativeParagraphStyleDefaults paragraphStyleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            RegisterNativeFontCandidates(
                EnumerateNativeParagraphOwnFontFamilies(paragraph).Concat(EnumerateNativeStyleFontFamilies(
                    characterStyleDefaults, paragraphStyleDefaults, tableRunStyleDefaults, default, includeDocument: false)),
                pdfOptions,
                registeredFamilies,
                registeredFontSlots,
                allowSystemFontEmbedding,
                nativeFontMap);
        }

        private static void RegisterNativeFontCandidate(string? familyName, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots, bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            if (!PdfCore.PdfOptions.TryAddOfficeFontFamilyKey(familyName, registeredFamilies, NormalizeNativeFontFamily, out string trimmedFamilyName)) {
                return;
            }

            string literalFamily = OfficeFontFamilyParser.Quote(trimmedFamilyName);
            if (pdfOptions.HasRenderingProfileFamilyPlanner(literalFamily)) {
                nativeFontMap.RegisterNamed(trimmedFamilyName, trimmedFamilyName);
                return;
            }

            if (pdfOptions.TryResolveFontFamilySubstitution(
                    literalFamily,
                    out PdfCore.PdfFontFamilySubstitution? substitution) &&
                substitution != null) {
                nativeFontMap.RegisterNamed(trimmedFamilyName, substitution.TargetFontFamily);
                nativeFontMap.ReportFontSubstitution(
                    pdfOptions,
                    trimmedFamilyName,
                    fallbackSlot: null,
                    substitution.TargetFontFamily,
                    sourceFontFamilyRequest: literalFamily);
                return;
            }

            if (pdfOptions.TryResolveNamedOfficeFontFamily(
                    literalFamily,
                    out string? registeredNamedFamily) &&
                registeredNamedFamily != null) {
                nativeFontMap.RegisterNamed(trimmedFamilyName, registeredNamedFamily);
                return;
            }

            if (pdfOptions.TryResolveEmbeddedFontFamilySlot(trimmedFamilyName, out PdfCore.PdfStandardFont embeddedSlot)) {
                registeredFontSlots.Add(embeddedSlot);
                nativeFontMap.Register(trimmedFamilyName, embeddedSlot);
                return;
            }

            if (TryRegisterNativeDeclaredFontCandidate(trimmedFamilyName, pdfOptions, allowSystemFontEmbedding, nativeFontMap)) {
                if (nativeFontMap.TryGetFontSlot(trimmedFamilyName, out PdfCore.PdfStandardFont declaredSlot)) {
                    registeredFontSlots.Add(declaredSlot);
                }
                return;
            }

            if (allowSystemFontEmbedding &&
                pdfOptions.TryRegisterNamedOfficeFontFamily(literalFamily, out string? registeredFamilyName) &&
                registeredFamilyName != null) {
                nativeFontMap.RegisterNamed(trimmedFamilyName, registeredFamilyName);
                return;
            }

            if (allowSystemFontEmbedding &&
                PdfCore.PdfStandardFontMapper.TryMapFontFamily(literalFamily, out PdfCore.PdfStandardFont mappedFont) &&
                registeredFontSlots.Contains(PdfCore.PdfStandardFontMapper.GetFontFamily(mappedFont)) &&
                !EmbeddedFontSlotMatchesFamily(pdfOptions, mappedFont, trimmedFamilyName) &&
                PdfCore.PdfEmbeddedFontFamily.TryFromSystem(trimmedFamilyName, out PdfCore.PdfEmbeddedFontFamily? distinctEmbeddedFamily) &&
                distinctEmbeddedFamily != null) {
                if (PdfCore.PdfOptions.TrySelectAvailableFontFamilySlot(trimmedFamilyName, registeredFontSlots, out PdfCore.PdfStandardFont distinctFontSlot)) {
                    registeredFontSlots.Add(distinctFontSlot);
                    pdfOptions.RegisterFontFamily(distinctFontSlot, distinctEmbeddedFamily);
                    nativeFontMap.Register(trimmedFamilyName, distinctFontSlot);
                    return;
                }

                nativeFontMap.ReportSlotExhaustion(
                    trimmedFamilyName,
                    mappedFont,
                    GetEmbeddedFontFamilyName(pdfOptions, mappedFont));
            }

            if (pdfOptions.TryRegisterMappedOfficeFontFamily(literalFamily, registeredFontSlots, allowSystemFontEmbedding, out PdfCore.PdfStandardFont fontFamily)) {
                nativeFontMap.Register(trimmedFamilyName, fontFamily);
                bool representedExactly =
                    EmbeddedFontSlotMatchesFamily(pdfOptions, fontFamily, trimmedFamilyName) ||
                    (!pdfOptions.HasEmbeddedStandardFontFamily(fontFamily) &&
                     PdfCore.PdfStandardFontMapper.IsStandardPdfFamilyEquivalent(literalFamily, fontFamily));
                if (!representedExactly) {
                    nativeFontMap.ReportFontSubstitution(
                        pdfOptions,
                        trimmedFamilyName,
                        fontFamily,
                        GetEmbeddedFontFamilyName(pdfOptions, fontFamily),
                        sourceFontEmbeddingAllowed: allowSystemFontEmbedding);
                }
                return;
            }

            if (allowSystemFontEmbedding &&
                PdfCore.PdfOptions.TrySelectAvailableFontFamilySlot(trimmedFamilyName, registeredFontSlots, out PdfCore.PdfStandardFont fontSlot) &&
                PdfCore.PdfEmbeddedFontFamily.TryFromSystem(trimmedFamilyName, out PdfCore.PdfEmbeddedFontFamily? embeddedFamily) &&
                embeddedFamily != null) {
                registeredFontSlots.Add(fontSlot);
                pdfOptions.RegisterFontFamily(fontSlot, embeddedFamily);
                nativeFontMap.Register(trimmedFamilyName, fontSlot);
                return;
            }

            PdfCore.PdfStandardFont fallback = PdfCore.PdfStandardFontMapper.TryMapFontFamily(literalFamily, out PdfCore.PdfStandardFont mappedFallback)
                ? mappedFallback
                : PdfCore.PdfStandardFont.Helvetica;
            nativeFontMap.ReportFontSubstitution(
                pdfOptions,
                trimmedFamilyName,
                fallback,
                GetEmbeddedFontFamilyName(pdfOptions, fallback),
                sourceFontEmbeddingAllowed: allowSystemFontEmbedding);
        }

        /// <summary>Tries the literal source, explicit fallback list and declared alternates before implicit substitution.</summary>
        private static bool TryRegisterNativeDeclaredFontCandidate(string sourceFamily, PdfCore.PdfOptions pdfOptions,
            bool allowSystemFontEmbedding, NativeFontMap nativeFontMap) {
            IReadOnlyList<string> alternates = nativeFontMap.GetDeclaredFontAlternates(sourceFamily);
            IReadOnlyList<string> listedFamilies = OfficeFontFamilyParser.Parse(sourceFamily, allowSemicolonSeparators: true);
            if (alternates.Count == 0 && listedFamilies.Count == 1 &&
                string.Equals(listedFamilies[0], sourceFamily, StringComparison.Ordinal)) return false;
            // A registered family may itself contain delimiters. Only consider the legacy
            // list interpretation after the complete literal name has proved unavailable.
            foreach (string candidate in new[] { sourceFamily }.Concat(listedFamilies).Concat(alternates)
                         .Distinct(StringComparer.OrdinalIgnoreCase)) {
                string quoted = OfficeFontFamilyParser.Quote(candidate);
                // An explicitly listed source candidate retains the existing list diagnostics.
                string selectedSourceFamily = listedFamilies.Contains(candidate, StringComparer.OrdinalIgnoreCase)
                    ? candidate : sourceFamily;
                string? named = null;
                if (pdfOptions.HasRenderingProfileFamilyPlanner(quoted)) named = candidate;
                else if (pdfOptions.TryResolveNamedOfficeFontFamily(quoted, out string? registered)) named = registered;
                else if (pdfOptions.TryResolveEmbeddedFontFamilySlot(candidate, out PdfCore.PdfStandardFont embeddedSlot)) {
                    nativeFontMap.Register(sourceFamily, embeddedSlot);
                    if (!string.Equals(NormalizeNativeFontFamily(selectedSourceFamily), NormalizeNativeFontFamily(candidate),
                            StringComparison.OrdinalIgnoreCase)) {
                        nativeFontMap.ReportFontSubstitution(pdfOptions, sourceFamily, embeddedSlot, candidate,
                            sourceFontEmbeddingAllowed: allowSystemFontEmbedding);
                    }
                    return true;
                } else if (allowSystemFontEmbedding &&
                         PdfCore.PdfEmbeddedFontFamily.TryFromSystem(candidate, out PdfCore.PdfEmbeddedFontFamily? installed) &&
                         installed != null && pdfOptions.TryRegisterNamedFontFamily(installed)) named = installed.FamilyName;
                if (named != null) {
                    nativeFontMap.RegisterNamed(sourceFamily, named);
                    if (!string.Equals(NormalizeNativeFontFamily(selectedSourceFamily), NormalizeNativeFontFamily(named),
                            StringComparison.OrdinalIgnoreCase)) {
                        nativeFontMap.ReportFontSubstitution(pdfOptions, sourceFamily, null, named,
                            sourceFontEmbeddingAllowed: allowSystemFontEmbedding);
                    }
                    return true;
                }
                if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(quoted, out PdfCore.PdfStandardFont standard) &&
                    PdfCore.PdfStandardFontMapper.IsStandardPdfFamilyEquivalent(quoted, standard) &&
                    (!pdfOptions.HasEmbeddedStandardFontFamily(standard) || EmbeddedFontSlotMatchesFamily(pdfOptions, standard, candidate))) {
                    nativeFontMap.Register(sourceFamily, standard);
                    if (!string.Equals(NormalizeNativeFontFamily(selectedSourceFamily), NormalizeNativeFontFamily(candidate),
                            StringComparison.OrdinalIgnoreCase)) {
                        nativeFontMap.ReportFontSubstitution(pdfOptions, sourceFamily, standard, candidate,
                            sourceFontEmbeddingAllowed: allowSystemFontEmbedding);
                    }
                    return true;
                }
            }
            return false;
        }

        internal static bool EmbeddedFontSlotMatchesFamily(PdfCore.PdfOptions options, PdfCore.PdfStandardFont slot, string familyName) {
            return options.EmbeddedFontFamilySlotMatches(slot, familyName);
        }

        private static string? GetEmbeddedFontFamilyName(PdfCore.PdfOptions options, PdfCore.PdfStandardFont slot) {
            PdfCore.PdfStandardFont normalizedSlot = PdfCore.PdfStandardFontMapper.GetFontFamily(slot);
            return options.GetEmbeddedFontFamilyName(normalizedSlot);
        }

    }
}
