using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static PdfCore.PdfOptions CreatePdfOptions(ExcelPdfSaveOptions options, out bool preserveConfiguredFontSlots) {
            PdfCore.PdfOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfCore.PdfOptions();
            pdfOptions.ReportDiagnosticsTo(options.Report, "OfficeIMO.Excel.Pdf");

            pdfOptions.CreateOutlineFromHeadings = true;
            preserveConfiguredFontSlots = options.HasExplicitPdfFontConfiguration;

            if (!string.IsNullOrWhiteSpace(options.FontFamily) &&
                TryApplyPdfFontFamily(options.FontFamily, pdfOptions, options.ResourcePolicy.AllowSystemFontEmbedding)) {
                preserveConfiguredFontSlots = true;
            }

            if (options.PageSize.HasValue) {
                pdfOptions.PageSize = options.PageSize.Value;
            }

            if (options.Margins.HasValue) {
                pdfOptions.Margins = options.Margins.Value;
            }

            return pdfOptions;
        }

        private static void ApplyTextFallbacks(
            PdfCore.PdfOptions pdfOptions,
            ExcelPdfSaveOptions options,
            bool preserveConfiguredFontSlots,
            IEnumerable<PdfCore.PdfStandardFont> reservedFontSlots) {
            if (!options.ResourcePolicy.AllowSystemFontEmbedding ||
                options.TextFallbacks == PdfCore.PdfTextFallbackFeatures.None) {
                return;
            }

            PdfCore.PdfTextFallbackFeatures fallbackFeatures = options.TextFallbacks;
            if (preserveConfiguredFontSlots || pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont)) {
                fallbackFeatures &= ~PdfCore.PdfTextFallbackFeatures.DocumentFont;
            }

            pdfOptions.UseTextFallbacks(
                fallbackFeatures,
                reservedFontSlots,
                options.ResourcePolicy.AllowSystemFontEmbedding,
                preserveConfiguredFontSlots);
        }

        private static bool TryApplyPdfFontFamily(string? familyName, PdfCore.PdfOptions pdfOptions, bool embedSystemFont, bool requireEmbeddedFont = false) {
            return pdfOptions.TryUseOfficeFontFamily(familyName, embedSystemFont, requireEmbeddedFont);
        }

        private static HashSet<PdfCore.PdfStandardFont> RegisterWorksheetFonts(PdfCore.PdfOptions pdfOptions, IReadOnlyList<WorksheetPdfExportPlan> exportPlans, ExcelPdfSaveOptions options, bool preserveConfiguredFontSlots) {
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots = pdfOptions.CreateRegisteredFontFamilySlots(preserveConfiguredFontSlots);

            var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            RegisterWorksheetHeaderFooterFonts(exportPlans, options, registeredFamilies, registeredFontSlots, pdfOptions);

            if (options.UseWorksheetCellStyles) {
                foreach (WorksheetPdfExportPlan plan in exportPlans) {
                    ExcelCellStyleSnapshot?[,]? styles = plan.ExportData.Styles;
                    if (styles == null) {
                        continue;
                    }

                    int rows = Math.Min(plan.ExportedRows, styles.GetLength(0));
                    int columns = styles.GetLength(1);
                    for (int row = 0; row < rows; row++) {
                        for (int column = 0; column < columns; column++) {
                            ExcelCellStyleSnapshot? style = styles[row, column];
                            RegisterWorksheetFontCandidate(
                                style?.FontName,
                                pdfOptions,
                                registeredFamilies,
                                registeredFontSlots,
                                options,
                                plan.SheetName,
                                reportSubstitution: style?.IsFontFamilyExplicit == true || !preserveConfiguredFontSlots);
                        }
                    }
                }
            }

            return registeredFontSlots;
        }

        private static void RegisterWorksheetHeaderFooterFonts(
            IReadOnlyList<WorksheetPdfExportPlan> exportPlans,
            ExcelPdfSaveOptions options,
            HashSet<string> registeredFamilies,
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
            PdfCore.PdfOptions pdfOptions) {
            if (!options.UseWorksheetHeadersAndFooters) {
                return;
            }

            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                ExcelSheet.ExcelHeaderFooterSnapshot? headerFooter = plan.HeaderFooter;
                if (headerFooter == null) {
                    continue;
                }

                foreach (string? text in EnumerateHeaderFooterTextZones(headerFooter)) {
                    RegisterWorksheetHeaderFooterFontCandidate(text, pdfOptions, registeredFamilies, registeredFontSlots, options, plan.SheetName);
                }
            }
        }

        private static IEnumerable<string?> EnumerateHeaderFooterTextZones(ExcelSheet.ExcelHeaderFooterSnapshot headerFooter) {
            yield return headerFooter.HeaderLeft;
            yield return headerFooter.HeaderCenter;
            yield return headerFooter.HeaderRight;
            yield return headerFooter.FooterLeft;
            yield return headerFooter.FooterCenter;
            yield return headerFooter.FooterRight;
            yield return headerFooter.FirstHeaderLeft;
            yield return headerFooter.FirstHeaderCenter;
            yield return headerFooter.FirstHeaderRight;
            yield return headerFooter.FirstFooterLeft;
            yield return headerFooter.FirstFooterCenter;
            yield return headerFooter.FirstFooterRight;
            yield return headerFooter.EvenHeaderLeft;
            yield return headerFooter.EvenHeaderCenter;
            yield return headerFooter.EvenHeaderRight;
            yield return headerFooter.EvenFooterLeft;
            yield return headerFooter.EvenFooterCenter;
            yield return headerFooter.EvenFooterRight;
        }

        private static void RegisterWorksheetHeaderFooterFontCandidate(
            string? text,
            PdfCore.PdfOptions pdfOptions,
            HashSet<string> registeredFamilies,
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
            ExcelPdfSaveOptions options,
            string sheetName) {
            if (string.IsNullOrWhiteSpace(text)) {
                return;
            }

            for (int i = 0; i < text!.Length - 1; i++) {
                if (text[i] == '&' && text[i + 1] == '&') {
                    i++;
                    continue;
                }

                if (text[i] != '&' || text[i + 1] != '"') {
                    continue;
                }

                if (!TryReadHeaderFooterQuotedToken(text, i + 1, out string token, out int endIndex)) {
                    continue;
                }

                string familyName = token.Split(new[] { ',' }, 2)[0];
                RegisterWorksheetFontCandidate(familyName, pdfOptions, registeredFamilies, registeredFontSlots, options, sheetName);
                i = endIndex;
            }
        }

        private static void RegisterWorksheetFontCandidate(
            string? familyName,
            PdfCore.PdfOptions pdfOptions,
            HashSet<string> registeredFamilies,
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
            ExcelPdfSaveOptions options,
            string sheetName,
            bool reportSubstitution = true) {
            if (PdfCore.PdfOptions.TryAddOfficeFontFamilyKey(familyName, registeredFamilies, normalizeKey: null, out string trimmedFamilyName)) {
                if (pdfOptions.HasRenderingProfileFamilyPlanner(trimmedFamilyName)) {
                    return;
                }

                if (pdfOptions.TryResolveFontFamilySubstitution(
                        trimmedFamilyName,
                        out PdfCore.PdfFontFamilySubstitution? substitution) &&
                    substitution != null) {
                    options.Report.Add(pdfOptions.CreateFontFamilySubstitutionWarning(
                        "OfficeIMO.Excel.Pdf",
                        "WorksheetFontFamilySubstituted",
                        sheetName,
                        trimmedFamilyName,
                        fallbackSlot: null,
                        resolvedFontFamily: substitution.TargetFontFamily,
                        additionalDetails: new Dictionary<string, string> {
                            ["sheetName"] = sheetName
                        }));
                    return;
                }

                if (pdfOptions.HasNamedFontFamily(trimmedFamilyName)) {
                    return;
                }

                bool embedSystemFont = options.ResourcePolicy.AllowSystemFontEmbedding &&
                    options.ResourcePolicy.AllowDocumentFontEmbedding;
                if (embedSystemFont && pdfOptions.TryRegisterNamedOfficeFontFamily(trimmedFamilyName, out _)) {
                    return;
                }

                bool mapped = pdfOptions.TryRegisterMappedOfficeFontFamily(
                    trimmedFamilyName,
                    registeredFontSlots,
                    embedSystemFont,
                    out PdfCore.PdfStandardFont fallback);
                bool representedExactly =
                    mapped &&
                    (pdfOptions.EmbeddedFontFamilySlotMatches(fallback, trimmedFamilyName) ||
                     (!pdfOptions.HasEmbeddedStandardFontFamily(fallback) &&
                      PdfCore.PdfStandardFontMapper.IsStandardPdfFamilyEquivalent(trimmedFamilyName, fallback)));
                if (reportSubstitution &&
                    !representedExactly) {
                    PdfCore.PdfStandardFont reportedFallback = mapped
                        ? fallback
                        : PdfCore.PdfStandardFont.Helvetica;
                    options.Report.Add(pdfOptions.CreateFontFamilySubstitutionWarning(
                        "OfficeIMO.Excel.Pdf",
                        "WorksheetFontFamilySubstituted",
                        sheetName,
                        trimmedFamilyName,
                        reportedFallback,
                        pdfOptions.GetEmbeddedFontFamilyName(reportedFallback),
                        new Dictionary<string, string> {
                            ["sheetName"] = sheetName
                        }));
                }
            }
        }

        private static IReadOnlyList<WorksheetPdfExportPlan> BuildWorksheetExportPlans(ExcelDocument document, ExcelDocumentReader reader, IReadOnlyList<string> sheetNames, ExcelPdfSaveOptions options, bool hasExplicitSheetSelection, PdfCore.PdfStandardFont defaultFontFamily) {
            var plans = new List<WorksheetPdfExportPlan>();
            for (int i = 0; i < sheetNames.Count; i++) {
                string sheetName = sheetNames[i];
                ExcelSheet? workbookSheet = GetWorkbookSheet(document, sheetName);
                if (ShouldSkipWorkbookSheet(workbookSheet, options, hasExplicitSheetSelection)) {
                    continue;
                }

                ExcelSheetReader sheet = reader.GetSheet(sheetName);
                bool boundedRead = IsBoundedWorksheetRead(options);
                ExcelSheet? metadataSheet = boundedRead ? null : workbookSheet;
                ExcelSheetPageSetup? pageSetup = options.UseWorksheetPageSetup ? metadataSheet?.GetPageSetup() : null;
                ExcelSheet.ExcelHeaderFooterSnapshot? headerFooter = (options.UseWorksheetHeadersAndFooters || options.UseWorksheetHeaderFooterImages) ? metadataSheet?.GetHeaderFooter() : null;
                string exportRange = GetExportRange(sheet, workbookSheet, options);
                string normalizedExportRange = NormalizeA1Range(exportRange);
                int sourceRows = GetRangeRowCount(normalizedExportRange);
                string materializedRange = options.MaxRowsPerSheet.HasValue
                    ? LimitRangeRows(normalizedExportRange, options.MaxRowsPerSheet.Value)
                    : normalizedExportRange;
                SheetExportData exportData = ReadSheetExportData(document, sheet, workbookSheet, materializedRange, options, defaultFontFamily);
                IReadOnlyList<int> manualRowBreaks = options.UseWorksheetPageBreaks && metadataSheet != null
                    ? metadataSheet.GetManualRowPageBreaks()
                    : Array.Empty<int>();
                IReadOnlyList<int> manualColumnBreaks = options.UseWorksheetPageBreaks && metadataSheet != null
                    ? metadataSheet.GetManualColumnPageBreaks()
                    : Array.Empty<int>();
                object?[,] values = exportData.Values;
                int rows = values.GetLength(0);
                int columns = values.GetLength(1);
                bool hasTable = rows > 0 && columns > 0 && HasWorksheetTableContent(exportData);
                int exportedRows = options.MaxRowsPerSheet.HasValue
                    ? Math.Min(rows, options.MaxRowsPerSheet.Value)
                    : rows;
                if (options.MaxRowsPerSheet.HasValue && sourceRows > exportedRows) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetRows",
                        $"Worksheet export was truncated from {sourceRows.ToString(CultureInfo.InvariantCulture)} to {exportedRows.ToString(CultureInfo.InvariantCulture)} rows because MaxRowsPerSheet is set.");
                }
                if (boundedRead) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetBoundedRead",
                        "The bounded worksheet path preserved values, direct cell styles, and built-in Excel table styles without materializing the full worksheet. Print settings, repeated titles, merges, hyperlinks, conditional formatting, row and column sizing, and worksheet media were not inspected.");
                }

                ISet<string>? exportedCellReferences = CreateExportedCellReferenceSet(exportData.CellReferences, exportedRows);
                bool filterMediaToExportedCells = !boundedRead && HasWorksheetPrintArea(workbookSheet, options) ||
                                                  options.MaxRowsPerSheet.HasValue ||
                                                  (options.RespectWorksheetHiddenRowsAndColumns && HasHiddenRowsOrColumns(workbookSheet));
                IReadOnlyList<WorksheetImageExportData> images = boundedRead
                    ? Array.Empty<WorksheetImageExportData>()
                    : FilterImagesByExportedCells(ReadWorksheetImages(workbookSheet, options, sheetName), exportedCellReferences, filterMediaToExportedCells);
                IReadOnlyList<WorksheetChartExportData> charts = boundedRead
                    ? Array.Empty<WorksheetChartExportData>()
                    : FilterChartsByExportedCells(ReadWorksheetCharts(workbookSheet, options, sheetName), exportedCellReferences, filterMediaToExportedCells);
                if (!hasTable && images.Count == 0 && charts.Count == 0) {
                    continue;
                }

                bool hasPrintArea = !boundedRead && HasWorksheetPrintArea(workbookSheet, options) && !ContainsMultiplePrintAreas(GetWorksheetPrintArea(workbookSheet, options)!);
                plans.Add(new WorksheetPdfExportPlan(
                    sheetName,
                    pageSetup,
                    headerFooter,
                    exportData,
                    images,
                    charts,
                    hasTable,
                    exportedRows,
                    manualRowBreaks,
                    manualColumnBreaks,
                    CreateSheetBookmarkName(sheetName, plans.Count + 1),
                    CreateWorksheetGeometry(metadataSheet, normalizedExportRange, options),
                    hasPrintArea || options.MaxRowsPerSheet.HasValue));
            }

            return plans;
        }

        private static bool HasWorksheetTableContent(SheetExportData exportData) {
            object?[,] values = exportData.Values;
            for (int row = 0; row < values.GetLength(0); row++) {
                for (int column = 0; column < values.GetLength(1); column++) {
                    object? value = values[row, column];
                    if (value is string text ? text.Length > 0 : value != null) {
                        return true;
                    }
                }
            }

            ExcelCellStyleSnapshot?[,]? styles = exportData.Styles;
            if (styles != null) {
                for (int row = 0; row < styles.GetLength(0); row++) {
                    for (int column = 0; column < styles.GetLength(1); column++) {
                        if (styles[row, column] is ExcelCellStyleSnapshot style && HasVisibleBlankCellStyle(style)) {
                            return true;
                        }
                    }
                }
            }

            string?[,]? cellReferences = exportData.CellReferences;
            if (cellReferences == null || exportData.StructuredTables.Count == 0) {
                return false;
            }

            for (int row = 0; row < cellReferences.GetLength(0); row++) {
                for (int column = 0; column < cellReferences.GetLength(1); column++) {
                    if (GetStructuredTableCellVisual(exportData.StructuredTables, cellReferences, row, column) != null) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasVisibleBlankCellStyle(ExcelCellStyleSnapshot style) =>
            style.FillColorArgb != null ||
            style.FillPatternType != null ||
            style.FillPatternForegroundColorArgb != null ||
            style.FillPatternBackgroundColorArgb != null ||
            style.FillGradientUnsupported ||
            style.FillGradientStartColorArgb != null ||
            style.FillGradientEndColorArgb != null ||
            style.FillGradientStops.Count > 0 ||
            style.FillGradientDegree.HasValue ||
            style.Border != null;

        private static int GetRangeRowCount(string normalizedRange) {
            return A1.TryParseRange(normalizedRange, out int firstRow, out _, out int lastRow, out _)
                ? Math.Max(0, lastRow - firstRow + 1)
                : 0;
        }

        private static string LimitRangeRows(string normalizedRange, int maximumRows) {
            if (maximumRows <= 0 ||
                !A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return normalizedRange;
            }

            int limitedLastRow = Math.Min(lastRow, firstRow + maximumRows - 1);
            return ToA1Range(firstRow, firstColumn, limitedLastRow, lastColumn);
        }

        private static WorksheetGeometryData CreateWorksheetGeometry(ExcelSheet? workbookSheet, string normalizedRange, ExcelPdfSaveOptions options) {
            A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn);
            return new WorksheetGeometryData(
                Math.Max(1, firstRow),
                Math.Max(1, firstColumn),
                Math.Max(Math.Max(1, firstRow), lastRow),
                Math.Max(Math.Max(1, firstColumn), lastColumn),
                workbookSheet?.DefaultColumnWidth is double defaultColumnWidth && defaultColumnWidth > 0D ? defaultColumnWidth : 8.43D,
                workbookSheet?.DefaultRowHeight is double defaultRowHeight && defaultRowHeight > 0D ? defaultRowHeight : 15D,
                workbookSheet?.GetColumnDefinitions() ?? Array.Empty<ExcelColumnSnapshot>(),
                workbookSheet?.GetRowDefinitions() ?? Array.Empty<ExcelRowSnapshot>(),
                options.UseWorksheetColumnWidths,
                options.UseWorksheetRowHeights,
                options.RespectWorksheetHiddenRowsAndColumns);
        }

        private static IReadOnlyDictionary<string, string> BuildSheetDestinationMap(IReadOnlyList<WorksheetPdfExportPlan> exportPlans) {
            var destinations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                destinations[plan.SheetName] = plan.BookmarkName;
            }

            return destinations;
        }

        private static IReadOnlyDictionary<string, string> BuildCellDestinationMap(IReadOnlyList<WorksheetPdfExportPlan> exportPlans) {
            var targetCells = CollectInternalHyperlinkTargetCells(exportPlans);
            if (targetCells.Count == 0) {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            var destinations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                string?[,]? references = plan.ExportData.CellReferences;
                if (references == null) {
                    continue;
                }

                int rows = Math.Min(plan.ExportedRows, references.GetLength(0));
                int columns = references.GetLength(1);
                for (int row = 0; row < rows; row++) {
                    for (int column = 0; column < columns; column++) {
                        string? cellReference = references[row, column];
                        if (string.IsNullOrWhiteSpace(cellReference)) {
                            continue;
                        }

                        string key = CreateCellDestinationKey(plan.SheetName, cellReference!);
                        if (targetCells.Contains(key)) {
                            destinations[key] = CreateCellBookmarkName(plan.BookmarkName, cellReference!);
                        }
                    }
                }
            }

            return destinations;
        }

        private static HashSet<string> CollectInternalHyperlinkTargetCells(IReadOnlyList<WorksheetPdfExportPlan> exportPlans) {
            var targetCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                ExcelHyperlinkSnapshot?[,]? hyperlinks = plan.ExportData.Hyperlinks;
                if (hyperlinks == null) {
                    continue;
                }

                int rows = Math.Min(plan.ExportedRows, hyperlinks.GetLength(0));
                int columns = hyperlinks.GetLength(1);
                for (int row = 0; row < rows; row++) {
                    for (int column = 0; column < columns; column++) {
                        ExcelHyperlinkSnapshot? hyperlink = hyperlinks[row, column];
                        if (hyperlink == null || hyperlink.IsExternal) {
                            continue;
                        }

                        if (TryParseInternalTarget(hyperlink.Target, plan.SheetName, out string? sheetName, out string? cellReference)) {
                            targetCells.Add(CreateCellDestinationKey(sheetName!, cellReference!));
                        }
                    }
                }
            }

            return targetCells;
        }

        private static string CreateSheetBookmarkName(string sheetName, int index) {
            var builder = new System.Text.StringBuilder();
            foreach (char value in sheetName) {
                if (char.IsLetterOrDigit(value)) {
                    builder.Append(char.ToLowerInvariant(value));
                } else if (builder.Length > 0 && builder[builder.Length - 1] != '-') {
                    builder.Append('-');
                }
            }

            string token = builder.ToString().Trim('-');
            if (token.Length > 40) {
                token = token.Substring(0, 40).Trim('-');
            }

            return token.Length == 0
                ? "excel-sheet-" + index.ToString(CultureInfo.InvariantCulture)
                : "excel-sheet-" + index.ToString(CultureInfo.InvariantCulture) + "-" + token;
        }

        private static string CreateCellBookmarkName(string sheetBookmarkName, string cellReference) {
            return sheetBookmarkName + "-" + cellReference.Replace("$", string.Empty).ToLowerInvariant();
        }

        private static string CreateCellDestinationKey(string sheetName, string cellReference) {
            return sheetName.Trim() + "!" + cellReference.Replace("$", string.Empty).ToUpperInvariant();
        }

    }
}
