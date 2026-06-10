using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static PdfCore.PdfOptions CreatePdfOptions(ExcelPdfSaveOptions options, out bool preserveConfiguredFontSlots) {
            PdfCore.PdfOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfCore.PdfOptions();
            pdfOptions.ReportDiagnosticsTo(options.ConversionReport, "OfficeIMO.Excel.Pdf");

            pdfOptions.CreateOutlineFromHeadings = true;
            preserveConfiguredFontSlots = options.PdfOptions != null;

            if (!string.IsNullOrWhiteSpace(options.FontFamily) &&
                TryApplyPdfFontFamily(options.FontFamily, pdfOptions)) {
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

        private static void ApplyDefaultEmbeddedFontFallback(PdfCore.PdfOptions pdfOptions, ExcelPdfSaveOptions options, bool preserveConfiguredFontSlots) {
            if (options.PdfOptions == null &&
                !preserveConfiguredFontSlots &&
                !pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont)) {
                pdfOptions.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true);
            }
        }

        private static bool TryApplyPdfFontFamily(string? familyName, PdfCore.PdfOptions pdfOptions, bool requireEmbeddedFont = false) {
            return pdfOptions.TryUseOfficeFontFamily(familyName, embedSystemFont: true, requireEmbeddedFont: requireEmbeddedFont);
        }

        private static void RegisterWorksheetFonts(PdfCore.PdfOptions pdfOptions, IReadOnlyList<WorksheetPdfExportPlan> exportPlans, ExcelPdfSaveOptions options, bool preserveConfiguredFontSlots) {
            if (!options.UseWorksheetCellStyles) {
                return;
            }

            var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots = CreateRegisteredFontSlots(pdfOptions, preserveConfiguredFontSlots);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                ExcelCellStyleSnapshot?[,]? styles = plan.ExportData.Styles;
                if (styles == null) {
                    continue;
                }

                int rows = Math.Min(plan.ExportedRows, styles.GetLength(0));
                int columns = styles.GetLength(1);
                for (int row = 0; row < rows; row++) {
                    for (int column = 0; column < columns; column++) {
                        RegisterWorksheetFontCandidate(styles[row, column]?.FontName, pdfOptions, registeredFamilies, registeredFontSlots);
                    }
                }
            }
        }

        private static void RegisterWorksheetFontCandidate(string? familyName, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies, HashSet<PdfCore.PdfStandardFont> registeredFontSlots) {
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

        private static HashSet<PdfCore.PdfStandardFont> CreateRegisteredFontSlots(PdfCore.PdfOptions pdfOptions, bool preserveConfiguredFontSlots) {
            var registeredFontSlots = new HashSet<PdfCore.PdfStandardFont>();
            if (preserveConfiguredFontSlots) {
                AddRegisteredFontSlot(registeredFontSlots, pdfOptions.DefaultFont);
                AddRegisteredFontSlot(registeredFontSlots, pdfOptions.HeaderFont);
                AddRegisteredFontSlot(registeredFontSlots, pdfOptions.FooterFont);
            }

            foreach (PdfCore.PdfStandardFont embeddedFont in pdfOptions.EmbeddedFonts.Keys) {
                AddRegisteredFontSlot(registeredFontSlots, embeddedFont);
            }

            return registeredFontSlots;
        }

        private static void AddRegisteredFontSlot(HashSet<PdfCore.PdfStandardFont> registeredFontSlots, PdfCore.PdfStandardFont font) {
            registeredFontSlots.Add(PdfCore.PdfStandardFontMapper.GetFontFamily(font));
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
                ExcelSheetPageSetup? pageSetup = options.UseWorksheetPageSetup ? workbookSheet?.GetPageSetup() : null;
                ExcelSheet.HeaderFooterSnapshot? headerFooter = (options.UseWorksheetHeadersAndFooters || options.UseWorksheetHeaderFooterImages) ? workbookSheet?.GetHeaderFooter() : null;
                string exportRange = GetExportRange(sheet, workbookSheet, options);
                SheetExportData exportData = ReadSheetExportData(sheet, workbookSheet, exportRange, options, defaultFontFamily);
                IReadOnlyList<int> manualRowBreaks = options.UseWorksheetPageBreaks && workbookSheet != null
                    ? workbookSheet.GetManualRowPageBreaks()
                    : Array.Empty<int>();
                IReadOnlyList<int> manualColumnBreaks = options.UseWorksheetPageBreaks && workbookSheet != null
                    ? workbookSheet.GetManualColumnPageBreaks()
                    : Array.Empty<int>();
                object?[,] values = exportData.Values;
                int rows = values.GetLength(0);
                int columns = values.GetLength(1);
                bool hasTable = rows > 0 && columns > 0;
                int exportedRows = options.MaxRowsPerSheet.HasValue
                    ? Math.Min(rows, options.MaxRowsPerSheet.Value)
                    : rows;
                if (options.MaxRowsPerSheet.HasValue && rows > exportedRows) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetRows",
                        $"Worksheet export was truncated from {rows.ToString(CultureInfo.InvariantCulture)} to {exportedRows.ToString(CultureInfo.InvariantCulture)} rows because MaxRowsPerSheet is set.");
                }

                ISet<string>? exportedCellReferences = CreateExportedCellReferenceSet(exportData.CellReferences, exportedRows);
                bool filterMediaToExportedCells = HasWorksheetPrintArea(workbookSheet, options) ||
                                                  options.MaxRowsPerSheet.HasValue ||
                                                  (options.RespectWorksheetHiddenRowsAndColumns && HasHiddenRowsOrColumns(workbookSheet));
                IReadOnlyList<WorksheetImageExportData> images = FilterImagesByExportedCells(ReadWorksheetImages(workbookSheet, options, sheetName), exportedCellReferences, filterMediaToExportedCells);
                IReadOnlyList<WorksheetChartExportData> charts = FilterChartsByExportedCells(ReadWorksheetCharts(workbookSheet, options, sheetName), exportedCellReferences, filterMediaToExportedCells);
                if (!hasTable && images.Count == 0 && charts.Count == 0) {
                    continue;
                }

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
                    CreateSheetBookmarkName(sheetName, plans.Count + 1)));
            }

            return plans;
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
