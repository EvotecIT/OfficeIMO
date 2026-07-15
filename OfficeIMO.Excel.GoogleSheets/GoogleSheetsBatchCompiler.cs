using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsBatchCompiler {
        private const string DefaultTableFooterColorArgb = "FFE8EAED";
        private const int DefaultGoogleSheetsRowCount = 1000;
        private const int DefaultGoogleSheetsColumnCount = 26;

        internal static GoogleSheetsBatch Build(ExcelDocument document, GoogleSheetsSaveOptions options) {
            var plan = GoogleSheetsPlanBuilder.Build(document, options);
            var report = plan.Report;
            var workbookSnapshot = document.CreateInspectionSnapshot(new ExcelReadOptions {
                UseCachedFormulaResult = true,
                TreatDatesUsingNumberFormat = true,
                NumericAsDecimal = false,
            });
            var title = ResolveTitle(document, workbookSnapshot, options);
            var batch = new GoogleSheetsBatch(title, plan, report);
            batch.ChartDataSheetName = BuildUniqueChartDataSheetName(workbookSnapshot.Worksheets.Select(worksheet => worksheet.Name));

            batch.Add(new GoogleSheetsUpdateSpreadsheetPropertiesRequest {
                Locale = options.Spreadsheet.Locale,
                TimeZone = options.Spreadsheet.TimeZone,
                RecalculationInterval = options.Spreadsheet.RecalculationInterval,
            });
            if (options.Identity.WriteDeveloperMetadata) {
                batch.Add(new GoogleSheetsAddDeveloperMetadataRequest { Key = options.Identity.SourceKey, Value = options.Identity.SourceValue });
                batch.Add(new GoogleSheetsAddDeveloperMetadataRequest { Key = options.Identity.SchemaKey, Value = options.Identity.SchemaValue });
            }

            bool styleNoticeAdded = false;
            bool builtInNameNoticeAdded = false;
            bool tableStyleNoticeAdded = false;
            bool multipleFilterNoticeAdded = false;
            bool protectionNoticeAdded = false;
            bool protectionPermissionsNoticeAdded = false;
            bool tableTotalsNoticeAdded = false;
            bool customFilterNoticeAdded = false;
            bool cellValidationNoticeAdded = false;

            foreach (var worksheet in workbookSnapshot.Worksheets) {
                ResolveGridSize(worksheet, workbookSnapshot.NamedRanges, out int rowCount, out int columnCount);
                batch.Add(new GoogleSheetsAddSheetRequest {
                    SheetName = worksheet.Name,
                    SheetIndex = worksheet.Index,
                    Hidden = worksheet.Hidden,
                    RightToLeft = worksheet.RightToLeft,
                    TabColorArgb = worksheet.TabColorArgb,
                    RowCount = rowCount,
                    ColumnCount = columnCount,
                    FrozenRowCount = worksheet.FrozenRowCount,
                    FrozenColumnCount = worksheet.FrozenColumnCount,
                    HideGridlines = !worksheet.ShowGridlines,
                });

                if (worksheet.Protection != null) {
                    if (!protectionNoticeAdded) {
                        report.Add(
                            TranslationSeverity.Info,
                            "SheetProtection",
                            "Worksheet protection is compiled into Google Sheets protected-sheet requests.");
                        protectionNoticeAdded = true;
                    }

                    if (!protectionPermissionsNoticeAdded && HasUnmappedProtectionPermissions(worksheet.Protection)) {
                        report.Add(
                            TranslationSeverity.Info,
                            "SheetProtectionPermissions",
                            "Excel worksheet-protection permission flags are preserved in the OfficeIMO inspection snapshot, but Google Sheets currently applies whole-sheet protection without exact per-operation permission parity.");
                        protectionPermissionsNoticeAdded = true;
                    }

                    batch.Add(new GoogleSheetsAddProtectedRangeRequest {
                        SheetName = worksheet.Name,
                        Description = BuildProtectionDescription(worksheet.Name, worksheet.Protection),
                        WarningOnly = options.Protection.WarningOnly,
                        DomainUsersCanEdit = options.Protection.DomainUsersCanEdit,
                        EditorEmailAddresses = options.Protection.EditorEmailAddresses.ToArray(),
                        UnprotectedA1Ranges = options.Protection.UnprotectedRangesBySheet.TryGetValue(worksheet.Name, out var unprotectedRanges)
                            ? unprotectedRanges.ToArray()
                            : Array.Empty<string>(),
                    });
                }

                foreach (var column in worksheet.Columns) {
                    batch.Add(new GoogleSheetsUpdateDimensionPropertiesRequest {
                        SheetName = worksheet.Name,
                        DimensionKind = GoogleSheetsDimensionKind.Columns,
                        StartIndex = column.StartIndex - 1,
                        EndIndexExclusive = column.EndIndex,
                        PixelSize = column.Width.HasValue ? ConvertExcelColumnWidthToPixels(column.Width.Value) : null,
                        Hidden = column.Hidden,
                        OutlineLevel = column.OutlineLevel,
                    });
                }

                foreach (var row in worksheet.Rows) {
                    batch.Add(new GoogleSheetsUpdateDimensionPropertiesRequest {
                        SheetName = worksheet.Name,
                        DimensionKind = GoogleSheetsDimensionKind.Rows,
                        StartIndex = row.Index - 1,
                        EndIndexExclusive = row.Index,
                        PixelSize = row.Height.HasValue ? ConvertPointsToPixels(row.Height.Value) : null,
                        Hidden = row.Hidden,
                        OutlineLevel = row.OutlineLevel,
                    });
                }

                AppendDimensionGroups(batch, worksheet);

                foreach (var table in worksheet.Tables) {
                    if (!tableStyleNoticeAdded && !string.IsNullOrWhiteSpace(table.StyleName)) {
                        report.Add(
                            TranslationSeverity.Info,
                            "TableStyles",
                            "Excel table structure is now compiled into native Google Sheets tables, but Excel table style names are currently preserved as diagnostics rather than translated into matching Google visual themes.");
                        tableStyleNoticeAdded = true;
                    }

                    if (!tableTotalsNoticeAdded && table.TotalsRowShown) {
                        report.Add(
                            TranslationSeverity.Info,
                            "TableTotals",
                            "Excel table totals rows now compile into native Google table footers while preserving worksheet totals-row cells and formulas. Per-column totals-row function metadata is still preserved primarily as diagnostics for now.");
                        tableTotalsNoticeAdded = true;
                    }

                    batch.Add(new GoogleSheetsAddTableRequest {
                        SheetName = worksheet.Name,
                        TableName = string.IsNullOrWhiteSpace(table.Name) ? $"Table_{worksheet.Name}_{table.StartRow}_{table.StartColumn}" : table.Name,
                        A1Range = table.A1Range,
                        StartRowIndex = table.StartRow - 1,
                        EndRowIndexExclusive = table.EndRow,
                        StartColumnIndex = table.StartColumn - 1,
                        EndColumnIndexExclusive = table.EndColumn,
                        HasHeaderRow = table.HasHeaderRow,
                        TotalsRowShown = table.TotalsRowShown,
                        StyleName = table.StyleName,
                        HeaderColorArgb = ResolveTableHeaderColorArgb(worksheet, table),
                        FirstBandColorArgb = ResolveTableFirstBandColorArgb(worksheet, table),
                        SecondBandColorArgb = ResolveTableSecondBandColorArgb(worksheet, table),
                        FooterColorArgb = ResolveTableFooterColorArgb(worksheet, table),
                        Columns = BuildTableColumns(workbookSnapshot, worksheet, table),
                    });
                }

                var filterRequests = BuildFilterRequests(worksheet, report, ref multipleFilterNoticeAdded, ref customFilterNoticeAdded);
                foreach (var filterRequest in filterRequests) {
                    batch.Add(filterRequest);
                }

                ExcelSheet? sourceSheet = document.Sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, worksheet.Name, StringComparison.OrdinalIgnoreCase));
                IReadOnlyList<NativePivotCompilation> nativePivots = sourceSheet == null
                    ? Array.Empty<NativePivotCompilation>()
                    : CompilePivotTables(sourceSheet, workbookSnapshot, report, options.UnsupportedFeatures.PivotTables);
                var updateCells = new GoogleSheetsUpdateCellsRequest {
                    SheetName = worksheet.Name
                };

                foreach (var cell in worksheet.Cells) {
                    if (nativePivots.Any(pivot => TryRangeContainsCell(pivot.OutputRange, cell.Row, cell.Column))) {
                        continue;
                    }

                    if (!styleNoticeAdded && cell.Style != null) {
                        report.Add(
                            TranslationSeverity.Info,
                            "Styles",
                            "The current Google Sheets compiler now emits basic cell styling, hyperlinks, and row/column dimensions alongside workbook structure, values, formulas, frozen panes, merges, and named ranges.");
                        styleNoticeAdded = true;
                    }

                    var cellValue = BuildCellValue(cell, options.Formulas);
                    updateCells.AddCell(new GoogleSheetsCellData {
                        RowIndex = cell.Row - 1,
                        ColumnIndex = cell.Column - 1,
                        Value = cellValue,
                        NumberFormatHint = GetNumberFormatHint(cell.Value, cell.Style),
                        Style = BuildCellStyle(cell.Style),
                        DataValidationRule = BuildCellValidationRule(workbookSnapshot, worksheet, cell.Row, cell.Column, report, ref cellValidationNoticeAdded),
                        Hyperlink = BuildHyperlink(cell.Hyperlink),
                        Comment = BuildComment(cell.Comment),
                        TextFormatRuns = BuildTextFormatRuns(cell.RichTextRuns),
                    });
                }

                if (updateCells.Cells.Count > 0) {
                    batch.Add(updateCells);
                }

                AppendValidationRanges(workbookSnapshot, worksheet, batch, report, ref cellValidationNoticeAdded);

                foreach (var mergedRange in worksheet.MergedRanges) {
                    batch.Add(new GoogleSheetsMergeCellsRequest {
                        SheetName = worksheet.Name,
                        A1Range = mergedRange.A1Range,
                        StartRowIndex = mergedRange.StartRow - 1,
                        EndRowIndexExclusive = mergedRange.EndRow,
                        StartColumnIndex = mergedRange.StartColumn - 1,
                        EndColumnIndexExclusive = mergedRange.EndColumn,
                    });
                }

                if (sourceSheet != null) {
                    AppendConditionalFormatting(batch, sourceSheet, report);
                    AppendCharts(batch, sourceSheet, worksheet, report, options.UnsupportedFeatures.Charts);
                    foreach (NativePivotCompilation pivot in nativePivots) {
                        batch.Add(pivot.Request);
                    }
                }
            }

            var reservedNamedRangeNames = new HashSet<string>(
                workbookSnapshot.NamedRanges
                    .Where(namedRange => !namedRange.IsBuiltIn && string.IsNullOrWhiteSpace(namedRange.SheetName))
                    .Select(namedRange => namedRange.Name),
                StringComparer.OrdinalIgnoreCase);
            foreach (var namedRange in workbookSnapshot.NamedRanges) {
                if (namedRange.IsBuiltIn) {
                    if (!builtInNameNoticeAdded) {
                        report.Add(
                            TranslationSeverity.Info,
                            "BuiltInNames",
                            "Built-in Excel names such as print areas are detected, but they are kept as diagnostics rather than emitted as Google named ranges in the first compiler slice.");
                        builtInNameNoticeAdded = true;
                    }
                    continue;
                }

                string targetName = string.IsNullOrWhiteSpace(namedRange.SheetName)
                    ? namedRange.Name
                    : BuildSheetScopedNamedRangeName(namedRange, reservedNamedRangeNames, report);
                batch.Add(new GoogleSheetsAddNamedRangeRequest {
                    Name = targetName,
                    SourceName = namedRange.Name,
                    SheetName = namedRange.SheetName,
                    A1Range = namedRange.ReferenceA1,
                });
            }

            return batch;
        }

        private static void ResolveGridSize(
            ExcelWorksheetSnapshot worksheet,
            IReadOnlyList<ExcelNamedRangeSnapshot> namedRanges,
            out int rowCount,
            out int columnCount) {
            rowCount = DefaultGoogleSheetsRowCount;
            columnCount = DefaultGoogleSheetsColumnCount;
            string usedRange = worksheet.UsedRangeA1.Replace("$", string.Empty);
            if (A1.TryParseRange(usedRange, out _, out _, out int lastRow, out int lastColumn)) {
                rowCount = Math.Max(rowCount, lastRow);
                columnCount = Math.Max(columnCount, lastColumn);
            }

            foreach (ExcelDataValidationSnapshot validation in worksheet.Validations) {
                foreach (string a1Range in validation.A1Ranges) {
                    ExpandGridToInclude(a1Range, ref rowCount, ref columnCount);
                }
            }

            foreach (ExcelMergedRangeSnapshot mergedRange in worksheet.MergedRanges) {
                ExpandGridToInclude(mergedRange.A1Range, ref rowCount, ref columnCount);
            }

            foreach (ExcelTableSnapshot table in worksheet.Tables) {
                ExpandGridToInclude(table.A1Range, ref rowCount, ref columnCount);
            }

            foreach (ExcelNamedRangeSnapshot namedRange in namedRanges.Where(range => !range.IsBuiltIn)) {
                string? targetSheetName = namedRange.SheetName;
                string rangeText = namedRange.ReferenceA1.TrimStart('=').Replace("$", string.Empty);
                if (TrySplitSheetQualifiedRange(rangeText, out string? explicitSheetName, out string unqualifiedRange)) {
                    targetSheetName = explicitSheetName;
                    rangeText = unqualifiedRange;
                }

                if (string.Equals(targetSheetName, worksheet.Name, StringComparison.OrdinalIgnoreCase)) {
                    ExpandGridToIncludeNamedRange(rangeText, ref rowCount, ref columnCount);
                }
            }

            rowCount = Math.Max(rowCount, worksheet.FrozenRowCount);
            columnCount = Math.Max(columnCount, worksheet.FrozenColumnCount);
        }

        private static void ExpandGridToInclude(string a1Range, ref int rowCount, ref int columnCount) {
            string normalizedRange = a1Range.Replace("$", string.Empty);
            if (!A1.TryParseRange(normalizedRange, out _, out _, out int lastRow, out int lastColumn)) {
                return;
            }

            rowCount = Math.Max(rowCount, lastRow);
            columnCount = Math.Max(columnCount, lastColumn);
        }

        private static void ExpandGridToIncludeNamedRange(string a1Range, ref int rowCount, ref int columnCount) {
            string normalizedRange = a1Range.Replace("$", string.Empty);
            if (!A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return;
            }

            if (firstRow != 1 || lastRow != A1.MaxRows) {
                rowCount = Math.Max(rowCount, lastRow);
            }
            if (firstColumn != 1 || lastColumn != A1.MaxColumns) {
                columnCount = Math.Max(columnCount, lastColumn);
            }
        }

        private static string ResolveTitle(ExcelDocument document, ExcelWorkbookSnapshot workbookSnapshot, GoogleSheetsSaveOptions options) {
            if (!string.IsNullOrWhiteSpace(options.Title)) {
                return options.Title!;
            }

            if (!string.IsNullOrWhiteSpace(workbookSnapshot.Title)) {
                return workbookSnapshot.Title!;
            }

            if (!string.IsNullOrWhiteSpace(workbookSnapshot.FilePath)) {
                return Path.GetFileNameWithoutExtension(workbookSnapshot.FilePath);
            }

            if (!string.IsNullOrWhiteSpace(document.FilePath)) {
                return Path.GetFileNameWithoutExtension(document.FilePath);
            }

            return "Workbook";
        }

        private static bool HasUnmappedProtectionPermissions(ExcelWorksheetProtectionSnapshot protection) {
            return protection.AllowSelectLockedCells
                || protection.AllowSelectUnlockedCells
                || protection.AllowFormatCells
                || protection.AllowFormatColumns
                || protection.AllowFormatRows
                || protection.AllowInsertColumns
                || protection.AllowInsertRows
                || protection.AllowInsertHyperlinks
                || protection.AllowDeleteColumns
                || protection.AllowDeleteRows
                || protection.AllowSort
                || protection.AllowAutoFilter
                || protection.AllowPivotTables;
        }

        private static string BuildProtectionDescription(string sheetName, ExcelWorksheetProtectionSnapshot protection) {
            var allowedOperations = new List<string>();

            if (protection.AllowSelectLockedCells) allowedOperations.Add("select locked cells");
            if (protection.AllowSelectUnlockedCells) allowedOperations.Add("select unlocked cells");
            if (protection.AllowFormatCells) allowedOperations.Add("format cells");
            if (protection.AllowFormatColumns) allowedOperations.Add("format columns");
            if (protection.AllowFormatRows) allowedOperations.Add("format rows");
            if (protection.AllowInsertColumns) allowedOperations.Add("insert columns");
            if (protection.AllowInsertRows) allowedOperations.Add("insert rows");
            if (protection.AllowInsertHyperlinks) allowedOperations.Add("insert hyperlinks");
            if (protection.AllowDeleteColumns) allowedOperations.Add("delete columns");
            if (protection.AllowDeleteRows) allowedOperations.Add("delete rows");
            if (protection.AllowSort) allowedOperations.Add("sort");
            if (protection.AllowAutoFilter) allowedOperations.Add("use autofilter");
            if (protection.AllowPivotTables) allowedOperations.Add("use pivot tables");

            if (allowedOperations.Count == 0) {
                return $"OfficeIMO worksheet protection for '{sheetName}'.";
            }

            return $"OfficeIMO worksheet protection for '{sheetName}' allows: {string.Join(", ", allowedOperations)}.";
        }
    }
}
