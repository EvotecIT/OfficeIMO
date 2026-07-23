using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Imports Google spreadsheets through either loss-minimizing Drive conversion or the native Sheets model.
    /// </summary>
    public sealed class GoogleSheetsImporter : IGoogleSheetsImporter {
        private const string NativeFields =
            "spreadsheetId,spreadsheetUrl,properties(title,locale,timeZone)," +
            "namedRanges(name,range)," +
            "sheets(properties(sheetId,title,index,hidden,rightToLeft,tabColor,gridProperties)," +
            "data(startRow,startColumn,rowData(values(userEnteredValue,effectiveValue,formattedValue,userEnteredFormat,textFormatRuns,note,dataValidation,pivotTable)),rowMetadata(hiddenByUser,pixelSize),columnMetadata(hiddenByUser,pixelSize))," +
            "merges,conditionalFormats,charts,tables,filterViews,basicFilter,rowGroups,columnGroups)";

        public async Task<GoogleSheetsImportResult> ImportAsync(
            string spreadsheetId,
            GoogleWorkspaceSession session,
            GoogleSheetsImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(spreadsheetId)) throw new ArgumentException("Spreadsheet ID is required.", nameof(spreadsheetId));
            if (session == null) throw new ArgumentNullException(nameof(session));

            GoogleSheetsImportOptions effectiveOptions = options ?? new GoogleSheetsImportOptions();
            ValidateOptions(effectiveOptions);
            return effectiveOptions.Mode == GoogleSheetsImportMode.DriveExport
                ? await ImportViaDriveAsync(spreadsheetId, session, effectiveOptions, cancellationToken).ConfigureAwait(false)
                : await ImportNativeAsync(spreadsheetId, session, effectiveOptions, cancellationToken).ConfigureAwait(false);
        }

        private static async Task<GoogleSheetsImportResult> ImportViaDriveAsync(
            string spreadsheetId,
            GoogleWorkspaceSession session,
            GoogleSheetsImportOptions options,
            CancellationToken cancellationToken) {
            var report = new TranslationReport();
            using var drive = new GoogleDriveClient(session);
            GoogleDriveFile source = await drive.GetFileAsync(spreadsheetId, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
            EnsureSpreadsheet(source, spreadsheetId);
            EnsureDownloadable(source, spreadsheetId);

            byte[] xlsx = await drive.ExportAsync(
                spreadsheetId,
                GoogleDriveMimeTypes.MicrosoftExcel,
                options.Progress,
                report,
                cancellationToken,
                options.MaxResponseBytes).ConfigureAwait(false);

            var stream = new MemoryStream(xlsx, writable: true);
            ExcelDocument document;
            try {
                document = ExcelDocument.Load(stream, options.LoadOptions);
            } catch {
                stream.Dispose();
                throw;
            }

            report.Add(
                TranslationSeverity.Info,
                "DriveExportImport",
                "The Google spreadsheet was exported to XLSX through Drive and loaded by OfficeIMO.",
                code: "SHEETS.IMPORT.DRIVE_EXPORT",
                action: TranslationAction.Preserve,
                targetId: spreadsheetId);
            if (options.Ranges.Count > 0) {
                report.Add(
                    TranslationSeverity.Warning,
                    "PartialImport",
                    "Drive-export import always returns the complete workbook; range selection applies only to native import.",
                    code: "SHEETS.IMPORT.RANGES_IGNORED",
                    action: TranslationAction.Preserve,
                    count: options.Ranges.Count);
            }

            return new GoogleSheetsImportResult(document, BuildReference(source, spreadsheetId, report), report);
        }

        private static async Task<GoogleSheetsImportResult> ImportNativeAsync(
            string spreadsheetId,
            GoogleWorkspaceSession session,
            GoogleSheetsImportOptions options,
            CancellationToken cancellationToken) {
            var report = new TranslationReport();
            using var drive = new GoogleDriveClient(session);
            GoogleDriveFile source = await drive.GetFileAsync(spreadsheetId, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
            EnsureSpreadsheet(source, spreadsheetId);
            EnsureDownloadable(source, spreadsheetId);

            GoogleWorkspaceAccessToken token = await session
                .AcquireAccessTokenAsync(new[] { GoogleWorkspaceScopeCatalog.SpreadsheetsReadonly }, cancellationToken)
                .ConfigureAwait(false);
            string uri = BuildNativeReadUri(spreadsheetId, options);
            GoogleSheetsNativeSpreadsheet spreadsheet;
            using (var transport = new GoogleWorkspaceHttpTransport(session.Options)) {
                spreadsheet = await transport.SendJsonAsync<GoogleSheetsNativeSpreadsheet>(
                    token.AccessToken,
                    HttpMethod.Get,
                    uri,
                    null,
                    GoogleWorkspaceRequestSafety.Safe,
                    "Google Sheets API",
                    report,
                    GoogleSheetsJsonSerializerContext.Default.GoogleSheetsNativeSpreadsheet,
                    cancellationToken,
                    options.MaxResponseBytes).ConfigureAwait(false);
            }

            ValidateNativeSpreadsheet(spreadsheet, options);
            ExcelDocument document = ProjectNativeSpreadsheet(spreadsheet, report);
            report.Add(
                TranslationSeverity.Info,
                "NativeImport",
                "Native Sheets values, formulas, core styles, sheet settings, merges, notes, and named ranges were projected into OfficeIMO.",
                code: "SHEETS.IMPORT.NATIVE",
                action: TranslationAction.Preserve,
                targetId: spreadsheetId);
            return new GoogleSheetsImportResult(document, BuildReference(source, spreadsheetId, report, spreadsheet), report);
        }

        private static ExcelDocument ProjectNativeSpreadsheet(
            GoogleSheetsNativeSpreadsheet spreadsheet,
            TranslationReport report) {
            var stream = new MemoryStream();
            ExcelDocument document = ExcelDocument.Create(stream);
            try {
                var sheetNames = new Dictionary<int, string>();
                foreach (GoogleSheetsNativeSheet nativeSheet in spreadsheet.Sheets.OrderBy(sheet => sheet.Properties.Index)) {
                    string title = string.IsNullOrWhiteSpace(nativeSheet.Properties.Title)
                        ? $"Sheet{nativeSheet.Properties.Index + 1}"
                        : nativeSheet.Properties.Title;
                    ExcelSheet sheet = document.AddWorksheet(title);
                    sheetNames[nativeSheet.Properties.SheetId] = sheet.Name;
                    ApplySheetProperties(sheet, nativeSheet.Properties);
                    ApplyGridData(sheet, nativeSheet.Data, report);
                    ApplyDimensionGroups(sheet, nativeSheet.Properties, nativeSheet.RowGroups, "ROWS", report);
                    ApplyDimensionGroups(sheet, nativeSheet.Properties, nativeSheet.ColumnGroups, "COLUMNS", report);
                    ApplyMerges(sheet, nativeSheet.Merges);
                    AddNativeObjectDiagnostics(nativeSheet, report);
                }

                foreach (GoogleSheetsNativeNamedRange namedRange in spreadsheet.NamedRanges) {
                    if (string.IsNullOrWhiteSpace(namedRange.Name)
                        || !sheetNames.TryGetValue(namedRange.Range.SheetId, out string? sheetName)
                        || !TryBuildA1Range(namedRange.Range, out string? a1Range)) {
                        report.Add(
                            TranslationSeverity.Warning,
                            "NamedRanges",
                            $"Native named range '{namedRange.Name}' could not be represented as an Excel range.",
                            code: "SHEETS.IMPORT.NAMED_RANGE_SKIPPED",
                            action: TranslationAction.Skip);
                        continue;
                    }

                    document.SetNamedRange(namedRange.Name, $"'{EscapeSheetName(sheetName)}'!{a1Range}", save: false);
                }

                document.Save();
                return document;
            } catch {
                document.Dispose();
                stream.Dispose();
                throw;
            }
        }

        private static void ApplySheetProperties(ExcelSheet sheet, GoogleSheetsNativeSheetProperties properties) {
            if (properties.Hidden) sheet.SetHidden(true);
            if (properties.RightToLeft) sheet.SetRightToLeft(true);
            if (properties.TabColor != null) sheet.SetTabColor(ToHex(properties.TabColor));
            if (properties.GridProperties != null) {
                if (properties.GridProperties.FrozenRowCount > 0 || properties.GridProperties.FrozenColumnCount > 0) {
                    sheet.Freeze(properties.GridProperties.FrozenRowCount, properties.GridProperties.FrozenColumnCount);
                }
                sheet.SetGridlinesVisible(!properties.GridProperties.HideGridlines);
            }
        }

        private static void ApplyGridData(
            ExcelSheet sheet,
            IReadOnlyList<GoogleSheetsNativeGridData> data,
            TranslationReport report) {
            foreach (GoogleSheetsNativeGridData block in data) {
                ApplyDimensionMetadata(sheet, block);
                for (int rowOffset = 0; rowOffset < block.RowData.Count; rowOffset++) {
                    GoogleSheetsNativeRowData row = block.RowData[rowOffset];
                    for (int columnOffset = 0; columnOffset < row.Values.Count; columnOffset++) {
                        int rowIndex = block.StartRow + rowOffset + 1;
                        int columnIndex = block.StartColumn + columnOffset + 1;
                        ApplyCell(sheet, rowIndex, columnIndex, row.Values[columnOffset], report);
                    }
                }
            }
        }

        private static void ApplyDimensionMetadata(ExcelSheet sheet, GoogleSheetsNativeGridData block) {
            for (int offset = 0; offset < block.RowMetadata.Count; offset++) {
                int rowIndex = block.StartRow + offset + 1;
                if (rowIndex > 1048576) break;
                GoogleSheetsNativeDimensionProperties metadata = block.RowMetadata[offset];
                if (metadata.PixelSize is int pixels && pixels > 0) {
                    sheet.SetRowHeight(rowIndex, pixels * 72d / 96d);
                }
                if (metadata.HiddenByUser == true) {
                    sheet.SetRowHidden(rowIndex, true);
                }
            }

            for (int offset = 0; offset < block.ColumnMetadata.Count; offset++) {
                int columnIndex = block.StartColumn + offset + 1;
                if (columnIndex > 16384) break;
                GoogleSheetsNativeDimensionProperties metadata = block.ColumnMetadata[offset];
                if (metadata.PixelSize is int pixels && pixels > 0) {
                    sheet.SetColumnWidth(columnIndex, ConvertPixelsToExcelColumnWidth(pixels));
                }
                if (metadata.HiddenByUser == true) {
                    sheet.SetColumnHidden(columnIndex, true);
                }
            }
        }

        private static void ApplyDimensionGroups(
            ExcelSheet sheet,
            GoogleSheetsNativeSheetProperties properties,
            IReadOnlyList<GoogleSheetsNativeDimensionGroup> groups,
            string dimension,
            TranslationReport report) {
            int maximum = string.Equals(dimension, "ROWS", StringComparison.Ordinal) ? 1048576 : 16384;
            var levels = new Dictionary<int, byte>();
            foreach (GoogleSheetsNativeDimensionGroup group in groups) {
                GoogleSheetsNativeDimensionRange range = group.Range;
                if (range.SheetId != properties.SheetId
                    || !string.Equals(range.Dimension, dimension, StringComparison.OrdinalIgnoreCase)
                    || range.StartIndex < 0
                    || range.EndIndex <= range.StartIndex
                    || group.Depth <= 0) {
                    report.Add(
                        TranslationSeverity.Warning,
                        "DimensionGroups",
                        $"A native {dimension.ToLowerInvariant()} group on '{properties.Title}' had an invalid range or depth and was skipped.",
                        path: properties.Title,
                        code: "SHEETS.IMPORT.DIMENSION_GROUP_SKIPPED",
                        action: TranslationAction.Skip);
                    continue;
                }

                byte level = (byte)Math.Min(7, group.Depth);
                if (group.Depth > 7) {
                    report.AddUnique(
                        TranslationSeverity.Warning,
                        "DimensionGroups",
                        $"Google Sheets outline depth exceeds Excel's maximum of 7 on '{properties.Title}'; deeper groups were flattened to level 7.",
                        path: properties.Title,
                        code: "SHEETS.IMPORT.DIMENSION_GROUP_DEPTH_FLATTENED",
                        action: TranslationAction.Flatten);
                }
                int end = Math.Min(maximum, range.EndIndex);
                for (int index = range.StartIndex; index < end; index++) {
                    int oneBasedIndex = index + 1;
                    if (!levels.TryGetValue(oneBasedIndex, out byte existing) || level > existing) {
                        levels[oneBasedIndex] = level;
                    }
                }
            }

            foreach (KeyValuePair<int, byte> item in levels.OrderBy(item => item.Key)) {
                if (string.Equals(dimension, "ROWS", StringComparison.Ordinal)) {
                    sheet.SetRowOutline(item.Key, item.Value, save: false);
                } else {
                    sheet.SetColumnOutline(item.Key, item.Value, save: false);
                }
            }
        }

        private static double ConvertPixelsToExcelColumnWidth(int pixels) {
            const double maximumDigitWidth = 7d;
            double roundingOffset = Math.Truncate(128d / maximumDigitWidth) / 256d;
            return Math.Max(0d, (pixels / maximumDigitWidth) - roundingOffset);
        }

        private static void ApplyCell(
            ExcelSheet sheet,
            int row,
            int column,
            GoogleSheetsNativeCellData cell,
            TranslationReport report) {
            GoogleSheetsNativeExtendedValue? entered = cell.UserEnteredValue;
            GoogleSheetsNativeExtendedValue? effective = cell.EffectiveValue;
            if (!string.IsNullOrWhiteSpace(entered?.FormulaValue)) {
                string formula = entered!.FormulaValue!;
                sheet.CellFormula(row, column, formula.StartsWith("=", StringComparison.Ordinal) ? formula.Substring(1) : formula);
            } else if (entered?.StringValue != null) {
                sheet.CellValue(row, column, entered.StringValue);
            } else if (entered?.NumberValue is double number) {
                sheet.CellValue(row, column, number);
            } else if (entered?.BoolValue is bool boolean) {
                sheet.CellValue(row, column, boolean);
            } else if (effective?.StringValue != null) {
                sheet.CellValue(row, column, effective.StringValue);
            } else if (effective?.NumberValue is double effectiveNumber) {
                sheet.CellValue(row, column, effectiveNumber);
            } else if (effective?.BoolValue is bool effectiveBoolean) {
                sheet.CellValue(row, column, effectiveBoolean);
            } else if (effective?.ErrorValue != null) {
                sheet.CellValue(row, column, cell.FormattedValue ?? effective.ErrorValue.Type ?? "#ERROR!");
                report.AddUnique(
                    TranslationSeverity.Warning,
                    "FormulaErrors",
                    "Native formula errors are imported as their displayed value.",
                    path: sheet.Name,
                    code: "SHEETS.IMPORT.FORMULA_ERROR",
                    action: TranslationAction.Flatten);
            } else if (!string.IsNullOrEmpty(cell.FormattedValue)) {
                sheet.CellValue(row, column, cell.FormattedValue!);
            }

            ApplyCellFormat(sheet, row, column, cell.UserEnteredFormat);
            ApplyRichText(sheet, row, column, cell, report);
            if (!string.IsNullOrWhiteSpace(cell.Note)) {
                sheet.SetComment(row, column, cell.Note!, "Google Sheets");
            }

            if (cell.DataValidation != null) {
                report.AddUnique(
                    TranslationSeverity.Warning,
                    "DataValidation",
                    "Native data-validation metadata is detected; use Drive-export import when exact validation preservation is required.",
                    path: sheet.Name,
                    code: "SHEETS.IMPORT.DATA_VALIDATION_FALLBACK",
                    action: TranslationAction.Flatten);
            }
            if (cell.PivotTable != null) {
                report.AddUnique(
                    TranslationSeverity.Warning,
                    "PivotTables",
                    "Native pivot tables are detected; use Drive-export import for pivot preservation.",
                    path: sheet.Name,
                    code: "SHEETS.IMPORT.PIVOT_FALLBACK",
                    action: TranslationAction.Flatten);
            }
        }

        private static void ApplyCellFormat(
            ExcelSheet sheet,
            int row,
            int column,
            GoogleSheetsNativeCellFormat? format) {
            if (format == null) return;
            if (!string.IsNullOrWhiteSpace(format.NumberFormat?.Pattern)) sheet.FormatCell(row, column, format.NumberFormat!.Pattern!);
            if (format.BackgroundColor != null) sheet.CellBackground(row, column, ToHex(format.BackgroundColor));
            if (format.TextFormat != null) {
                GoogleSheetsNativeTextFormat text = format.TextFormat;
                if (text.Bold == true) sheet.CellBold(row, column);
                if (text.Italic == true) sheet.CellItalic(row, column);
                if (text.Underline == true) sheet.CellUnderline(row, column);
                if (text.Strikethrough == true) sheet.CellStrikethrough(row, column);
                if (!string.IsNullOrWhiteSpace(text.FontFamily)) sheet.CellFontName(row, column, text.FontFamily!);
                if (text.FontSize is int fontSize && fontSize > 0) sheet.CellFontSize(row, column, fontSize);
                if (text.ForegroundColor != null) sheet.CellFontColor(row, column, ToHex(text.ForegroundColor));
            }

            switch (format.HorizontalAlignment) {
                case "LEFT": sheet.CellAlign(row, column, HorizontalAlignmentValues.Left); break;
                case "CENTER": sheet.CellAlign(row, column, HorizontalAlignmentValues.Center); break;
                case "RIGHT": sheet.CellAlign(row, column, HorizontalAlignmentValues.Right); break;
                case "JUSTIFY": sheet.CellAlign(row, column, HorizontalAlignmentValues.Justify); break;
            }
            switch (format.VerticalAlignment) {
                case "TOP": sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Top); break;
                case "MIDDLE": sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center); break;
                case "BOTTOM": sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Bottom); break;
            }
            if (string.Equals(format.WrapStrategy, "WRAP", StringComparison.Ordinal)) sheet.CellWrapText(row, column);
            if (format.TextRotation?.Vertical == true) sheet.CellTextRotation(row, column, 255);
            else if (format.TextRotation?.Angle is int angle) sheet.CellTextRotation(row, column, angle >= 0 ? angle : 90 - angle);
        }

        private static void ApplyRichText(
            ExcelSheet sheet,
            int row,
            int column,
            GoogleSheetsNativeCellData cell,
            TranslationReport report) {
            if (cell.TextFormatRuns.Count == 0) return;
            string? text = cell.UserEnteredValue?.StringValue;
            if (text == null || !HasValidTextFormatRunIndexes(cell.TextFormatRuns, text.Length)) {
                report.AddUnique(
                    TranslationSeverity.Warning,
                    "RichText",
                    "Native rich-text runs could not be projected because their source value or character indexes were invalid.",
                    path: sheet.Name,
                    code: "SHEETS.IMPORT.RICH_TEXT_FALLBACK",
                    action: TranslationAction.Flatten);
                return;
            }

            GoogleSheetsNativeTextFormat? baseFormat = cell.UserEnteredFormat?.TextFormat;
            var projected = new List<ExcelRichTextRun>();
            int cursor = 0;
            for (int index = 0; index < cell.TextFormatRuns.Count; index++) {
                GoogleSheetsNativeTextFormatRun sourceRun = cell.TextFormatRuns[index];
                if (sourceRun.StartIndex > cursor) {
                    projected.Add(CreateRichTextRun(text.Substring(cursor, sourceRun.StartIndex - cursor), baseFormat, null));
                }

                int endIndex = index + 1 < cell.TextFormatRuns.Count
                    ? cell.TextFormatRuns[index + 1].StartIndex
                    : text.Length;
                if (endIndex > sourceRun.StartIndex) {
                    projected.Add(CreateRichTextRun(text.Substring(sourceRun.StartIndex, endIndex - sourceRun.StartIndex), baseFormat, sourceRun.Format));
                }
                cursor = endIndex;
            }

            if (projected.Count > 0) sheet.SetRichText(row, column, projected);
        }

        private static bool HasValidTextFormatRunIndexes(IReadOnlyList<GoogleSheetsNativeTextFormatRun> runs, int textLength) {
            int previous = -1;
            foreach (GoogleSheetsNativeTextFormatRun run in runs) {
                if (run.StartIndex < 0 || run.StartIndex > textLength || run.StartIndex <= previous) return false;
                previous = run.StartIndex;
            }
            return true;
        }

        private static ExcelRichTextRun CreateRichTextRun(
            string text,
            GoogleSheetsNativeTextFormat? baseFormat,
            GoogleSheetsNativeTextFormat? runFormat) {
            GoogleSheetsNativeColor? color = runFormat?.ForegroundColor ?? baseFormat?.ForegroundColor;
            return new ExcelRichTextRun(text) {
                Bold = runFormat?.Bold ?? baseFormat?.Bold ?? false,
                Italic = runFormat?.Italic ?? baseFormat?.Italic ?? false,
                Underline = runFormat?.Underline ?? baseFormat?.Underline ?? false,
                Strikethrough = runFormat?.Strikethrough ?? baseFormat?.Strikethrough ?? false,
                FontName = runFormat?.FontFamily ?? baseFormat?.FontFamily,
                FontSize = runFormat?.FontSize ?? baseFormat?.FontSize,
                FontColor = color == null ? null : ToHex(color),
            };
        }

        private static void ApplyMerges(ExcelSheet sheet, IReadOnlyList<GoogleSheetsNativeGridRange> ranges) {
            foreach (GoogleSheetsNativeGridRange range in ranges) {
                if (TryBuildA1Range(range, out string? a1Range)) {
                    sheet.MergeRange(a1Range!);
                }
            }
        }

        private static void AddNativeObjectDiagnostics(GoogleSheetsNativeSheet sheet, TranslationReport report) {
            AddFallbackNotice(report, sheet.Properties.Title, "ConditionalFormatting", "conditional-format rules", sheet.ConditionalFormats.Count);
            AddFallbackNotice(report, sheet.Properties.Title, "Charts", "charts", sheet.Charts.Count);
            AddFallbackNotice(report, sheet.Properties.Title, "Tables", "tables", sheet.Tables.Count);
            AddFallbackNotice(report, sheet.Properties.Title, "FilterViews", "filter views", sheet.FilterViews.Count);
            AddFallbackNotice(report, sheet.Properties.Title, "BasicFilters", "basic filters", sheet.BasicFilter == null ? 0 : 1);
        }

        private static void AddFallbackNotice(TranslationReport report, string sheet, string feature, string description, int count) {
            if (count == 0) return;
            report.Add(
                TranslationSeverity.Warning,
                feature,
                $"Native import detected {count} {description} on '{sheet}'. Use Drive-export import for broad object preservation.",
                path: sheet,
                code: "SHEETS.IMPORT." + feature.ToUpperInvariant() + "_FALLBACK",
                action: TranslationAction.Flatten,
                count: count);
        }

        private static string BuildNativeReadUri(string spreadsheetId, GoogleSheetsImportOptions options) {
            var uri = new StringBuilder("https://sheets.googleapis.com/v4/spreadsheets/")
                .Append(Uri.EscapeDataString(spreadsheetId))
                .Append("?fields=")
                .Append(Uri.EscapeDataString(string.IsNullOrWhiteSpace(options.Fields) ? NativeFields : options.Fields!));
            foreach (string range in options.Ranges.Where(range => !string.IsNullOrWhiteSpace(range))) {
                uri.Append("&ranges=").Append(Uri.EscapeDataString(range));
            }
            return uri.ToString();
        }

        private static void ValidateOptions(GoogleSheetsImportOptions options) {
            if (options.MaxResponseBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxResponseBytes));
            if (options.MaxSheets <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSheets));
            if (options.MaxCells <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxCells));
            if (options.MaxDimensionGroupMembers <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxDimensionGroupMembers));
        }

        private static void ValidateNativeSpreadsheet(
            GoogleSheetsNativeSpreadsheet spreadsheet,
            GoogleSheetsImportOptions options) {
            if (spreadsheet.Sheets.Count > options.MaxSheets) {
                throw new InvalidDataException($"Google Sheets native import exceeded the configured {options.MaxSheets} sheet limit.");
            }

            long cells = 0;
            long dimensionGroupMembers = 0;
            foreach (GoogleSheetsNativeSheet sheet in spreadsheet.Sheets) {
                foreach (GoogleSheetsNativeGridData block in sheet.Data) {
                    foreach (GoogleSheetsNativeRowData row in block.RowData) {
                        if (row.Values.Count > options.MaxCells - cells) {
                            throw new InvalidDataException($"Google Sheets native import exceeded the configured {options.MaxCells} cell limit.");
                        }
                        cells += row.Values.Count;
                    }
                }
                CountDimensionGroupMembers(sheet.RowGroups, sheet.Properties.SheetId, "ROWS", 1048576,
                    options, ref dimensionGroupMembers);
                CountDimensionGroupMembers(sheet.ColumnGroups, sheet.Properties.SheetId, "COLUMNS", 16384,
                    options, ref dimensionGroupMembers);
            }
        }

        private static void CountDimensionGroupMembers(
            IReadOnlyList<GoogleSheetsNativeDimensionGroup> groups,
            int sheetId,
            string dimension,
            int maximum,
            GoogleSheetsImportOptions options,
            ref long total) {
            foreach (GoogleSheetsNativeDimensionGroup group in groups) {
                GoogleSheetsNativeDimensionRange range = group.Range;
                if (range.SheetId != sheetId
                    || !string.Equals(range.Dimension, dimension, StringComparison.OrdinalIgnoreCase)
                    || range.StartIndex < 0
                    || range.StartIndex >= maximum
                    || range.EndIndex <= range.StartIndex
                    || group.Depth <= 0) {
                    continue;
                }

                long members = Math.Min(maximum, range.EndIndex) - (long)range.StartIndex;
                if (members > options.MaxDimensionGroupMembers - total) {
                    throw new InvalidDataException(
                        $"Google Sheets native import exceeded the configured {options.MaxDimensionGroupMembers} dimension-group member limit.");
                }
                total += members;
            }
        }

        private static void EnsureSpreadsheet(GoogleDriveFile source, string spreadsheetId) {
            if (!string.Equals(source.MimeType, GoogleDriveMimeTypes.Spreadsheet, StringComparison.Ordinal)) {
                throw new InvalidOperationException($"Drive file '{spreadsheetId}' is not a Google spreadsheet (mimeType: '{source.MimeType}').");
            }
        }

        private static void EnsureDownloadable(GoogleDriveFile source, string spreadsheetId) {
            if (source.Capabilities != null && !source.Capabilities.CanDownload) {
                throw new InvalidOperationException($"Drive file '{spreadsheetId}' cannot be downloaded or exported by the current principal.");
            }
        }

        private static GoogleSpreadsheetReference BuildReference(
            GoogleDriveFile source,
            string spreadsheetId,
            TranslationReport report,
            GoogleSheetsNativeSpreadsheet? native = null) {
            return new GoogleSpreadsheetReference {
                SpreadsheetId = native?.SpreadsheetId ?? source.Id ?? spreadsheetId,
                FileId = source.Id ?? spreadsheetId,
                Name = native?.Properties?.Title ?? source.Name,
                MimeType = source.MimeType,
                WebViewLink = native?.SpreadsheetUrl ?? source.WebViewLink,
                DriveVersion = source.Version,
                ModifiedTime = source.ModifiedTime,
                Report = report,
            };
        }

        private static bool TryBuildA1Range(GoogleSheetsNativeGridRange range, out string? a1Range) {
            a1Range = null;
            bool hasRows = range.StartRowIndex.HasValue && range.EndRowIndex.HasValue;
            bool hasColumns = range.StartColumnIndex.HasValue && range.EndColumnIndex.HasValue;
            if ((!hasRows && (range.StartRowIndex.HasValue || range.EndRowIndex.HasValue))
                || (!hasColumns && (range.StartColumnIndex.HasValue || range.EndColumnIndex.HasValue))
                || (!hasRows && !hasColumns)
                || (hasRows && range.EndRowIndex <= range.StartRowIndex)
                || (hasColumns && range.EndColumnIndex <= range.StartColumnIndex)) {
                return false;
            }

            if (!hasRows) {
                a1Range = ToColumnName(range.StartColumnIndex!.Value + 1)
                    + "1:"
                    + ToColumnName(range.EndColumnIndex!.Value)
                    + "1048576";
            } else if (!hasColumns) {
                a1Range = "A"
                    + (range.StartRowIndex!.Value + 1).ToString(CultureInfo.InvariantCulture)
                    + ":"
                    + "XFD"
                    + range.EndRowIndex!.Value.ToString(CultureInfo.InvariantCulture);
            } else {
                a1Range = ToColumnName(range.StartColumnIndex!.Value + 1)
                    + (range.StartRowIndex!.Value + 1).ToString(CultureInfo.InvariantCulture)
                    + ":"
                    + ToColumnName(range.EndColumnIndex!.Value)
                    + range.EndRowIndex!.Value.ToString(CultureInfo.InvariantCulture);
            }
            return true;
        }

        private static string ToColumnName(int column) {
            var builder = new StringBuilder();
            while (column > 0) {
                column--;
                builder.Insert(0, (char)('A' + (column % 26)));
                column /= 26;
            }
            return builder.ToString();
        }

        private static string ToHex(GoogleSheetsNativeColor color) {
            int red = ToByte(color.Red);
            int green = ToByte(color.Green);
            int blue = ToByte(color.Blue);
            return $"#{red:X2}{green:X2}{blue:X2}";
        }

        private static int ToByte(double component) => Math.Max(0, Math.Min(255, (int)Math.Round(component * 255d)));
        private static string EscapeSheetName(string name) => name.Replace("'", "''");
    }
}
