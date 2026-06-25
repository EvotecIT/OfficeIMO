using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static class LegacyXlsWorkbookProjector {
        internal static ExcelDocument ToExcelDocument(LegacyXlsWorkbook workbook) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            var packageStream = new MemoryStream();
            ExcelDocument document = ExcelDocument.Create(packageStream, autoSave: false);
            foreach (LegacyXlsWorksheet legacySheet in workbook.Worksheets) {
                ExcelSheet sheet = document.AddWorkSheet(legacySheet.Name);
                ProjectWorksheet(workbook, legacySheet, sheet);
            }

            ProjectDefinedNames(workbook, document);
            ProjectAutoFilters(workbook, document);
            ProjectExternalReferences(workbook, document);

            if (workbook.Protection?.IsProtected == true) {
                document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                    LegacyPasswordHash = workbook.Protection.LegacyPasswordHash
                });
            }

            return document;
        }

        private static void ProjectWorksheet(LegacyXlsWorkbook workbook, LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            sheet.Batch(currentSheet => {
                foreach (LegacyXlsCell cell in legacySheet.Cells) {
                    LegacyXlsCellFormat? format = workbook.GetEffectiveCellFormat(cell.StyleIndex);
                    if (cell.Kind == LegacyXlsCellValueKind.Blank) {
                        ApplyNumberFormat(currentSheet, cell, format);
                        ApplyFont(currentSheet, workbook, cell, format);
                        ApplyFill(currentSheet, workbook, cell, format);
                        ApplyAlignment(currentSheet, cell, format);
                        ApplyBorder(currentSheet, workbook, cell, format);
                        ApplyProtection(currentSheet, cell, format);
                        ApplyQuotePrefix(currentSheet, cell, format);
                        continue;
                    }

                    object? value = GetProjectedCellValue(workbook, cell, format);
                    if (cell.Kind == LegacyXlsCellValueKind.Error && value is string errorText) {
                        currentSheet.SetLegacyErrorCellValue(cell.Row, cell.Column, errorText);
                    } else {
                        currentSheet.CellValue(cell.Row, cell.Column, value);
                    }

                    if (cell.IsFormula && !string.IsNullOrWhiteSpace(cell.FormulaText)) {
                        currentSheet.CellFormula(cell.Row, cell.Column, cell.FormulaText!);
                    }

                    ApplyNumberFormat(currentSheet, cell, format);
                    ApplyFont(currentSheet, workbook, cell, format);
                    ApplyFill(currentSheet, workbook, cell, format);
                    ApplyAlignment(currentSheet, cell, format);
                    ApplyBorder(currentSheet, workbook, cell, format);
                    ApplyProtection(currentSheet, cell, format);
                    ApplyQuotePrefix(currentSheet, cell, format);
                }

                if (legacySheet.DefaultColumnWidth.HasValue) {
                    currentSheet.SetDefaultColumnWidth(legacySheet.DefaultColumnWidth.Value, save: false);
                }

                foreach (LegacyXlsColumnLayout column in legacySheet.Columns) {
                    LegacyXlsCellFormat? columnFormat = workbook.GetEffectiveCellFormat(column.StyleIndex);
                    uint? projectedColumnStyleIndex = columnFormat != null
                        ? currentSheet.GetOrCreateLegacyCellFormatStyleIndex(workbook, columnFormat)
                        : null;

                    for (int columnIndex = column.StartColumn; columnIndex <= column.EndColumn; columnIndex++) {
                        if (column.Width > 0) {
                            currentSheet.SetColumnWidth(columnIndex, column.Width);
                        }

                        if (column.Hidden) {
                            currentSheet.SetColumnHidden(columnIndex, true);
                        }

                        if (column.OutlineLevel > 0 || column.Collapsed) {
                            currentSheet.SetColumnOutline(columnIndex, column.OutlineLevel, column.Collapsed, save: false);
                        }

                        if (projectedColumnStyleIndex.HasValue && projectedColumnStyleIndex.Value != 0U) {
                            currentSheet.SetColumnStyleIndex(columnIndex, projectedColumnStyleIndex.Value, save: false);
                        }
                    }
                }

                if (legacySheet.DefaultRowHeight.HasValue) {
                    currentSheet.SetDefaultRowHeight(legacySheet.DefaultRowHeight.Value, legacySheet.DefaultRowsHidden, save: false);
                }

                foreach (LegacyXlsRowLayout row in legacySheet.Rows) {
                    if (row.CustomHeight && row.Height > 0) {
                        currentSheet.SetRowHeight(row.Row, row.Height);
                    }

                    if (row.Hidden) {
                        currentSheet.SetRowHidden(row.Row, true);
                    }

                    if (row.OutlineLevel > 0 || row.Collapsed) {
                        currentSheet.SetRowOutline(row.Row, row.OutlineLevel, row.Collapsed, save: false);
                    }

                    if (row.StyleIndex.HasValue) {
                        LegacyXlsCellFormat? rowFormat = workbook.GetEffectiveCellFormat(row.StyleIndex.Value);
                        if (rowFormat != null) {
                            uint projectedRowStyleIndex = currentSheet.GetOrCreateLegacyCellFormatStyleIndex(workbook, rowFormat);
                            if (projectedRowStyleIndex != 0U) {
                                currentSheet.SetRowStyleIndex(row.Row, projectedRowStyleIndex, save: false);
                            }
                        }
                    }
                }

                foreach (LegacyXlsMergedRange mergedRange in legacySheet.MergedRanges) {
                    currentSheet.MergeRange(ToA1Range(mergedRange));
                }

                if (legacySheet.FreezePane != null) {
                    currentSheet.Freeze(legacySheet.FreezePane.TopRows, legacySheet.FreezePane.LeftColumns);
                }

                if (legacySheet.ZoomScale.HasValue) {
                    currentSheet.SetZoomScale(legacySheet.ZoomScale.Value, save: false);
                }

                if (legacySheet.ShowGridLines.HasValue) {
                    currentSheet.SetGridlinesVisible(legacySheet.ShowGridLines.Value);
                }

                if (legacySheet.ShowRowColumnHeadings.HasValue) {
                    currentSheet.SetRowColumnHeadingsVisible(legacySheet.ShowRowColumnHeadings.Value);
                }

                if (legacySheet.ShowZeroValues.HasValue) {
                    currentSheet.SetZeroValuesVisible(legacySheet.ShowZeroValues.Value);
                }

                if (legacySheet.RightToLeft.HasValue) {
                    currentSheet.SetRightToLeft(legacySheet.RightToLeft.Value);
                }

                foreach (LegacyXlsSelection selection in legacySheet.Selections) {
                    ProjectSelection(currentSheet, selection, legacySheet.FreezePane != null);
                }

                foreach (LegacyXlsHyperlink hyperlink in legacySheet.Hyperlinks) {
                    string reference = ToA1Range(hyperlink);
                    if (hyperlink.IsExternal) {
                        currentSheet.AddExternalHyperlinkReference(reference, hyperlink.Target);
                    } else {
                        currentSheet.AddInternalHyperlinkReference(reference, hyperlink.Target, normalizeLocation: false);
                    }
                }

                foreach (LegacyXlsDataValidation validation in legacySheet.DataValidations) {
                    LegacyXlsDataValidationProjector.Project(workbook, currentSheet, validation);
                }

                foreach (LegacyXlsConditionalFormatting conditionalFormatting in legacySheet.ConditionalFormattings) {
                    LegacyXlsConditionalFormattingProjector.Project(currentSheet, conditionalFormatting);
                }
            });

            foreach (LegacyXlsComment comment in legacySheet.Comments) {
                if (TryCreateCommentRichTextRuns(workbook, comment, out IReadOnlyList<ExcelRichTextRun> richTextRuns)) {
                    sheet.SetCommentRichText(comment.Row, comment.Column, richTextRuns, comment.Author);
                } else {
                    sheet.SetComment(comment.Row, comment.Column, comment.Text, comment.Author);
                }
            }

            if (legacySheet.Protection?.IsProtected == true) {
                sheet.Protect(new ExcelSheetProtectionOptions {
                    LegacyPasswordHash = legacySheet.Protection.LegacyPasswordHash,
                    ProtectObjects = legacySheet.Protection.ProtectObjects,
                    ProtectScenarios = legacySheet.Protection.ProtectScenarios
                });
            }

            ProjectPageSetup(legacySheet, sheet);
            ProjectPageBreaks(legacySheet, sheet);

            if (legacySheet.Visibility == 2) {
                sheet.SetVeryHidden(true);
            } else if (legacySheet.Visibility != 0) {
                sheet.SetHidden(true);
            }
        }

        private static void ProjectSelection(ExcelSheet sheet, LegacyXlsSelection selection, bool hasFrozenPane) {
            string activeCell = A1.CellReference(selection.ActiveRow, selection.ActiveColumn);
            IReadOnlyList<string> selectedRanges = selection.SelectedRanges
                .Select(range => range.Reference)
                .ToArray();
            sheet.SetWorksheetSelection(activeCell, selectedRanges, hasFrozenPane ? ToPane(selection.Pane) : null, save: false);
        }

        private static PaneValues? ToPane(byte pane) {
            return pane switch {
                0x00 => PaneValues.BottomRight,
                0x01 => PaneValues.TopRight,
                0x02 => PaneValues.BottomLeft,
                _ => null
            };
        }

        private static void ProjectDefinedNames(LegacyXlsWorkbook workbook, ExcelDocument document) {
            foreach (LegacyXlsDefinedName definedName in workbook.DefinedNames) {
                ExcelSheet? scope = definedName.LocalSheetIndex.HasValue && definedName.LocalSheetIndex.Value < document.Sheets.Count
                    ? document.Sheets[definedName.LocalSheetIndex.Value]
                    : null;
                if (scope != null
                    && string.Equals(definedName.Name, "_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase)
                    && TryParsePrintTitles(definedName.Reference, scope, out int? firstRow, out int? lastRow, out int? firstColumn, out int? lastColumn)) {
                    document.SetPrintTitles(scope, firstRow, lastRow, firstColumn, lastColumn, save: false);
                    continue;
                }

                if (scope != null
                    && string.Equals(definedName.Name, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)
                    && TryParseScopedRange(definedName.Reference, scope, out string autoFilterRange)) {
                    scope.AddAutoFilter(autoFilterRange);
                }

                document.SetNamedRange(
                    definedName.Name,
                    definedName.Reference,
                    scope,
                    save: false,
                    hidden: definedName.Hidden,
                    validationMode: NameValidationMode.Strict);
            }
        }

        private static void ProjectAutoFilters(LegacyXlsWorkbook workbook, ExcelDocument document) {
            for (int i = 0; i < workbook.Worksheets.Count && i < document.Sheets.Count; i++) {
                LegacyXlsWorksheet legacySheet = workbook.Worksheets[i];
                if (legacySheet.AutoFilterCriteria.Count == 0) {
                    continue;
                }

                ExcelSheet sheet = document.Sheets[i];
                if (!TryGetAutoFilterRange(workbook, sheet, i, out string autoFilterRange)) {
                    continue;
                }

                LegacyXlsAutoFilterProjector.Project(sheet, autoFilterRange, legacySheet.AutoFilterCriteria);
            }
        }

        private static bool TryGetAutoFilterRange(LegacyXlsWorkbook workbook, ExcelSheet sheet, int sheetIndex, out string autoFilterRange) {
            autoFilterRange = string.Empty;
            foreach (LegacyXlsDefinedName definedName in workbook.DefinedNames) {
                if (definedName.LocalSheetIndex == sheetIndex
                    && string.Equals(definedName.Name, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)
                    && TryParseScopedRange(definedName.Reference, sheet, out autoFilterRange)) {
                    return true;
                }
            }

            return false;
        }

        private static void ProjectExternalReferences(LegacyXlsWorkbook workbook, ExcelDocument document) {
            WorkbookPart workbookPart = document.WorkbookPartRoot;
            Workbook workbookRoot = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is null.");

            foreach (LegacyXlsExternalReference reference in workbook.ExternalReferences) {
                if (reference.Kind != LegacyXlsExternalReferenceKind.ExternalWorkbook
                    || string.IsNullOrWhiteSpace(reference.Target)) {
                    continue;
                }

                if (!TryCreateExternalTargetUri(reference.Target!, out Uri? targetUri) || targetUri == null) {
                    continue;
                }

                ExternalWorkbookPart externalWorkbookPart = workbookPart.AddNewPart<ExternalWorkbookPart>();
                ExternalRelationship relationship = externalWorkbookPart.AddExternalRelationship(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
                    targetUri);

                var externalBook = new ExternalBook { Id = relationship.Id };
                if (reference.SheetNames.Count > 0) {
                    externalBook.SheetNames = new SheetNames(reference.SheetNames
                        .Where(sheetName => !string.IsNullOrWhiteSpace(sheetName))
                        .Select(sheetName => new SheetName { Val = sheetName }));
                }

                ExternalDefinedNames? externalDefinedNames = CreateExternalDefinedNames(reference);
                if (externalDefinedNames != null) {
                    externalBook.ExternalDefinedNames = externalDefinedNames;
                }

                externalWorkbookPart.ExternalLink = new ExternalLink(externalBook);
                externalWorkbookPart.ExternalLink.Save();

                ExternalReferences externalReferences = GetOrCreateExternalReferences(workbookRoot);
                externalReferences.Append(new DocumentFormat.OpenXml.Spreadsheet.ExternalReference {
                    Id = workbookPart.GetIdOfPart(externalWorkbookPart)
                });
            }
        }

        private static ExternalReferences GetOrCreateExternalReferences(Workbook workbook) {
            if (workbook.ExternalReferences != null) {
                return workbook.ExternalReferences;
            }

            var externalReferences = new ExternalReferences();
            OpenXmlElement? before = workbook.GetFirstChild<DefinedNames>();
            before ??= workbook.GetFirstChild<CalculationProperties>();
            before ??= workbook.GetFirstChild<OleSize>();
            before ??= workbook.GetFirstChild<CustomWorkbookViews>();
            before ??= workbook.GetFirstChild<PivotCaches>();
            before ??= workbook.GetFirstChild<WebPublishing>();
            before ??= workbook.GetFirstChild<FileRecoveryProperties>();
            before ??= workbook.GetFirstChild<WebPublishObjects>();
            before ??= workbook.GetFirstChild<WorkbookExtensionList>();
            if (before != null) {
                workbook.InsertBefore(externalReferences, before);
            } else {
                workbook.Append(externalReferences);
            }

            return externalReferences;
        }

        private static ExternalDefinedNames? CreateExternalDefinedNames(LegacyXlsExternalReference reference) {
            List<ExternalDefinedName> names = reference.ExternalNames
                .Where(name => !string.IsNullOrWhiteSpace(name.Name))
                .Select(name => {
                    var projectedName = new ExternalDefinedName {
                        Name = name.Name
                    };
                    if (name.LocalSheetIndex.HasValue && name.LocalSheetIndex.Value >= 0) {
                        projectedName.SheetId = (uint)name.LocalSheetIndex.Value;
                    }

                    return projectedName;
                })
                .ToList();

            return names.Count == 0 ? null : new ExternalDefinedNames(names);
        }

        private static bool TryCreateExternalTargetUri(string target, out Uri? uri) {
            string normalized = RemoveControlCharacters(target);
            uri = null;
            if (string.IsNullOrWhiteSpace(normalized)) {
                return false;
            }

            return Uri.TryCreate(normalized, UriKind.RelativeOrAbsolute, out uri);
        }

        private static string RemoveControlCharacters(string value) {
            StringBuilder? sanitized = null;
            for (int i = 0; i < value.Length; i++) {
                if (!char.IsControl(value[i])) {
                    sanitized?.Append(value[i]);
                    continue;
                }

                sanitized ??= new StringBuilder(value.Length).Append(value, 0, i);
            }

            return sanitized == null ? value : sanitized.ToString();
        }

        private static bool TryCreateCommentRichTextRuns(
            LegacyXlsWorkbook workbook,
            LegacyXlsComment comment,
            out IReadOnlyList<ExcelRichTextRun> richTextRuns) {
            richTextRuns = Array.Empty<ExcelRichTextRun>();
            if (comment.FormattingRuns.Count == 0 || string.IsNullOrEmpty(comment.Text)) {
                return false;
            }

            List<LegacyXlsCommentFormattingRun> runs = comment.FormattingRuns
                .Where(run => run.StartCharacter < comment.Text.Length)
                .OrderBy(run => run.StartCharacter)
                .ToList();
            if (runs.Count == 0) {
                return false;
            }

            if (runs[0].StartCharacter != 0) {
                runs.Insert(0, new LegacyXlsCommentFormattingRun(0, 0));
            }

            var projectedRuns = new List<ExcelRichTextRun>(runs.Count);
            for (int i = 0; i < runs.Count; i++) {
                int start = runs[i].StartCharacter;
                int end = i + 1 < runs.Count ? runs[i + 1].StartCharacter : comment.Text.Length;
                if (end <= start) {
                    continue;
                }

                var projectedRun = new ExcelRichTextRun(comment.Text.Substring(start, end - start));
                LegacyXlsFont? font = workbook.GetFont(runs[i].FontIndex);
                if (font != null) {
                    projectedRun.Bold = font.Bold;
                    projectedRun.Italic = font.Italic;
                    projectedRun.Underline = font.Underline;
                    projectedRun.FontName = font.Name;
                    projectedRun.FontSize = font.Size;
                    if (workbook.TryResolveColor(font.ColorIndex, out string? fontColor)) {
                        projectedRun.FontColor = fontColor;
                    }
                }

                projectedRuns.Add(projectedRun);
            }

            richTextRuns = projectedRuns;
            return projectedRuns.Count > 0;
        }

        private static bool TryParsePrintTitles(
            string reference,
            ExcelSheet scope,
            out int? firstRow,
            out int? lastRow,
            out int? firstColumn,
            out int? lastColumn) {
            firstRow = null;
            lastRow = null;
            firstColumn = null;
            lastColumn = null;

            bool parsedAny = false;
            foreach (string part in SplitDefinedNameReferenceList(reference)) {
                if (!SheetNameLookup.TryParseSheetQualifiedReference(part, out string sheetName, out string localReference, allowExternalWorkbookReferences: false)
                    || !SheetNameLookup.Matches(scope.Name, sheetName)) {
                    return false;
                }

                string normalized = localReference.Replace("$", string.Empty);
                int separator = normalized.IndexOf(':');
                if (separator <= 0 || separator >= normalized.Length - 1) {
                    return false;
                }

                string start = normalized.Substring(0, separator);
                string end = normalized.Substring(separator + 1);
                if (int.TryParse(start, NumberStyles.None, CultureInfo.InvariantCulture, out int parsedFirstRow)
                    && int.TryParse(end, NumberStyles.None, CultureInfo.InvariantCulture, out int parsedLastRow)
                    && parsedFirstRow > 0
                    && parsedLastRow >= parsedFirstRow) {
                    firstRow = parsedFirstRow;
                    lastRow = parsedLastRow;
                    parsedAny = true;
                    continue;
                }

                int parsedFirstColumn = A1.ColumnLettersToIndex(start);
                int parsedLastColumn = A1.ColumnLettersToIndex(end);
                if (parsedFirstColumn <= 0 || parsedLastColumn < parsedFirstColumn) {
                    return false;
                }

                firstColumn = parsedFirstColumn;
                lastColumn = parsedLastColumn;
                parsedAny = true;
            }

            return parsedAny;
        }

        private static bool TryParseScopedRange(string reference, ExcelSheet scope, out string localRange) {
            localRange = string.Empty;
            if (!SheetNameLookup.TryParseSheetQualifiedReference(reference, out string sheetName, out string parsedReference, allowExternalWorkbookReferences: false)
                || !SheetNameLookup.Matches(scope.Name, sheetName)) {
                return false;
            }

            localRange = parsedReference.Replace("$", string.Empty);
            return localRange.Length > 0;
        }

        private static IReadOnlyList<string> SplitDefinedNameReferenceList(string text) {
            var parts = new List<string>();
            var current = new System.Text.StringBuilder(text.Length);
            bool inQuote = false;

            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch == '\'') {
                    current.Append(ch);
                    if (inQuote && i + 1 < text.Length && text[i + 1] == '\'') {
                        current.Append(text[++i]);
                    } else {
                        inQuote = !inQuote;
                    }
                    continue;
                }

                if (ch == ',' && !inQuote) {
                    string part = current.ToString().Trim();
                    if (part.Length > 0) {
                        parts.Add(part);
                    }

                    current.Clear();
                    continue;
                }

                current.Append(ch);
            }

            string finalPart = current.ToString().Trim();
            if (finalPart.Length > 0) {
                parts.Add(finalPart);
            }

            return parts;
        }

        private static void ProjectPageBreaks(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            foreach (LegacyXlsPageBreak pageBreak in legacySheet.RowPageBreaks) {
                sheet.AddManualRowPageBreak(pageBreak.Position, save: false);
            }

            foreach (LegacyXlsPageBreak pageBreak in legacySheet.ColumnPageBreaks) {
                sheet.AddManualColumnPageBreak(pageBreak.Position, save: false);
            }
        }

        private static void ProjectPageSetup(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            LegacyXlsPageSetup? pageSetup = legacySheet.PageSetup;
            if (pageSetup == null) {
                return;
            }

            if (pageSetup.LeftMargin.HasValue
                || pageSetup.RightMargin.HasValue
                || pageSetup.TopMargin.HasValue
                || pageSetup.BottomMargin.HasValue
                || pageSetup.HeaderMargin.HasValue
                || pageSetup.FooterMargin.HasValue) {
                sheet.SetMargins(
                    pageSetup.LeftMargin ?? 0.7d,
                    pageSetup.RightMargin ?? 0.7d,
                    pageSetup.TopMargin ?? 0.75d,
                    pageSetup.BottomMargin ?? 0.75d,
                    pageSetup.HeaderMargin ?? 0.3d,
                    pageSetup.FooterMargin ?? 0.3d);
            }

            if (pageSetup.Landscape.HasValue) {
                sheet.SetOrientation(pageSetup.Landscape.Value ? ExcelPageOrientation.Landscape : ExcelPageOrientation.Portrait);
            }

            if (pageSetup.PrintGridLines.HasValue
                || pageSetup.PrintHeadings.HasValue
                || pageSetup.HorizontalCentered.HasValue
                || pageSetup.VerticalCentered.HasValue) {
                sheet.SetPrintOptions(
                    pageSetup.PrintGridLines,
                    pageSetup.PrintHeadings,
                    pageSetup.HorizontalCentered,
                    pageSetup.VerticalCentered,
                    save: false);
            }

            uint? fitToWidth = pageSetup.FitToWidth.HasValue && pageSetup.FitToWidth.Value > 0
                ? pageSetup.FitToWidth.Value
                : null;
            uint? fitToHeight = pageSetup.FitToHeight.HasValue && pageSetup.FitToHeight.Value > 0
                ? pageSetup.FitToHeight.Value
                : null;
            uint? scale = pageSetup.Scale.HasValue ? pageSetup.Scale.Value : null;
            if (fitToWidth.HasValue || fitToHeight.HasValue || scale.HasValue) {
                sheet.SetPageSetup(fitToWidth, fitToHeight, scale);
            }

            if (pageSetup.FitToPage.HasValue) {
                sheet.SetFitToPage(pageSetup.FitToPage.Value);
            }

            if (!string.IsNullOrEmpty(pageSetup.HeaderText) || !string.IsNullOrEmpty(pageSetup.FooterText)) {
                var header = SplitHeaderFooterText(pageSetup.HeaderText);
                var footer = SplitHeaderFooterText(pageSetup.FooterText);
                sheet.SetHeaderFooter(
                    header.Left,
                    header.Center,
                    header.Right,
                    footer.Left,
                    footer.Center,
                    footer.Right);
            }
        }

        private static (string? Left, string? Center, string? Right) SplitHeaderFooterText(string? text) {
            if (string.IsNullOrEmpty(text)) {
                return (null, null, null);
            }

            if (!ContainsHeaderFooterSectionMarker(text!)) {
                return (null, text, null);
            }

            string? left = null;
            string? center = null;
            string? right = null;
            int i = 0;
            while (i < text!.Length) {
                if (text[i] != '&' || i + 1 >= text.Length) {
                    i++;
                    continue;
                }

                char section = text[i + 1];
                if (section != 'L' && section != 'C' && section != 'R') {
                    i += 2;
                    continue;
                }

                i += 2;
                int start = i;
                while (i < text.Length) {
                    if (text[i] == '&' && i + 1 < text.Length) {
                        char next = text[i + 1];
                        if (next == 'L' || next == 'C' || next == 'R') {
                            break;
                        }
                    }

                    i++;
                }

                string value = text.Substring(start, i - start);
                if (section == 'L') {
                    left = AppendHeaderFooterSection(left, value);
                } else if (section == 'C') {
                    center = AppendHeaderFooterSection(center, value);
                } else {
                    right = AppendHeaderFooterSection(right, value);
                }
            }

            return (left, center, right);
        }

        private static bool ContainsHeaderFooterSectionMarker(string text) {
            return text.IndexOf("&L", StringComparison.Ordinal) >= 0
                || text.IndexOf("&C", StringComparison.Ordinal) >= 0
                || text.IndexOf("&R", StringComparison.Ordinal) >= 0;
        }

        private static string? AppendHeaderFooterSection(string? current, string value) {
            if (value.Length == 0) {
                return current;
            }

            return current == null ? value : current + value;
        }

        private static object? GetProjectedCellValue(
            LegacyXlsWorkbook workbook,
            LegacyXlsCell cell,
            LegacyXlsCellFormat? format) {
            if (cell.Kind != LegacyXlsCellValueKind.Number || format?.IsDateLike != true || cell.Value is not double serial) {
                return cell.Value;
            }

            return LegacyXlsDateSerialConverter.TryConvert(serial, workbook.Uses1904DateSystem, out DateTime value)
                ? value
                : cell.Value;
        }

        private static void ApplyNumberFormat(ExcelSheet sheet, LegacyXlsCell cell, LegacyXlsCellFormat? format) {
            if (format?.NumberFormatCode == null || format.NumberFormatId == 0) {
                return;
            }

            if (format.IsBuiltInNumberFormat) {
                sheet.FormatCellBuiltInNumberFormat(cell.Row, cell.Column, format.NumberFormatId);
                return;
            }

            sheet.FormatCell(cell.Row, cell.Column, format.NumberFormatCode);
        }

        private static void ApplyFont(
            ExcelSheet sheet,
            LegacyXlsWorkbook workbook,
            LegacyXlsCell cell,
            LegacyXlsCellFormat? format) {
            if (format == null) {
                return;
            }

            LegacyXlsFont? font = workbook.GetFont(format.FontIndex);
            if (font == null) {
                return;
            }

            workbook.TryResolveColor(font.ColorIndex, out string? fontColor);
            VerticalAlignmentRunValues? verticalTextAlignment = ToVerticalTextAlignment(font.Escapement);
            if (font.Name == null && !font.Size.HasValue && fontColor == null && !font.Bold && !font.Italic && !font.Underline && !font.Strikeout && !verticalTextAlignment.HasValue) {
                return;
            }

            sheet.FormatCellFont(
                cell.Row,
                cell.Column,
                font.Name,
                font.Size,
                fontColor,
                font.Bold,
                font.Italic,
                font.Underline,
                font.Strikeout,
                verticalTextAlignment);
        }

        private static VerticalAlignmentRunValues? ToVerticalTextAlignment(LegacyXlsFontEscapement escapement) {
            return escapement == LegacyXlsFontEscapement.Superscript
                ? VerticalAlignmentRunValues.Superscript
                : escapement == LegacyXlsFontEscapement.Subscript
                    ? VerticalAlignmentRunValues.Subscript
                    : null;
        }

        private static void ApplyFill(
            ExcelSheet sheet,
            LegacyXlsWorkbook workbook,
            LegacyXlsCell cell,
            LegacyXlsCellFormat? format) {
            if (format == null || format.FillPattern == 0) {
                return;
            }

            PatternValues? pattern = ToFillPattern(format.FillPattern);
            if (!pattern.HasValue) {
                return;
            }

            string? foregroundColor = ResolveColor(workbook, format.FillForegroundColorIndex);
            string? backgroundColor = ResolveColor(workbook, format.FillBackgroundColorIndex);
            if (foregroundColor == null && backgroundColor == null) {
                return;
            }

            if (format.FillPattern == 1 && foregroundColor != null) {
                sheet.FormatCellFill(cell.Row, cell.Column, PatternValues.Solid, foregroundColor, foregroundColor);
                return;
            }

            sheet.FormatCellFill(cell.Row, cell.Column, pattern.Value, foregroundColor, backgroundColor);
        }

        private static void ApplyAlignment(ExcelSheet sheet, LegacyXlsCell cell, LegacyXlsCellFormat? format) {
            if (format?.ApplyAlignment != true) {
                return;
            }

            sheet.FormatCellAlignment(
                cell.Row,
                cell.Column,
                ToHorizontalAlignment(format.HorizontalAlignment),
                ToVerticalAlignment(format.VerticalAlignment),
                format.WrapText,
                ToTextRotation(format.TextRotation),
                format.Indent,
                format.ShrinkToFit,
                ToReadingOrder(format.ReadingOrder));
        }

        private static HorizontalAlignmentValues? ToHorizontalAlignment(byte alignment) {
            return alignment switch {
                1 => HorizontalAlignmentValues.Left,
                2 => HorizontalAlignmentValues.Center,
                3 => HorizontalAlignmentValues.Right,
                4 => HorizontalAlignmentValues.Fill,
                5 => HorizontalAlignmentValues.Justify,
                6 => HorizontalAlignmentValues.CenterContinuous,
                7 => HorizontalAlignmentValues.Distributed,
                _ => null
            };
        }

        private static VerticalAlignmentValues? ToVerticalAlignment(byte alignment) {
            return alignment switch {
                0 => VerticalAlignmentValues.Top,
                1 => VerticalAlignmentValues.Center,
                2 => VerticalAlignmentValues.Bottom,
                3 => VerticalAlignmentValues.Justify,
                4 => VerticalAlignmentValues.Distributed,
                _ => null
            };
        }

        private static uint? ToTextRotation(byte rotation) {
            return rotation <= 180 || rotation == 255 ? rotation : null;
        }

        private static uint? ToReadingOrder(byte readingOrder) {
            return readingOrder <= 2 ? readingOrder : null;
        }

        private static void ApplyBorder(
            ExcelSheet sheet,
            LegacyXlsWorkbook workbook,
            LegacyXlsCell cell,
            LegacyXlsCellFormat? format) {
            if (format?.Border == null) {
                return;
            }

            LegacyXlsBorder border = format.Border;
            sheet.FormatCellBorder(
                cell.Row,
                cell.Column,
                ToBorderStyle(border.LeftStyle),
                ResolveColor(workbook, border.LeftColorIndex),
                ToBorderStyle(border.RightStyle),
                ResolveColor(workbook, border.RightColorIndex),
                ToBorderStyle(border.TopStyle),
                ResolveColor(workbook, border.TopColorIndex),
                ToBorderStyle(border.BottomStyle),
                ResolveColor(workbook, border.BottomColorIndex),
                ToBorderStyle(border.DiagonalStyle),
                ResolveColor(workbook, border.DiagonalColorIndex),
                border.DiagonalUp,
                border.DiagonalDown);
        }

        private static void ApplyProtection(ExcelSheet sheet, LegacyXlsCell cell, LegacyXlsCellFormat? format) {
            if (format?.ApplyProtection != true) {
                return;
            }

            sheet.FormatCellProtection(cell.Row, cell.Column, format.Locked, format.FormulaHidden);
        }

        private static void ApplyQuotePrefix(ExcelSheet sheet, LegacyXlsCell cell, LegacyXlsCellFormat? format) {
            if (format?.QuotePrefix != true) {
                return;
            }

            sheet.FormatCellQuotePrefix(cell.Row, cell.Column, true);
        }

        private static BorderStyleValues? ToBorderStyle(byte style) {
            return style switch {
                1 => BorderStyleValues.Thin,
                2 => BorderStyleValues.Medium,
                3 => BorderStyleValues.Dashed,
                4 => BorderStyleValues.Dotted,
                5 => BorderStyleValues.Thick,
                6 => BorderStyleValues.Double,
                7 => BorderStyleValues.Hair,
                8 => BorderStyleValues.MediumDashed,
                9 => BorderStyleValues.DashDot,
                10 => BorderStyleValues.MediumDashDot,
                11 => BorderStyleValues.DashDotDot,
                12 => BorderStyleValues.MediumDashDotDot,
                13 => BorderStyleValues.SlantDashDot,
                _ => null
            };
        }

        private static PatternValues? ToFillPattern(byte pattern) {
            return pattern switch {
                1 => PatternValues.Solid,
                2 => PatternValues.MediumGray,
                3 => PatternValues.DarkGray,
                4 => PatternValues.LightGray,
                5 => PatternValues.DarkHorizontal,
                6 => PatternValues.DarkVertical,
                7 => PatternValues.DarkDown,
                8 => PatternValues.DarkUp,
                9 => PatternValues.DarkGrid,
                10 => PatternValues.DarkTrellis,
                11 => PatternValues.LightHorizontal,
                12 => PatternValues.LightVertical,
                13 => PatternValues.LightDown,
                14 => PatternValues.LightUp,
                15 => PatternValues.LightGrid,
                16 => PatternValues.LightTrellis,
                17 => PatternValues.Gray125,
                18 => PatternValues.Gray0625,
                _ => null
            };
        }

        private static string? ResolveColor(LegacyXlsWorkbook workbook, ushort colorIndex) {
            return workbook.TryResolveColor(colorIndex, out string? color) ? color : null;
        }

        private static string ToA1Range(LegacyXlsMergedRange mergedRange) {
            string start = A1.CellReference(mergedRange.StartRow, mergedRange.StartColumn);
            string end = A1.CellReference(mergedRange.EndRow, mergedRange.EndColumn);
            return start == end ? start : start + ":" + end;
        }

        private static string ToA1Range(LegacyXlsHyperlink hyperlink) {
            string start = A1.CellReference(hyperlink.StartRow, hyperlink.StartColumn);
            string end = A1.CellReference(hyperlink.EndRow, hyperlink.EndColumn);
            return start == end ? start : start + ":" + end;
        }
    }
}
