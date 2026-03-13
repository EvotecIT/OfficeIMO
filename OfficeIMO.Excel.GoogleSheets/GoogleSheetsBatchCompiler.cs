using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static class GoogleSheetsBatchCompiler {
        private const string DefaultTableFooterColorArgb = "FFE8EAED";

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

            bool styleNoticeAdded = false;
            bool formulaNoticeAdded = false;
            bool builtInNameNoticeAdded = false;
            bool tableStyleNoticeAdded = false;
            bool multipleFilterNoticeAdded = false;
            bool protectionNoticeAdded = false;
            bool protectionPermissionsNoticeAdded = false;
            bool tableTotalsNoticeAdded = false;
            bool customFilterNoticeAdded = false;
            bool cellValidationNoticeAdded = false;

            foreach (var worksheet in workbookSnapshot.Worksheets) {
                batch.Add(new GoogleSheetsAddSheetRequest {
                    SheetName = worksheet.Name,
                    SheetIndex = worksheet.Index,
                    Hidden = worksheet.Hidden,
                    RightToLeft = worksheet.RightToLeft,
                    TabColorArgb = worksheet.TabColorArgb,
                    FrozenRowCount = worksheet.FrozenRowCount,
                    FrozenColumnCount = worksheet.FrozenColumnCount,
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
                        WarningOnly = false,
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
                    });
                }

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

                var updateCells = new GoogleSheetsUpdateCellsRequest {
                    SheetName = worksheet.Name
                };
                var emittedCellKeys = new HashSet<string>(StringComparer.Ordinal);

                foreach (var cell in worksheet.Cells) {
                    if (!styleNoticeAdded && cell.Style != null) {
                        report.Add(
                            TranslationSeverity.Info,
                            "Styles",
                            "The current Google Sheets compiler now emits basic cell styling, hyperlinks, and row/column dimensions alongside workbook structure, values, formulas, frozen panes, merges, and named ranges.");
                        styleNoticeAdded = true;
                    }

                    var cellValue = BuildCellValue(cell, options, report, ref formulaNoticeAdded);
                    emittedCellKeys.Add(CreateCellKey(cell.Row, cell.Column));
                    updateCells.AddCell(new GoogleSheetsCellData {
                        RowIndex = cell.Row - 1,
                        ColumnIndex = cell.Column - 1,
                        Value = cellValue,
                        NumberFormatHint = GetNumberFormatHint(cell.Value, cell.Style),
                        Style = BuildCellStyle(cell.Style),
                        DataValidationRule = BuildCellValidationRule(workbookSnapshot, worksheet, cell.Row, cell.Column, report, ref cellValidationNoticeAdded),
                        Hyperlink = BuildHyperlink(cell.Hyperlink),
                        Comment = BuildComment(cell.Comment),
                    });
                }

                AppendValidationOnlyCells(workbookSnapshot, worksheet, updateCells, emittedCellKeys, report, ref cellValidationNoticeAdded);

                if (updateCells.Cells.Count > 0) {
                    batch.Add(updateCells);
                }

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
            }

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

                batch.Add(new GoogleSheetsAddNamedRangeRequest {
                    Name = namedRange.Name,
                    SheetName = namedRange.SheetName,
                    A1Range = namedRange.ReferenceA1,
                });
            }

            return batch;
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

        private static GoogleSheetsCellValue BuildCellValue(
            ExcelCellSnapshot cell,
            GoogleSheetsSaveOptions options,
            TranslationReport report,
            ref bool formulaNoticeAdded) {
            if (!string.IsNullOrWhiteSpace(cell.Formula)) {
                if (!formulaNoticeAdded) {
                    var message = options.PreserveUnsupportedFormulasAsText
                        ? "Formula cells are compiled as formulas first; unsupported-formula fallback-to-text still needs an execution-stage compatibility map."
                        : "Formula cells are compiled as formulas, but function-by-function compatibility validation still needs to be implemented.";
                    report.Add(TranslationSeverity.Info, "FormulaExecution", message);
                    formulaNoticeAdded = true;
                }

                return GoogleSheetsCellValue.Formula(NormalizeFormula(cell.Formula!));
            }

            var typedValue = cell.Value;
            if (typedValue == null) {
                return GoogleSheetsCellValue.Blank();
            }

            if (typedValue is bool booleanValue) {
                return GoogleSheetsCellValue.Boolean(booleanValue);
            }

            if (typedValue is DateTime dateTimeValue) {
                return GoogleSheetsCellValue.DateTime(dateTimeValue);
            }

            if (typedValue is DateTimeOffset dateTimeOffsetValue) {
                return GoogleSheetsCellValue.DateTime(dateTimeOffsetValue.LocalDateTime);
            }

            if (typedValue is byte || typedValue is sbyte || typedValue is short || typedValue is ushort
                || typedValue is int || typedValue is uint || typedValue is long || typedValue is ulong
                || typedValue is float || typedValue is double || typedValue is decimal) {
                return GoogleSheetsCellValue.Number(Convert.ToDouble(typedValue, System.Globalization.CultureInfo.InvariantCulture));
            }

            return GoogleSheetsCellValue.String(Convert.ToString(typedValue, System.Globalization.CultureInfo.InvariantCulture));
        }

        private static string NormalizeFormula(string formulaText) {
            if (string.IsNullOrWhiteSpace(formulaText)) {
                return "=";
            }

            return formulaText[0] == '=' ? formulaText : "=" + formulaText;
        }

        private static string? GetNumberFormatHint(object? typedValue, ExcelCellStyleSnapshot? style) {
            if (style?.IsDateLike == true || typedValue is DateTime || typedValue is DateTimeOffset) {
                return "DateTime";
            }

            if (!string.IsNullOrWhiteSpace(style?.NumberFormatCode)) {
                return style!.NumberFormatCode;
            }

            return null;
        }

        private static GoogleSheetsCellStyle? BuildCellStyle(ExcelCellStyleSnapshot? style) {
            if (style == null) {
                return null;
            }

            return new GoogleSheetsCellStyle {
                SourceStyleIndex = style.StyleIndex,
                NumberFormatId = style.NumberFormatId,
                NumberFormatCode = style.NumberFormatCode,
                IsDateLike = style.IsDateLike,
                Bold = style.Bold,
                Italic = style.Italic,
                Underline = style.Underline,
                FontColorArgb = style.FontColorArgb,
                FillColorArgb = style.FillColorArgb,
                Borders = BuildBorders(style.Border),
                HorizontalAlignment = style.HorizontalAlignment,
                VerticalAlignment = style.VerticalAlignment,
                WrapText = style.WrapText,
            };
        }

        private static GoogleSheetsCellBorders? BuildBorders(ExcelCellBorderSnapshot? border) {
            if (border == null) {
                return null;
            }

            var left = BuildBorderSide(border.Left);
            var right = BuildBorderSide(border.Right);
            var top = BuildBorderSide(border.Top);
            var bottom = BuildBorderSide(border.Bottom);

            if (left == null && right == null && top == null && bottom == null) {
                return null;
            }

            return new GoogleSheetsCellBorders {
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom,
            };
        }

        private static GoogleSheetsBorderSide? BuildBorderSide(ExcelBorderSideSnapshot? side) {
            if (side == null) {
                return null;
            }

            if (string.IsNullOrWhiteSpace(side.Style) && string.IsNullOrWhiteSpace(side.ColorArgb)) {
                return null;
            }

            return new GoogleSheetsBorderSide {
                Style = side.Style,
                ColorArgb = side.ColorArgb,
            };
        }

        private static GoogleSheetsHyperlink? BuildHyperlink(ExcelHyperlinkSnapshot? hyperlink) {
            if (hyperlink == null) {
                return null;
            }

            return new GoogleSheetsHyperlink {
                IsExternal = hyperlink.IsExternal,
                Target = hyperlink.Target,
            };
        }

        private static GoogleSheetsComment? BuildComment(ExcelCommentSnapshot? comment) {
            if (comment == null || string.IsNullOrWhiteSpace(comment.Text)) {
                return null;
            }

            return new GoogleSheetsComment {
                Author = string.IsNullOrWhiteSpace(comment.Author) ? null : comment.Author,
                Text = comment.Text,
            };
        }

        private static IReadOnlyList<GoogleSheetsRequest> BuildFilterRequests(
            ExcelWorksheetSnapshot worksheet,
            TranslationReport report,
            ref bool multipleFilterNoticeAdded,
            ref bool customFilterNoticeAdded) {
            var requests = new List<GoogleSheetsRequest>();
            var filterSources = new List<(ExcelAutoFilterSnapshot Filter, string Title)>();

            if (worksheet.AutoFilter != null) {
                filterSources.Add((worksheet.AutoFilter, worksheet.Name + " Filter"));
            }

            foreach (var table in worksheet.Tables) {
                if (table.AutoFilter != null) {
                    var title = string.IsNullOrWhiteSpace(table.Name)
                        ? worksheet.Name + " Table Filter"
                        : table.Name + " Filter";
                    filterSources.Add((table.AutoFilter, title));
                }
            }

            if (filterSources.Count == 0) {
                return requests;
            }

            if (filterSources.Count > 1 && !multipleFilterNoticeAdded) {
                report.Add(
                    TranslationSeverity.Info,
                    "MultipleFilters",
                    "When multiple Excel filter ranges exist on one sheet, the first is emitted as the sheet basic filter and the rest are emitted as Google filter views.");
                multipleFilterNoticeAdded = true;
            }

            for (int i = 0; i < filterSources.Count; i++) {
                var source = filterSources[i];
                var criteria = BuildFilterCriteria(worksheet, source.Filter, report, ref customFilterNoticeAdded);
                if (i == 0) {
                    requests.Add(new GoogleSheetsSetBasicFilterRequest {
                        SheetName = worksheet.Name,
                        A1Range = source.Filter.A1Range,
                        StartRowIndex = source.Filter.StartRow - 1,
                        EndRowIndexExclusive = source.Filter.EndRow,
                        StartColumnIndex = source.Filter.StartColumn - 1,
                        EndColumnIndexExclusive = source.Filter.EndColumn,
                        Criteria = criteria,
                    });
                } else {
                    requests.Add(new GoogleSheetsAddFilterViewRequest {
                        SheetName = worksheet.Name,
                        Title = source.Title,
                        A1Range = source.Filter.A1Range,
                        StartRowIndex = source.Filter.StartRow - 1,
                        EndRowIndexExclusive = source.Filter.EndRow,
                        StartColumnIndex = source.Filter.StartColumn - 1,
                        EndColumnIndexExclusive = source.Filter.EndColumn,
                        Criteria = criteria,
                    });
                }
            }

            return requests;
        }

        private static IReadOnlyList<GoogleSheetsFilterColumnCriteria> BuildFilterCriteria(
            ExcelWorksheetSnapshot worksheet,
            ExcelAutoFilterSnapshot filter,
            TranslationReport report,
            ref bool customFilterNoticeAdded) {
            var criteria = new List<GoogleSheetsFilterColumnCriteria>();
            if (filter.Columns.Count == 0) {
                return criteria;
            }

            var cellMap = worksheet.Cells.ToDictionary(
                cell => GetWorksheetCellKey(cell.Row, cell.Column),
                cell => cell);

            foreach (var filterColumn in filter.Columns) {
                GoogleSheetsBooleanCondition? condition = null;
                if (filterColumn.CustomFilters != null) {
                    condition = BuildBooleanCondition(filterColumn.CustomFilters, report, ref customFilterNoticeAdded);
                }

                List<string> hiddenValues = new List<string>();
                if (filterColumn.Values.Count == 0) {
                } else {
                    var allowedValues = new HashSet<string>(filterColumn.Values, StringComparer.OrdinalIgnoreCase);
                    var observedValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var absoluteColumn = filter.StartColumn + filterColumn.ColumnId;

                    for (int row = filter.StartRow + 1; row <= filter.EndRow; row++) {
                        if (cellMap.TryGetValue(GetWorksheetCellKey(row, absoluteColumn), out var cell)) {
                            observedValues.Add(ConvertCellToFilterText(cell));
                        } else {
                            observedValues.Add(string.Empty);
                        }
                    }

                    hiddenValues = observedValues
                        .Where(value => !allowedValues.Contains(value))
                        .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }

                if (hiddenValues.Count == 0 && condition == null) {
                    continue;
                }

                criteria.Add(new GoogleSheetsFilterColumnCriteria {
                    ColumnId = filterColumn.ColumnId,
                    HiddenValues = hiddenValues,
                    Condition = condition,
                });
            }

            return criteria;
        }

        private static GoogleSheetsBooleanCondition? BuildBooleanCondition(
            ExcelCustomFiltersSnapshot customFilters,
            TranslationReport report,
            ref bool customFilterNoticeAdded) {
            if (customFilters.Conditions.Count == 2) {
                if (TryBuildNumericRangeCondition(customFilters, out var rangeCondition)) {
                    return rangeCondition;
                }

                AddUnsupportedCustomFilterNotice(report, ref customFilterNoticeAdded);
                return null;
            }

            if (customFilters.Conditions.Count != 1 || customFilters.MatchAll) {
                AddUnsupportedCustomFilterNotice(report, ref customFilterNoticeAdded);
                return null;
            }

            var condition = customFilters.Conditions[0];
            if (string.IsNullOrWhiteSpace(condition.Value)) {
                return null;
            }

            var value = condition.Value;
            var filterOperator = condition.Operator ?? "equal";
            if (TryBuildTextWildcardCondition(filterOperator, value, out var textCondition)) {
                return textCondition;
            }

            if (double.TryParse(value, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, System.Globalization.CultureInfo.InvariantCulture, out _)
                && TryBuildNumericCondition(filterOperator, value, out var numericCondition)) {
                return numericCondition;
            }

            if (TryBuildTextEqualityCondition(filterOperator, value, out var equalityCondition)) {
                return equalityCondition;
            }

            AddUnsupportedCustomFilterNotice(report, ref customFilterNoticeAdded);
            return null;
        }

        private static bool TryBuildNumericRangeCondition(
            ExcelCustomFiltersSnapshot customFilters,
            out GoogleSheetsBooleanCondition? condition) {
            condition = null;
            if (customFilters.Conditions.Count != 2) {
                return false;
            }

            var first = customFilters.Conditions[0];
            var second = customFilters.Conditions[1];
            if (!TryNormalizeNumericCondition(first, out var firstOperator, out var firstValue)
                || !TryNormalizeNumericCondition(second, out var secondOperator, out var secondValue)) {
                return false;
            }

            if (customFilters.MatchAll
                && TryGetBetweenBounds(firstOperator, firstValue, secondOperator, secondValue, out var lowerInclusive, out var upperInclusive)) {
                condition = new GoogleSheetsBooleanCondition {
                    Type = "NUMBER_BETWEEN",
                    Values = new[] { lowerInclusive, upperInclusive },
                };
                return true;
            }

            if (!customFilters.MatchAll
                && TryGetOutsideBounds(firstOperator, firstValue, secondOperator, secondValue, out var lowerExclusive, out var upperExclusive)) {
                condition = new GoogleSheetsBooleanCondition {
                    Type = "NUMBER_NOT_BETWEEN",
                    Values = new[] { lowerExclusive, upperExclusive },
                };
                return true;
            }

            return false;
        }

        private static bool TryBuildTextWildcardCondition(
            string filterOperator,
            string value,
            out GoogleSheetsBooleanCondition? condition) {
            condition = null;
            var normalizedOperator = filterOperator.Trim().ToLowerInvariant();
            var startsWithWildcard = value.StartsWith("*", StringComparison.Ordinal);
            var endsWithWildcard = value.EndsWith("*", StringComparison.Ordinal);
            var unwrappedValue = value.Trim('*');

            if (string.IsNullOrEmpty(unwrappedValue) || (!startsWithWildcard && !endsWithWildcard)) {
                return false;
            }

            string? conditionType = (normalizedOperator, startsWithWildcard, endsWithWildcard) switch {
                ("equal", true, true) => "TEXT_CONTAINS",
                ("equal", false, true) => "TEXT_STARTS_WITH",
                ("equal", true, false) => "TEXT_ENDS_WITH",
                ("notequal", true, true) => "TEXT_NOT_CONTAINS",
                _ => null,
            };

            if (conditionType == null) {
                return false;
            }

            condition = new GoogleSheetsBooleanCondition {
                Type = conditionType,
                Values = new[] { unwrappedValue },
            };
            return true;
        }

        private static bool TryBuildTextEqualityCondition(
            string filterOperator,
            string value,
            out GoogleSheetsBooleanCondition? condition) {
            condition = null;
            var normalizedOperator = filterOperator.Trim().ToLowerInvariant();
            string? conditionType = normalizedOperator switch {
                "equal" => "TEXT_EQ",
                "notequal" => "TEXT_NOT_EQ",
                _ => null,
            };

            if (conditionType == null) {
                return false;
            }

            condition = new GoogleSheetsBooleanCondition {
                Type = conditionType,
                Values = new[] { value },
            };
            return true;
        }

        private static bool TryBuildNumericCondition(
            string filterOperator,
            string value,
            out GoogleSheetsBooleanCondition? condition) {
            condition = null;
            var normalizedOperator = filterOperator.Trim().ToLowerInvariant();
            string? conditionType = normalizedOperator switch {
                "equal" => "NUMBER_EQ",
                "notequal" => "NUMBER_NOT_EQ",
                "greaterthan" => "NUMBER_GREATER",
                "greaterthanorequal" => "NUMBER_GREATER_THAN_EQ",
                "lessthan" => "NUMBER_LESS",
                "lessthanorequal" => "NUMBER_LESS_THAN_EQ",
                _ => null,
            };

            if (conditionType == null) {
                return false;
            }

            condition = new GoogleSheetsBooleanCondition {
                Type = conditionType,
                Values = new[] { value },
            };
            return true;
        }

        private static bool TryNormalizeNumericCondition(
            ExcelCustomFilterConditionSnapshot condition,
            out string normalizedOperator,
            out string normalizedValue) {
            normalizedOperator = string.Empty;
            normalizedValue = string.Empty;

            if (condition == null || string.IsNullOrWhiteSpace(condition.Value)) {
                return false;
            }

            normalizedOperator = (condition.Operator ?? string.Empty).Trim().ToLowerInvariant();
            if (!double.TryParse(condition.Value, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, System.Globalization.CultureInfo.InvariantCulture, out var numericValue)) {
                return false;
            }

            normalizedValue = numericValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return normalizedOperator is "greaterthan" or "greaterthanorequal" or "lessthan" or "lessthanorequal";
        }

        private static bool TryGetBetweenBounds(
            string firstOperator,
            string firstValue,
            string secondOperator,
            string secondValue,
            out string lowerInclusive,
            out string upperInclusive) {
            lowerInclusive = string.Empty;
            upperInclusive = string.Empty;

            if (TryMatchLowerUpper(firstOperator, firstValue, secondOperator, secondValue, out lowerInclusive, out upperInclusive)) {
                return true;
            }

            return TryMatchLowerUpper(secondOperator, secondValue, firstOperator, firstValue, out lowerInclusive, out upperInclusive);
        }

        private static bool TryGetOutsideBounds(
            string firstOperator,
            string firstValue,
            string secondOperator,
            string secondValue,
            out string lowerExclusive,
            out string upperExclusive) {
            lowerExclusive = string.Empty;
            upperExclusive = string.Empty;

            if (TryMatchOutsideRange(firstOperator, firstValue, secondOperator, secondValue, out lowerExclusive, out upperExclusive)) {
                return true;
            }

            return TryMatchOutsideRange(secondOperator, secondValue, firstOperator, firstValue, out lowerExclusive, out upperExclusive);
        }

        private static bool TryMatchLowerUpper(
            string lowerOperator,
            string lowerValue,
            string upperOperator,
            string upperValue,
            out string lowerInclusive,
            out string upperInclusive) {
            lowerInclusive = string.Empty;
            upperInclusive = string.Empty;

            if (lowerOperator != "greaterthanorequal" || upperOperator != "lessthanorequal") {
                return false;
            }

            if (double.Parse(lowerValue, System.Globalization.CultureInfo.InvariantCulture) > double.Parse(upperValue, System.Globalization.CultureInfo.InvariantCulture)) {
                return false;
            }

            lowerInclusive = lowerValue;
            upperInclusive = upperValue;
            return true;
        }

        private static bool TryMatchOutsideRange(
            string lowerOperator,
            string lowerValue,
            string upperOperator,
            string upperValue,
            out string lowerExclusive,
            out string upperExclusive) {
            lowerExclusive = string.Empty;
            upperExclusive = string.Empty;

            if (lowerOperator != "lessthan" || upperOperator != "greaterthan") {
                return false;
            }

            if (double.Parse(lowerValue, System.Globalization.CultureInfo.InvariantCulture) > double.Parse(upperValue, System.Globalization.CultureInfo.InvariantCulture)) {
                return false;
            }

            lowerExclusive = lowerValue;
            upperExclusive = upperValue;
            return true;
        }

        private static void AddUnsupportedCustomFilterNotice(
            TranslationReport report,
            ref bool customFilterNoticeAdded) {
            if (customFilterNoticeAdded) {
                return;
            }

            report.Add(
                TranslationSeverity.Info,
                "CustomFilters",
                "Single-condition Excel custom filters are translated into native Google filter conditions when possible. More complex custom filter combinations are currently preserved as diagnostics only.");
            customFilterNoticeAdded = true;
        }

        private static IReadOnlyList<GoogleSheetsTableColumn> BuildTableColumns(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            var columns = new List<GoogleSheetsTableColumn>();
            foreach (var tableColumn in table.Columns) {
                var absoluteColumn = table.StartColumn + tableColumn.Index - 1;
                columns.Add(new GoogleSheetsTableColumn {
                    ColumnIndex = tableColumn.Index - 1,
                    Name = tableColumn.Name,
                    ColumnType = InferTableColumnType(workbookSnapshot, worksheet, table, absoluteColumn),
                    TotalsRowFunction = tableColumn.TotalsRowFunction,
                    DataValidationRule = BuildTableColumnValidationRule(workbookSnapshot, worksheet, table, absoluteColumn),
                });
            }

            return columns;
        }

        private static string? ResolveTableFooterColorArgb(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            if (!table.TotalsRowShown) {
                return null;
            }

            var footerColors = GetTableRowFillColors(worksheet, table, table.EndRow);

            if (footerColors.Count > 0) {
                return footerColors[0];
            }

            // A footer color is what prompts native Sheets table footer creation.
            return DefaultTableFooterColorArgb;
        }

        private static string? ResolveTableHeaderColorArgb(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            if (!table.HasHeaderRow) {
                return null;
            }

            var headerColors = GetTableRowFillColors(worksheet, table, table.StartRow);
            return headerColors.Count > 0 ? headerColors[0] : null;
        }

        private static string? ResolveTableFirstBandColorArgb(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            var firstDataRow = GetFirstDataRowIndex(table);
            if (!firstDataRow.HasValue) {
                return null;
            }

            var colors = GetTableRowFillColors(worksheet, table, firstDataRow.Value);
            return colors.Count > 0 ? colors[0] : null;
        }

        private static string? ResolveTableSecondBandColorArgb(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            var firstDataRow = GetFirstDataRowIndex(table);
            var lastDataRow = GetLastDataRowIndex(table);
            if (!firstDataRow.HasValue || !lastDataRow.HasValue) {
                return null;
            }

            var secondDataRow = firstDataRow.Value + 1;
            if (secondDataRow > lastDataRow.Value) {
                return null;
            }

            var colors = GetTableRowFillColors(worksheet, table, secondDataRow);
            return colors.Count > 0 ? colors[0] : null;
        }

        private static int? GetFirstDataRowIndex(ExcelTableSnapshot table) {
            var startRow = table.HasHeaderRow ? table.StartRow + 1 : table.StartRow;
            var lastDataRow = GetLastDataRowIndex(table);
            if (!lastDataRow.HasValue || startRow > lastDataRow.Value) {
                return null;
            }

            return startRow;
        }

        private static int? GetLastDataRowIndex(ExcelTableSnapshot table) {
            var endRow = table.TotalsRowShown ? table.EndRow - 1 : table.EndRow;
            if (endRow < table.StartRow) {
                return null;
            }

            return endRow;
        }

        private static List<string> GetTableRowFillColors(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table,
            int rowIndex) {
            return worksheet.Cells
                .Where(cell => cell.Row == rowIndex
                    && cell.Column >= table.StartColumn
                    && cell.Column <= table.EndColumn
                    && !string.IsNullOrWhiteSpace(cell.Style?.FillColorArgb))
                .Select(cell => cell.Style!.FillColorArgb!)
                .GroupBy(color => color, StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(group => group.Count())
                .ThenBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.Key)
                .ToList();
        }

        private static string InferTableColumnType(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table,
            int absoluteColumn) {
            var validationRule = BuildTableColumnValidationRule(workbookSnapshot, worksheet, table, absoluteColumn);
            if (validationRule != null) {
                return "DROPDOWN";
            }

            bool seenValue = false;
            bool allBoolean = true;
            bool allDateLike = true;
            bool allNumeric = true;
            bool anyPercent = false;
            bool anyCurrency = false;

            var startRow = table.HasHeaderRow ? table.StartRow + 1 : table.StartRow;
            var endRow = table.TotalsRowShown ? table.EndRow - 1 : table.EndRow;
            if (endRow < startRow) {
                return "TEXT";
            }

            foreach (var cell in worksheet.Cells.Where(c => c.Column == absoluteColumn && c.Row >= startRow && c.Row <= endRow)) {
                var value = cell.Value;
                if (value == null && string.IsNullOrWhiteSpace(cell.Formula)) {
                    continue;
                }

                seenValue = true;

                if (value is bool) {
                    allNumeric = false;
                    allDateLike = false;
                    continue;
                }

                allBoolean = false;

                if (value is DateTime || value is DateTimeOffset || cell.Style?.IsDateLike == true) {
                    allNumeric = false;
                    continue;
                }

                allDateLike = false;

                if (value is byte || value is sbyte || value is short || value is ushort
                    || value is int || value is uint || value is long || value is ulong
                    || value is float || value is double || value is decimal) {
                    var formatCode = cell.Style?.NumberFormatCode ?? string.Empty;
                    if (formatCode.IndexOf('%') >= 0) {
                        anyPercent = true;
                    }
                    if (formatCode.IndexOf('$') >= 0 || formatCode.IndexOf("z", StringComparison.OrdinalIgnoreCase) >= 0) {
                        anyCurrency = true;
                    }
                    continue;
                }

                allNumeric = false;
            }

            if (!seenValue) {
                return "TEXT";
            }
            if (allBoolean) {
                return "BOOLEAN";
            }
            if (allDateLike) {
                return "DATE_TIME";
            }
            if (allNumeric) {
                if (anyPercent) {
                    return "PERCENT";
                }
                if (anyCurrency) {
                    return "CURRENCY";
                }
                return "NUMBER";
            }

            return "TEXT";
        }

        private static GoogleSheetsDataValidationRule? BuildTableColumnValidationRule(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table,
            int absoluteColumn) {
            var validation = FindMatchingListValidation(worksheet, table, absoluteColumn);
            if (validation == null) {
                return null;
            }

            var values = ResolveListValidationValues(workbookSnapshot, worksheet, validation);
            if (values.Count == 0) {
                return null;
            }

            return new GoogleSheetsDataValidationRule {
                ConditionType = "ONE_OF_LIST",
                Values = values,
                Strict = true,
                ShowCustomUi = true,
            };
        }

        private static void AppendValidationOnlyCells(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            GoogleSheetsUpdateCellsRequest updateCells,
            ISet<string> emittedCellKeys,
            TranslationReport report,
            ref bool cellValidationNoticeAdded) {
            foreach (var validation in worksheet.Validations) {
                if (!IsSupportedDirectCellValidation(validation)) {
                    continue;
                }

                foreach (var (row, column) in EnumerateValidationCells(validation)) {
                    var cellKey = CreateCellKey(row, column);
                    if (emittedCellKeys.Contains(cellKey)) {
                        continue;
                    }

                    var validationRule = BuildDirectCellValidationRule(workbookSnapshot, worksheet, validation, row, column);
                    if (validationRule == null) {
                        continue;
                    }

                    if (!cellValidationNoticeAdded) {
                        report.Add(
                            TranslationSeverity.Info,
                            "CellValidations",
                            "List, whole-number, decimal, date, and text-length Excel data validations now compile into native Google Sheets cell validation rules for populated and empty target cells within explicit ranges.");
                        cellValidationNoticeAdded = true;
                    }

                    emittedCellKeys.Add(cellKey);
                    updateCells.AddCell(new GoogleSheetsCellData {
                        RowIndex = row - 1,
                        ColumnIndex = column - 1,
                        Value = GoogleSheetsCellValue.Blank(),
                        DataValidationRule = validationRule,
                    });
                }
            }
        }

        private static bool IsSupportedDirectCellValidation(ExcelDataValidationSnapshot validation) {
            return string.Equals(validation.Type, "list", StringComparison.OrdinalIgnoreCase)
                || string.Equals(validation.Type, "whole", StringComparison.OrdinalIgnoreCase)
                || string.Equals(validation.Type, "decimal", StringComparison.OrdinalIgnoreCase)
                || string.Equals(validation.Type, "date", StringComparison.OrdinalIgnoreCase)
                || string.Equals(validation.Type, "textlength", StringComparison.OrdinalIgnoreCase);
        }

        private static GoogleSheetsDataValidationRule? BuildCellValidationRule(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            int row,
            int column,
            TranslationReport report,
            ref bool cellValidationNoticeAdded) {
            foreach (var validation in worksheet.Validations) {
                if (!ValidationAppliesToCell(validation, row, column)) {
                    continue;
                }

                var rule = BuildDirectCellValidationRule(workbookSnapshot, worksheet, validation, row, column);
                if (rule == null) {
                    continue;
                }

                if (!cellValidationNoticeAdded) {
                    report.Add(
                        TranslationSeverity.Info,
                        "CellValidations",
                        "List, whole-number, decimal, date, and text-length Excel data validations now compile into native Google Sheets cell validation rules for populated and empty target cells within explicit ranges.");
                    cellValidationNoticeAdded = true;
                }

                return rule;
            }

            return null;
        }

        private static GoogleSheetsDataValidationRule? BuildDirectCellValidationRule(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            ExcelDataValidationSnapshot validation,
            int row,
            int column) {
            if (string.Equals(validation.Type, "list", StringComparison.OrdinalIgnoreCase)) {
                if (IsListValidationHandledByNativeTable(worksheet, validation, row, column)) {
                    return null;
                }

                var listValues = ResolveListValidationValues(workbookSnapshot, worksheet, validation);
                if (listValues.Count == 0) {
                    return null;
                }

                return new GoogleSheetsDataValidationRule {
                    ConditionType = "ONE_OF_LIST",
                    Values = listValues,
                    Strict = true,
                    ShowCustomUi = true,
                };
            }

            if (string.Equals(validation.Type, "whole", StringComparison.OrdinalIgnoreCase)
                || string.Equals(validation.Type, "decimal", StringComparison.OrdinalIgnoreCase)) {
                if (!TryMapNumericValidationConditionType(validation.Operator, out var numericConditionType, out var numericRequiresSecondValue)) {
                    return null;
                }

                if (!TryParseValidationNumber(validation.Formula1, out var firstNumberValue)) {
                    return null;
                }

                var numericValues = new List<string> { firstNumberValue };
                if (numericRequiresSecondValue) {
                    if (!TryParseValidationNumber(validation.Formula2, out var secondNumberValue)) {
                        return null;
                    }

                    numericValues.Add(secondNumberValue);
                }

                return new GoogleSheetsDataValidationRule {
                    ConditionType = numericConditionType,
                    Values = numericValues,
                    Strict = true,
                    ShowCustomUi = false,
                };
            }

            if (string.Equals(validation.Type, "date", StringComparison.OrdinalIgnoreCase)) {
                if (!TryMapDateValidationConditionType(validation.Operator, out var dateConditionType, out var dateRequiresSecondValue)) {
                    return null;
                }

                if (!TryParseValidationDate(validation.Formula1, out var firstDateValue)) {
                    return null;
                }

                var dateValues = new List<string> { firstDateValue };
                if (dateRequiresSecondValue) {
                    if (!TryParseValidationDate(validation.Formula2, out var secondDateValue)) {
                        return null;
                    }

                    dateValues.Add(secondDateValue);
                }

                return new GoogleSheetsDataValidationRule {
                    ConditionType = dateConditionType,
                    Values = dateValues,
                    Strict = true,
                    ShowCustomUi = false,
                };
            }

            if (string.Equals(validation.Type, "textlength", StringComparison.OrdinalIgnoreCase)) {
                if (!TryBuildTextLengthValidationFormula(validation, row, column, out var textLengthFormula)) {
                    return null;
                }

                return new GoogleSheetsDataValidationRule {
                    ConditionType = "CUSTOM_FORMULA",
                    Values = new[] { textLengthFormula },
                    Strict = true,
                    ShowCustomUi = false,
                };
            }

            return null;
        }

        private static bool IsListValidationHandledByNativeTable(
            ExcelWorksheetSnapshot worksheet,
            ExcelDataValidationSnapshot validation,
            int row,
            int column) {
            foreach (var table in worksheet.Tables) {
                var firstDataRow = GetFirstDataRowIndex(table);
                var lastDataRow = GetLastDataRowIndex(table);
                if (!firstDataRow.HasValue || !lastDataRow.HasValue) {
                    continue;
                }

                if (row < firstDataRow.Value || row > lastDataRow.Value) {
                    continue;
                }

                if (column < table.StartColumn || column > table.EndColumn) {
                    continue;
                }

                var matchingValidation = FindMatchingListValidation(worksheet, table, column);
                if (ReferenceEquals(matchingValidation, validation)) {
                    return true;
                }
            }

            return false;
        }

        private static IReadOnlyList<string> ResolveListValidationValues(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            ExcelDataValidationSnapshot validation) {
            var explicitValues = ParseExplicitListValidationValues(validation.Formula1);
            if (explicitValues.Count > 0) {
                return explicitValues;
            }

            var referencedRangeValues = ResolveReferencedRangeValidationValues(workbookSnapshot, worksheet.Name, validation.Formula1);
            if (referencedRangeValues.Count > 0) {
                return referencedRangeValues;
            }

            return ResolveNamedRangeValidationValues(workbookSnapshot, worksheet.Name, validation.Formula1);
        }

        private static ExcelDataValidationSnapshot? FindMatchingListValidation(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table,
            int absoluteColumn) {
            var firstDataRow = GetFirstDataRowIndex(table);
            var lastDataRow = GetLastDataRowIndex(table);
            if (!firstDataRow.HasValue || !lastDataRow.HasValue) {
                return null;
            }

            var columnAddress = A1.ColumnIndexToLetters(absoluteColumn);
            var expectedRange = $"{columnAddress}{firstDataRow.Value}:{columnAddress}{lastDataRow.Value}";
            return worksheet.Validations.FirstOrDefault(validation =>
                string.Equals(validation.Type, "list", StringComparison.OrdinalIgnoreCase)
                && validation.A1Ranges.Count == 1
                && string.Equals(validation.A1Ranges[0], expectedRange, StringComparison.OrdinalIgnoreCase));
        }

        private static IEnumerable<(int Row, int Column)> EnumerateValidationCells(ExcelDataValidationSnapshot validation) {
            foreach (var a1Range in validation.A1Ranges) {
                if (string.IsNullOrWhiteSpace(a1Range)) {
                    continue;
                }

                var normalizedRange = a1Range.Replace("$", string.Empty);
                int startRow;
                int startColumn;
                int endRow;
                int endColumn;

                if (!A1.TryParseRange(normalizedRange, out startRow, out startColumn, out endRow, out endColumn)) {
                    var (singleRow, singleColumn) = A1.ParseCellRef(normalizedRange);
                    if (singleRow <= 0 || singleColumn <= 0) {
                        continue;
                    }

                    startRow = endRow = singleRow;
                    startColumn = endColumn = singleColumn;
                }

                if (startRow <= 0 || startColumn <= 0 || endRow < startRow || endColumn < startColumn) {
                    continue;
                }

                for (var row = startRow; row <= endRow; row++) {
                    for (var column = startColumn; column <= endColumn; column++) {
                        yield return (row, column);
                    }
                }
            }
        }

        private static bool ValidationAppliesToCell(ExcelDataValidationSnapshot validation, int row, int column) {
            foreach (var a1Range in validation.A1Ranges) {
                if (TryRangeContainsCell(a1Range, row, column)) {
                    return true;
                }
            }

            return false;
        }

        private static string CreateCellKey(int row, int column) {
            return row.ToString(CultureInfo.InvariantCulture) + ":" + column.ToString(CultureInfo.InvariantCulture);
        }

        private static bool TryRangeContainsCell(string a1Range, int row, int column) {
            if (string.IsNullOrWhiteSpace(a1Range)) {
                return false;
            }

            var normalizedRange = a1Range.Replace("$", string.Empty);
            int startRow;
            int startColumn;
            int endRow;
            int endColumn;

            if (!A1.TryParseRange(normalizedRange, out startRow, out startColumn, out endRow, out endColumn)) {
                var (singleRow, singleColumn) = A1.ParseCellRef(normalizedRange);
                if (singleRow <= 0 || singleColumn <= 0) {
                    return false;
                }

                startRow = endRow = singleRow;
                startColumn = endColumn = singleColumn;
            }

            return row >= startRow
                && row <= endRow
                && column >= startColumn
                && column <= endColumn;
        }

        private static bool TryMapNumericValidationConditionType(
            string? validationOperator,
            out string conditionType,
            out bool requiresSecondValue) {
            requiresSecondValue = false;

            switch (validationOperator) {
                case "between":
                    conditionType = "NUMBER_BETWEEN";
                    requiresSecondValue = true;
                    return true;
                case "notBetween":
                    conditionType = "NUMBER_NOT_BETWEEN";
                    requiresSecondValue = true;
                    return true;
                case "equal":
                    conditionType = "NUMBER_EQ";
                    return true;
                case "notEqual":
                    conditionType = "NUMBER_NOT_EQ";
                    return true;
                case "greaterThan":
                    conditionType = "NUMBER_GREATER";
                    return true;
                case "greaterThanOrEqual":
                    conditionType = "NUMBER_GREATER_THAN_EQ";
                    return true;
                case "lessThan":
                    conditionType = "NUMBER_LESS";
                    return true;
                case "lessThanOrEqual":
                    conditionType = "NUMBER_LESS_THAN_EQ";
                    return true;
                default:
                    conditionType = string.Empty;
                    return false;
            }
        }

        private static bool TryMapDateValidationConditionType(
            string? validationOperator,
            out string conditionType,
            out bool requiresSecondValue) {
            requiresSecondValue = false;

            switch (validationOperator) {
                case "between":
                    conditionType = "DATE_BETWEEN";
                    requiresSecondValue = true;
                    return true;
                case "notBetween":
                    conditionType = "DATE_NOT_BETWEEN";
                    requiresSecondValue = true;
                    return true;
                case "equal":
                    conditionType = "DATE_EQ";
                    return true;
                case "greaterThan":
                    conditionType = "DATE_AFTER";
                    return true;
                case "greaterThanOrEqual":
                    conditionType = "DATE_ON_OR_AFTER";
                    return true;
                case "lessThan":
                    conditionType = "DATE_BEFORE";
                    return true;
                case "lessThanOrEqual":
                    conditionType = "DATE_ON_OR_BEFORE";
                    return true;
                default:
                    conditionType = string.Empty;
                    return false;
            }
        }

        private static bool TryParseValidationNumber(string? value, out string normalizedNumber) {
            normalizedNumber = string.Empty;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            if (!double.TryParse(value, NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var parsed)) {
                return false;
            }

            normalizedNumber = parsed.ToString("G15", CultureInfo.InvariantCulture);
            return true;
        }

        private static bool TryParseValidationDate(string? value, out string normalizedDate) {
            normalizedDate = string.Empty;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            if (!double.TryParse(value, NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var serialDate)) {
                return false;
            }

            try {
                normalizedDate = DateTime.FromOADate(serialDate).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

        private static bool TryBuildTextLengthValidationFormula(
            ExcelDataValidationSnapshot validation,
            int row,
            int column,
            out string formula) {
            formula = string.Empty;

            if (row <= 0 || column <= 0) {
                return false;
            }

            var cellReference = A1.ColumnIndexToLetters(column) + row.ToString(CultureInfo.InvariantCulture);
            var lengthExpression = $"LEN({cellReference})";

            if (!TryParseValidationNumber(validation.Formula1, out var firstValue)) {
                return false;
            }

            switch (validation.Operator) {
                case "equal":
                    formula = $"={lengthExpression}={firstValue}";
                    return true;
                case "notEqual":
                    formula = $"={lengthExpression}<>{firstValue}";
                    return true;
                case "greaterThan":
                    formula = $"={lengthExpression}>{firstValue}";
                    return true;
                case "greaterThanOrEqual":
                    formula = $"={lengthExpression}>={firstValue}";
                    return true;
                case "lessThan":
                    formula = $"={lengthExpression}<{firstValue}";
                    return true;
                case "lessThanOrEqual":
                    formula = $"={lengthExpression}<={firstValue}";
                    return true;
                case "between":
                    if (!TryParseValidationNumber(validation.Formula2, out var secondBetweenValue)) {
                        return false;
                    }

                    formula = $"=AND({lengthExpression}>={firstValue},{lengthExpression}<={secondBetweenValue})";
                    return true;
                case "notBetween":
                    if (!TryParseValidationNumber(validation.Formula2, out var secondNotBetweenValue)) {
                        return false;
                    }

                    formula = $"=OR({lengthExpression}<{firstValue},{lengthExpression}>{secondNotBetweenValue})";
                    return true;
                default:
                    return false;
            }
        }

        private static IReadOnlyList<string> ResolveReferencedRangeValidationValues(
            ExcelWorkbookSnapshot workbookSnapshot,
            string sourceSheetName,
            string? formula1) {
            if (string.IsNullOrWhiteSpace(formula1)) {
                return Array.Empty<string>();
            }

            var referenceText = formula1!.Trim();
            if (referenceText.StartsWith("=", StringComparison.Ordinal)) {
                referenceText = referenceText.Substring(1).Trim();
            }

            if (string.IsNullOrWhiteSpace(referenceText)) {
                return Array.Empty<string>();
            }

            var targetSheetName = sourceSheetName;
            if (TrySplitSheetQualifiedRange(referenceText, out var explicitSheetName, out var unqualifiedRange)) {
                targetSheetName = explicitSheetName!;
                referenceText = unqualifiedRange;
            }

            return ResolveWorksheetRangeValidationValues(workbookSnapshot, targetSheetName, referenceText);
        }

        private static IReadOnlyList<string> ResolveNamedRangeValidationValues(
            ExcelWorkbookSnapshot workbookSnapshot,
            string sourceSheetName,
            string? formula1) {
            if (string.IsNullOrWhiteSpace(formula1)) {
                return Array.Empty<string>();
            }

            var namedRangeName = formula1!.Trim();
            if (namedRangeName.StartsWith("=", StringComparison.Ordinal)) {
                namedRangeName = namedRangeName.Substring(1).Trim();
            }

            if (string.IsNullOrWhiteSpace(namedRangeName)) {
                return Array.Empty<string>();
            }

            var namedRange = workbookSnapshot.NamedRanges.FirstOrDefault(range =>
                string.Equals(range.Name, namedRangeName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(range.SheetName, sourceSheetName, StringComparison.OrdinalIgnoreCase))
                ?? workbookSnapshot.NamedRanges.FirstOrDefault(range =>
                    string.Equals(range.Name, namedRangeName, StringComparison.OrdinalIgnoreCase)
                    && string.IsNullOrWhiteSpace(range.SheetName));

            if (namedRange == null) {
                return Array.Empty<string>();
            }

            var rangeText = namedRange.ReferenceA1.Replace("$", string.Empty);
            var targetSheetName = namedRange.SheetName;
            if (TrySplitSheetQualifiedRange(rangeText, out var explicitSheetName, out var unqualifiedRange)) {
                targetSheetName = explicitSheetName;
                rangeText = unqualifiedRange;
            }

            if (string.IsNullOrWhiteSpace(targetSheetName)) {
                return Array.Empty<string>();
            }

            return ResolveWorksheetRangeValidationValues(workbookSnapshot, targetSheetName!, rangeText);
        }

        private static IReadOnlyList<string> ResolveWorksheetRangeValidationValues(
            ExcelWorkbookSnapshot workbookSnapshot,
            string targetSheetName,
            string rangeText) {
            if (string.IsNullOrWhiteSpace(targetSheetName) || string.IsNullOrWhiteSpace(rangeText)) {
                return Array.Empty<string>();
            }

            var targetWorksheet = workbookSnapshot.Worksheets.FirstOrDefault(worksheet =>
                string.Equals(worksheet.Name, targetSheetName, StringComparison.OrdinalIgnoreCase));
            if (targetWorksheet == null) {
                return Array.Empty<string>();
            }

            int startRow;
            int startColumn;
            int endRow;
            int endColumn;
            if (!A1.TryParseRange(rangeText, out startRow, out startColumn, out endRow, out endColumn)) {
                var (row, column) = A1.ParseCellRef(rangeText);
                if (row <= 0 || column <= 0) {
                    return Array.Empty<string>();
                }

                startRow = endRow = row;
                startColumn = endColumn = column;
            }

            if (startRow != endRow && startColumn != endColumn) {
                return Array.Empty<string>();
            }

            return targetWorksheet.Cells
                .Where(cell => cell.Row >= startRow
                    && cell.Row <= endRow
                    && cell.Column >= startColumn
                    && cell.Column <= endColumn)
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .Select(cell => ConvertCellValueToValidationItem(cell.Value))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!)
                .ToList();
        }

        private static IReadOnlyList<string> ParseExplicitListValidationValues(string? formula1) {
            if (string.IsNullOrWhiteSpace(formula1) || formula1!.Length < 2 || formula1[0] != '"' || formula1[formula1.Length - 1] != '"') {
                return Array.Empty<string>();
            }

            var values = new List<string>();
            var current = new System.Text.StringBuilder();
            var inner = formula1.Substring(1, formula1.Length - 2);

            for (int index = 0; index < inner.Length; index++) {
                var character = inner[index];
                if (character == '"'
                    && index + 1 < inner.Length
                    && inner[index + 1] == '"') {
                    current.Append('"');
                    index++;
                    continue;
                }

                if (character == ',') {
                    values.Add(current.ToString());
                    current.Clear();
                    continue;
                }

                current.Append(character);
            }

            values.Add(current.ToString());
            return values
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .ToList();
        }

        private static string? ConvertCellValueToValidationItem(object? value) {
            if (value == null) {
                return null;
            }

            return value switch {
                string text => text,
                bool boolean => boolean ? "TRUE" : "FALSE",
                DateTime dateTime => dateTime.ToString("O", CultureInfo.InvariantCulture),
                DateTimeOffset dateTimeOffset => dateTimeOffset.ToString("O", CultureInfo.InvariantCulture),
                IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
                _ => Convert.ToString(value, CultureInfo.InvariantCulture),
            };
        }

        private static bool TrySplitSheetQualifiedRange(string value, out string? sheetName, out string unqualifiedRange) {
            sheetName = null;
            unqualifiedRange = value;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            var bangIndex = value.LastIndexOf('!');
            if (bangIndex <= 0 || bangIndex >= value.Length - 1) {
                return false;
            }

            var sheetPart = value.Substring(0, bangIndex).Trim();
            var rangePart = value.Substring(bangIndex + 1).Trim();
            if (sheetPart.Length >= 2 && sheetPart[0] == '\'' && sheetPart[sheetPart.Length - 1] == '\'') {
                sheetPart = sheetPart.Substring(1, sheetPart.Length - 2).Replace("''", "'");
            }

            sheetName = sheetPart;
            unqualifiedRange = rangePart.Replace("$", string.Empty);
            return true;
        }

        private static string ConvertCellToFilterText(ExcelCellSnapshot cell) {
            if (cell.Value == null) {
                return string.Empty;
            }

            return cell.Value switch {
                DateTime dateTime => dateTime.ToString("o", System.Globalization.CultureInfo.InvariantCulture),
                DateTimeOffset dateTimeOffset => dateTimeOffset.ToString("o", System.Globalization.CultureInfo.InvariantCulture),
                bool booleanValue => booleanValue ? "TRUE" : "FALSE",
                _ => Convert.ToString(cell.Value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
            };
        }

        private static long GetWorksheetCellKey(int row, int column) {
            return ((long)row << 20) | (uint)column;
        }

        private static int ConvertExcelColumnWidthToPixels(double widthUnits) {
            const double mdw = 7.0;
            var pixels = Math.Truncate((256.0 * widthUnits + Math.Truncate(128.0 / mdw)) / 256.0 * mdw);
            return Math.Max(0, (int)Math.Round(pixels));
        }

        private static int ConvertPointsToPixels(double points) {
            return Math.Max(0, (int)Math.Round(points * 96.0 / 72.0));
        }
    }
}
