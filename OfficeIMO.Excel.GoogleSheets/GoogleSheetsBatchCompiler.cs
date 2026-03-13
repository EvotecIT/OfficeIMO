using OfficeIMO.GoogleWorkspace;
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
                        FooterColorArgb = ResolveTableFooterColorArgb(worksheet, table),
                        Columns = BuildTableColumns(worksheet, table),
                    });
                }

                var filterRequests = BuildFilterRequests(worksheet, report, ref multipleFilterNoticeAdded, ref customFilterNoticeAdded);
                foreach (var filterRequest in filterRequests) {
                    batch.Add(filterRequest);
                }

                var updateCells = new GoogleSheetsUpdateCellsRequest {
                    SheetName = worksheet.Name
                };

                foreach (var cell in worksheet.Cells) {
                    if (!styleNoticeAdded && cell.Style != null) {
                        report.Add(
                            TranslationSeverity.Info,
                            "Styles",
                            "The current Google Sheets compiler now emits basic cell styling, hyperlinks, and row/column dimensions alongside workbook structure, values, formulas, frozen panes, merges, and named ranges.");
                        styleNoticeAdded = true;
                    }

                    var cellValue = BuildCellValue(cell, options, report, ref formulaNoticeAdded);
                    updateCells.AddCell(new GoogleSheetsCellData {
                        RowIndex = cell.Row - 1,
                        ColumnIndex = cell.Column - 1,
                        Value = cellValue,
                        NumberFormatHint = GetNumberFormatHint(cell.Value, cell.Style),
                        Style = BuildCellStyle(cell.Style),
                        Hyperlink = BuildHyperlink(cell.Hyperlink),
                        Comment = BuildComment(cell.Comment),
                    });
                }

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
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            var columns = new List<GoogleSheetsTableColumn>();
            foreach (var tableColumn in table.Columns) {
                var absoluteColumn = table.StartColumn + tableColumn.Index - 1;
                columns.Add(new GoogleSheetsTableColumn {
                    ColumnIndex = tableColumn.Index - 1,
                    Name = tableColumn.Name,
                    ColumnType = InferTableColumnType(worksheet, table, absoluteColumn),
                    TotalsRowFunction = tableColumn.TotalsRowFunction,
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

            var footerColors = worksheet.Cells
                .Where(cell => cell.Row == table.EndRow
                    && cell.Column >= table.StartColumn
                    && cell.Column <= table.EndColumn
                    && !string.IsNullOrWhiteSpace(cell.Style?.FillColorArgb))
                .Select(cell => cell.Style!.FillColorArgb!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (footerColors.Count > 0) {
                return footerColors[0];
            }

            // A footer color is what prompts native Sheets table footer creation.
            return DefaultTableFooterColorArgb;
        }

        private static string InferTableColumnType(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table,
            int absoluteColumn) {
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
