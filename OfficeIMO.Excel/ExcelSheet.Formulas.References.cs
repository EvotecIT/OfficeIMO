using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static bool TryParseCellReference(string reference, out int row, out int column) {
            row = 0;
            column = 0;
            try {
                var cellRef = A1.ParseCellRef(reference.Replace("$", string.Empty));
                row = cellRef.Row;
                column = cellRef.Col;
                return row > 0 && column > 0 && row <= A1.MaxRows && column <= A1.MaxColumns;
            } catch (ArgumentException) {
                return false;
            }
        }

        private bool TryResolveNumericOperand(string token, out double value) {
            token = token.Trim();
            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) {
                return true;
            }

            if (!TryParseQualifiedFormulaCellReference(token, out ExcelSheet sheet, out int row, out int column)
                && (!TryResolveFormulaRangeReference(token, out sheet, out row, out column, out int r2, out int c2)
                    || row != r2
                    || column != c2)) {
                return false;
            }

            var argument = ResolveFormulaDependency(sheet, row, column);
            if (argument.Number.HasValue) {
                value = argument.Number.Value;
                return true;
            }

            return false;
        }

        private bool TryParseQualifiedFormulaCellReference(string token, out ExcelSheet sheet, out int row, out int column) {
            return TryParseQualifiedFormulaCellReference(token, null, out sheet, out row, out column);
        }

        private bool TryResolveFormulaReferenceArgument(string token, out ExcelSheet sheet, out int row, out int column) {
            if (TryParseQualifiedFormulaCellReference(token, out sheet, out row, out column)) {
                return true;
            }

            if (TryResolveFormulaRangeReference(token, out sheet, out row, out column, out int endRow, out int endColumn)
                && row == endRow
                && column == endColumn) {
                return true;
            }

            sheet = this;
            row = 0;
            column = 0;
            return false;
        }

        private bool TryEvaluateSingleReferenceFormulaValue(string formula, out FormulaArgumentValue value) {
            string token = formula.Trim();
            if (token.StartsWith("=", StringComparison.Ordinal)) {
                token = token.Substring(1).Trim();
            }

            if (!TryResolveFormulaReferenceArgument(token, out ExcelSheet sheet, out int row, out int column)) {
                value = default;
                return false;
            }

            value = ResolveFormulaDependency(sheet, row, column);
            return value.HasValue && !value.IsUnresolvedFormula;
        }

        private bool TryParseQualifiedFormulaCellReference(string token, ExcelSheet? defaultSheet, out ExcelSheet sheet, out int row, out int column) {
            sheet = this;
            row = 0;
            column = 0;

            if (!TrySplitQualifiedReference(token, out string? sheetName, out string reference)) {
                return false;
            }

            if (sheetName != null) {
                if (!TryGetFormulaReferenceSheet(sheetName, out sheet)) {
                    return false;
                }
            } else if (defaultSheet != null) {
                sheet = defaultSheet;
            }

            var cellRef = A1.ParseCellRef(reference.Replace("$", string.Empty));
            row = cellRef.Row;
            column = cellRef.Col;
            return row > 0
                && column > 0
                && row <= A1.MaxRows
                && column <= A1.MaxColumns;
        }

        private bool TryParseQualifiedFormulaRange(string token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
            return TryParseQualifiedFormulaRange(token, null, out sheet, out r1, out c1, out r2, out c2);
        }

        private bool TryParseQualifiedFormulaRange(string token, ExcelSheet? defaultSheet, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
            sheet = this;
            r1 = 0;
            c1 = 0;
            r2 = 0;
            c2 = 0;

            if (!TrySplitQualifiedReference(token, out string? sheetName, out string reference)) {
                return false;
            }

            if (sheetName != null) {
                if (!TryGetFormulaReferenceSheet(sheetName, out sheet)) {
                    return false;
                }
            } else if (defaultSheet != null) {
                sheet = defaultSheet;
            }

            return A1.TryParseRange(reference.Replace("$", string.Empty), out r1, out c1, out r2, out c2);
        }

        private bool TryParseQualifiedFormulaWholeRange(
            string token,
            ExcelSheet? defaultSheet,
            out ExcelSheet sheet,
            out int r1,
            out int c1,
            out int r2,
            out int c2,
            out string address) {
            sheet = this;
            r1 = c1 = r2 = c2 = 0;
            address = string.Empty;
            if (!TrySplitQualifiedReference(token, out string? sheetName, out string reference)) {
                return false;
            }

            if (sheetName != null) {
                if (!TryGetFormulaReferenceSheet(sheetName, out sheet)) {
                    return false;
                }
            } else if (defaultSheet != null) {
                sheet = defaultSheet;
            }

            if (A1.TryParseWholeColumnRange(reference, out c1, out c2)) {
                r1 = 1;
                r2 = A1.MaxRows;
                address = A1.ColumnIndexToLetters(c1) + ":" + A1.ColumnIndexToLetters(c2);
                return true;
            }

            if (A1.TryParseWholeRowRange(reference, out r1, out r2)) {
                c1 = 1;
                c2 = A1.MaxColumns;
                address = r1.ToString(CultureInfo.InvariantCulture) + ":" + r2.ToString(CultureInfo.InvariantCulture);
                return true;
            }

            return false;
        }

        private bool TryResolveFormulaRangeReference(string token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
            int? currentRow = null;
            if (_formulaEvaluationCellReference != null
                && TryParseCellReference(_formulaEvaluationCellReference, out int evaluationRow, out _)) {
                currentRow = evaluationRow;
            }

            return TryResolveFormulaRangeReference(token, currentRow, out sheet, out r1, out c1, out r2, out c2);
        }

        private bool TryResolveFormulaRangeReference(
            string token,
            int? currentRow,
            out ExcelSheet sheet,
            out int r1,
            out int c1,
            out int r2,
            out int c2) {
            if (ExcelFormulaExpressionParser.TryParseSupportedFunctionCall(token, out ExcelFormulaFunctionCallSyntax? call)
                && call!.Name.Equals("OFFSET", StringComparison.OrdinalIgnoreCase)) {
                bool resolved = TryResolveOffsetRange(call.Arguments, currentRow, out sheet, out r1, out c1, out r2, out c2,
                    out FormulaArgumentValue error);
                if (resolved && error.IsError && !_formulaReferenceError.IsError) _formulaReferenceError = error;
                return resolved && !error.IsError;
            }
            if (TryParseQualifiedFormulaRange(token, out sheet, out r1, out c1, out r2, out c2)) {
                return true;
            }

            if (TryParseQualifiedFormulaCellReference(token, out sheet, out r1, out c1)) {
                r2 = r1;
                c2 = c1;
                return true;
            }

            if (TryResolveTableReferenceRange(token, currentRow, out sheet, out r1, out c1, out r2, out c2)) {
                return true;
            }

            if (TryResolveUnqualifiedCurrentRowTableReferenceRange(token, currentRow, out sheet, out r1, out c1, out r2, out c2)) {
                return true;
            }

            return TryResolveDefinedNameRange(token, currentRow, out sheet, out r1, out c1, out r2, out c2);
        }

        private bool TryResolveTableReferenceRange(
            string token,
            int? currentRow,
            out ExcelSheet sheet,
            out int r1,
            out int c1,
            out int r2,
            out int c2) {
            sheet = this;
            r1 = 0;
            c1 = 0;
            r2 = 0;
            c2 = 0;

            if (!TryParseStructuredTableReference(token, out string tableName, out var sections)) {
                return false;
            }

            WorkbookPart? workbookPart = _spreadSheetDocument.WorkbookPart;
            if (workbookPart == null) {
                return false;
            }

            foreach (var sheetElement in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (sheetElement.Id?.Value == null) {
                    continue;
                }

                if (workbookPart.GetPartById(sheetElement.Id.Value) is not WorksheetPart worksheetPart) {
                    continue;
                }

                foreach (var tablePart in worksheetPart.TableDefinitionParts) {
                    Table? table = tablePart.Table;
                    if (table == null
                        || (!string.Equals(table.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(table.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase))) {
                        continue;
                    }

                    sheet = string.Equals(Name, sheetElement.Name?.Value, StringComparison.OrdinalIgnoreCase)
                        ? this
                        : new ExcelSheet(_excelDocument, _spreadSheetDocument, sheetElement) {
                            _formulaEvaluationCache = _formulaEvaluationCache,
                            _formulaEvaluationDepthCache = _formulaEvaluationDepthCache,
                            _formulaEvaluationStack = _formulaEvaluationStack,
                            _formulaEvaluationDepthFrames = _formulaEvaluationDepthFrames,
                            _formulaEvaluationGuardState = _formulaEvaluationGuardState
                        };
                    return TryResolveTableReferenceRange(table, sections, currentRow, out r1, out c1, out r2, out c2);
                }
            }

            return false;
        }

        private readonly struct FormulaStructuredTableReference {
            internal FormulaStructuredTableReference(
                string area,
                bool areaIsExplicit,
                string? firstColumn,
                string? lastColumn,
                string? additionalArea = null) {
                Area = area;
                AreaIsExplicit = areaIsExplicit;
                FirstColumn = firstColumn;
                LastColumn = lastColumn;
                AdditionalArea = additionalArea;
            }

            internal string Area { get; }
            internal bool AreaIsExplicit { get; }
            internal string? FirstColumn { get; }
            internal string? LastColumn { get; }
            internal string? AdditionalArea { get; }

            internal FormulaStructuredTableReference WithArea(string area) {
                return new FormulaStructuredTableReference(area, true, FirstColumn, LastColumn, AdditionalArea);
            }
        }

        private static bool TryParseStructuredTableReference(
            string token,
            out string tableName,
            out FormulaStructuredTableReference reference) {
            string value = token.Trim();
            tableName = string.Empty;
            reference = default;
            if (value.Length == 0 || value.IndexOf('!') >= 0) {
                return false;
            }

            int bracketStart = value.IndexOf('[');
            tableName = bracketStart < 0 ? value : value.Substring(0, bracketStart);
            if (!IsFormulaDefinedNameToken(tableName)) {
                return false;
            }

            if (bracketStart < 0) {
                reference = new FormulaStructuredTableReference("#Data", false, null, null);
                return true;
            }

            string specifier = value.Substring(bracketStart);
            if (specifier.Length < 2 || specifier[0] != '[' || specifier[specifier.Length - 1] != ']') {
                return false;
            }

            string content = specifier.Substring(1, specifier.Length - 2).Trim();
            if (content.Length == 0) {
                return false;
            }

            if (content[0] == '@') {
                return TryParseStructuredCurrentRowColumns(content.Substring(1).Trim(), out reference);
            }

            if (content[0] != '[') {
                if (ContainsUnescapedStructuredReferenceBracket(content)) {
                    return false;
                }

                reference = IsStructuredTableAreaSpecifier(content)
                    ? new FormulaStructuredTableReference(content, true, null, null)
                    : new FormulaStructuredTableReference("#Data", false, content, content);
                return true;
            }

            if (!TryParseStructuredTableSectionSequence(content, out List<string> sections, out List<char> separators)) {
                return false;
            }

            if (sections.Count == 1) {
                string section = sections[0];
                reference = IsStructuredTableAreaSpecifier(section)
                    ? new FormulaStructuredTableReference(section, true, null, null)
                    : new FormulaStructuredTableReference("#Data", false, section, section);
                return true;
            }

            if (!IsStructuredTableAreaSpecifier(sections[0])) {
                if (sections.Count != 2 || separators[0] != ':') {
                    return false;
                }

                reference = new FormulaStructuredTableReference("#Data", false, sections[0], sections[1]);
                return true;
            }

            string area = sections[0];
            string? additionalArea = null;
            int areaCount = 1;
            if (sections.Count > 1 && IsStructuredTableAreaSpecifier(sections[1])) {
                if (separators[0] != ',' || !IsSupportedStructuredTableAreaCombination(area, sections[1])) {
                    return false;
                }

                additionalArea = sections[1];
                areaCount = 2;
            }

            int columnCount = sections.Count - areaCount;
            if (columnCount == 0) {
                reference = new FormulaStructuredTableReference(area, true, null, null, additionalArea);
                return true;
            }

            if (separators[areaCount - 1] != ',') {
                return false;
            }

            string firstColumn = sections[areaCount];
            if (columnCount == 1) {
                reference = new FormulaStructuredTableReference(area, true, firstColumn, firstColumn, additionalArea);
                return true;
            }

            if (columnCount == 2 && separators[areaCount] == ':') {
                reference = new FormulaStructuredTableReference(
                    area,
                    true,
                    firstColumn,
                    sections[areaCount + 1],
                    additionalArea);
                return true;
            }

            return false;
        }

        private static bool IsSupportedStructuredTableAreaCombination(string first, string second) {
            return (string.Equals(first, "#Headers", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(second, "#Data", StringComparison.OrdinalIgnoreCase))
                || (string.Equals(first, "#Data", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(second, "#Totals", StringComparison.OrdinalIgnoreCase));
        }

        private static bool TryParseStructuredCurrentRowColumns(
            string value,
            out FormulaStructuredTableReference reference) {
            reference = default;
            if (value.Length == 0) {
                return false;
            }

            if (value[0] != '[') {
                if (ContainsUnescapedStructuredReferenceBracket(value)) {
                    return false;
                }

                reference = new FormulaStructuredTableReference("#This Row", true, value, value);
                return true;
            }

            if (!TryParseStructuredTableSectionSequence(value, out List<string> sections, out List<char> separators)) {
                return false;
            }

            if (sections.Count == 1) {
                reference = new FormulaStructuredTableReference("#This Row", true, sections[0], sections[0]);
                return true;
            }

            if (sections.Count == 2 && separators[0] == ':') {
                reference = new FormulaStructuredTableReference("#This Row", true, sections[0], sections[1]);
                return true;
            }

            return false;
        }

        private static bool TryParseStructuredTableSectionSequence(
            string value,
            out List<string> sections,
            out List<char> separators) {
            sections = new List<string>();
            separators = new List<char>();
            int index = 0;
            while (index < value.Length) {
                while (index < value.Length && char.IsWhiteSpace(value[index])) {
                    index++;
                }
                if (index == value.Length) {
                    return sections.Count > 0 && separators.Count == sections.Count - 1;
                }

                if (value[index] != '[') {
                    return false;
                }

                int end = FindUnescapedStructuredReferenceBracket(value, index + 1, ']');
                if (end < 0) {
                    return false;
                }

                string section = value.Substring(index + 1, end - index - 1).Trim();
                if (section.Length == 0 || ContainsUnescapedStructuredReferenceBracket(section)) {
                    return false;
                }

                sections.Add(section);
                index = end + 1;
                while (index < value.Length && char.IsWhiteSpace(value[index])) {
                    index++;
                }
                if (index == value.Length) {
                    return separators.Count == sections.Count - 1;
                }

                if (value[index] != ',' && value[index] != ':') {
                    return false;
                }

                separators.Add(value[index]);
                index++;
            }

            return sections.Count > 0 && separators.Count == sections.Count - 1;
        }

        private static bool ContainsUnescapedStructuredReferenceBracket(string value) {
            return FindUnescapedStructuredReferenceBracket(value, 0, '[') >= 0
                || FindUnescapedStructuredReferenceBracket(value, 0, ']') >= 0;
        }

        private static int FindUnescapedStructuredReferenceBracket(string value, int start, char bracket) {
            for (int index = start; index < value.Length; index++) {
                if (value[index] == bracket && !IsEscapedStructuredReferenceBracket(value, index)) {
                    return index;
                }
            }
            return -1;
        }

        private static bool TryResolveTableReferenceRange(
            Table table,
            FormulaStructuredTableReference reference,
            int? currentRow,
            out int r1,
            out int c1,
            out int r2,
            out int c2) {
            r1 = 0;
            c1 = 0;
            r2 = 0;
            c2 = 0;

            if (table.Reference?.Value == null
                || !A1.TryParseRange(table.Reference.Value.Replace("$", string.Empty), out int tableR1, out int tableC1, out int tableR2, out int tableC2)) {
                return false;
            }

            uint headerRows = table.HeaderRowCount?.Value ?? 1U;
            uint totalsRows = table.TotalsRowShown?.Value == true
                ? Math.Max(1U, table.TotalsRowCount?.Value ?? 1U)
                : 0U;

            if (!TryResolveTableAreaRows(reference.Area, tableR1, tableR2, headerRows, totalsRows, currentRow, out r1, out r2)) {
                return false;
            }
            if (reference.AdditionalArea != null) {
                if (!TryResolveTableAreaRows(
                    reference.AdditionalArea,
                    tableR1,
                    tableR2,
                    headerRows,
                    totalsRows,
                    currentRow,
                    out int additionalR1,
                    out int additionalR2)
                    || additionalR1 > r2 + 1) {
                    return false;
                }

                r1 = Math.Min(r1, additionalR1);
                r2 = Math.Max(r2, additionalR2);
            }

            c1 = tableC1;
            c2 = tableC2;
            if (!string.IsNullOrWhiteSpace(reference.FirstColumn)) {
                int firstOffset = ResolveTableColumnOffset(table, reference.FirstColumn!);
                int lastOffset = ResolveTableColumnOffset(table, reference.LastColumn ?? reference.FirstColumn!);
                if (firstOffset < 0 || lastOffset < firstOffset) {
                    return false;
                }

                c1 = tableC1 + firstOffset;
                c2 = tableC1 + lastOffset;
            }

            return r1 <= r2 && c1 <= c2;
        }

        private static bool TryResolveTableAreaRows(
            string area,
            int tableR1,
            int tableR2,
            uint headerRows,
            uint totalsRows,
            int? currentRow,
            out int r1,
            out int r2) {
            r1 = tableR1;
            r2 = tableR2;
            if (string.Equals(area, "#All", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (string.Equals(area, "#Headers", StringComparison.OrdinalIgnoreCase)) {
                if (headerRows == 0) {
                    return false;
                }

                r2 = tableR1 + (int)headerRows - 1;
                return r2 <= tableR2;
            }

            if (string.Equals(area, "#Totals", StringComparison.OrdinalIgnoreCase)) {
                if (totalsRows == 0) {
                    return false;
                }

                r1 = tableR2 - (int)totalsRows + 1;
                return r1 >= tableR1;
            }

            if (string.Equals(area, "#This Row", StringComparison.OrdinalIgnoreCase)) {
                int dataR1 = tableR1 + (int)headerRows;
                int dataR2 = tableR2 - (int)totalsRows;
                if (!currentRow.HasValue || currentRow.Value < dataR1 || currentRow.Value > dataR2) {
                    return false;
                }

                r1 = r2 = currentRow.Value;
                return true;
            }

            if (!string.Equals(area, "#Data", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            r1 = tableR1 + (int)headerRows;
            r2 = tableR2 - (int)totalsRows;
            return r1 <= r2;
        }

        private static bool IsStructuredTableAreaSpecifier(string section) {
            return string.Equals(section, "#All", StringComparison.OrdinalIgnoreCase)
                || string.Equals(section, "#Data", StringComparison.OrdinalIgnoreCase)
                || string.Equals(section, "#Headers", StringComparison.OrdinalIgnoreCase)
                || string.Equals(section, "#Totals", StringComparison.OrdinalIgnoreCase)
                || string.Equals(section, "#This Row", StringComparison.OrdinalIgnoreCase);
        }

        private static int ResolveTableColumnOffset(Table table, string columnName) {
            string decodedColumnName = ExcelFormulaSyntaxTree.DecodeStructuredColumnName(columnName);
            int index = 0;
            foreach (var tableColumn in table.TableColumns?.Elements<TableColumn>() ?? Enumerable.Empty<TableColumn>()) {
                if (string.Equals(tableColumn.Name?.Value, decodedColumnName, StringComparison.OrdinalIgnoreCase)) {
                    return index;
                }

                index++;
            }

            return -1;
        }

        private bool TryGetFormulaReferenceSheet(string sheetName, out ExcelSheet sheet) {
            sheet = this;
            if (string.Equals(Name, sheetName, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            var sheetElement = WorkbookRoot.Sheets?
                .Elements<Sheet>()
                .FirstOrDefault(candidate => string.Equals(candidate.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheetElement?.Id == null) {
                return false;
            }

            sheet = new ExcelSheet(_excelDocument, _spreadSheetDocument, sheetElement) {
                _formulaEvaluationCache = _formulaEvaluationCache,
                _formulaEvaluationDepthCache = _formulaEvaluationDepthCache,
                _formulaEvaluationStack = _formulaEvaluationStack,
                _formulaEvaluationDepthFrames = _formulaEvaluationDepthFrames,
                _formulaEvaluationGuardState = _formulaEvaluationGuardState
            };
            return true;
        }

        private static bool TrySplitQualifiedReference(string token, out string? sheetName, out string reference) {
            string value = token.Trim();
            sheetName = null;
            reference = value;
            if (value.Length == 0) {
                return false;
            }

            int separator = value.LastIndexOf('!');
            if (separator < 0) {
                return true;
            }

            if (separator == 0 || separator == value.Length - 1) {
                return false;
            }

            sheetName = NormalizeFormulaSheetName(value.Substring(0, separator));
            reference = value.Substring(separator + 1).Trim();
            return !string.IsNullOrWhiteSpace(sheetName) && reference.Length > 0;
        }

        private static string NormalizeFormulaSheetName(string token) {
            string value = token.Trim();
            if (value.Length >= 2 && value[0] == '\'' && value[value.Length - 1] == '\'') {
                value = value.Substring(1, value.Length - 2).Replace("''", "'");
            }

            return value;
        }

        private FormulaArgumentValue ResolveFormulaDependency(ExcelSheet sheet, int row, int column) {
            FormulaArgumentValue value = sheet.ResolveCellArgument(row, column);
            if (value.IsUnevaluatedFormulaCache && _formulaEvaluationDepthFrames?.Count > 0)
                _formulaEvaluationDepthFrames.Peek().MarkUnevaluatedFormulaCache();
            return value;
        }

        private FormulaArgumentValue ResolveCellArgument(int row, int column) {
            var cell = TryGetExistingCell(row, column);
            bool unresolvedFormula = false;
            if (cell?.CellFormula != null && _formulaEvaluationCache != null) {
                if (TryEvaluateFormulaCellValue(cell, out FormulaArgumentValue formulaResult)) {
                    return formulaResult.AsReferencedValue();
                }

                if (_formulaEvaluationDepthFrames != null
                    && _formulaEvaluationDepthFrames.Count > 0
                    && _formulaEvaluationDepthFrames.Peek().DependencyGuardBlocked) {
                    return FormulaArgumentValue.UnresolvedFormula();
                }

                unresolvedFormula = true;
            }

            bool unevaluatedFormulaCache = unresolvedFormula
                || (cell?.CellFormula != null && _formulaEvaluationCache == null);
            return ResolveFormulaArgumentSnapshot(row, column, unresolvedFormula,
                cell?.CellFormula == null, unevaluatedFormulaCache);
        }

        // Keep snapshot decoding and array-child lookup out of the recursive
        // formula dependency frame. Their locals are needed only at a leaf.
        private FormulaArgumentValue ResolveFormulaArgumentSnapshot(int row, int column, bool unresolvedFormula,
            bool withoutFormula, bool unevaluatedFormulaCache) {
            if (withoutFormula && _formulaEvaluationCache != null
                && TryResolveFixedArrayChild(row, column, out FormulaArgumentValue arrayValue)) return arrayValue.AsReferencedValue();
            var value = GetCellValueSnapshot(row, column);
            if (unresolvedFormula && value.Value == null && string.IsNullOrEmpty(value.CachedText)) {
                return FormulaArgumentValue.UnresolvedFormula();
            }

            if (unresolvedFormula
                && _formulaEvaluationDepthFrames != null
                && _formulaEvaluationDepthFrames.Count > 0) {
                _formulaEvaluationDepthFrames.Peek().IncludeChild(1);
            }

            var cachedType = value.HasFormula ? TryGetExistingCell(row, column)?.DataType?.Value : null;
            if (value.Kind == ExcelCellDataKind.Error
                || cachedType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error) {
                return FormulaArgumentValue.Error(value.CachedText ?? value.Value?.ToString() ?? "#VALUE!");
            }

            ExcelCellDataKind sourceKind = value.Kind == ExcelCellDataKind.Formula
                ? value.Value is bool ? ExcelCellDataKind.Boolean
                    : value.Value is double ? ExcelCellDataKind.Number
                    : value.Value != null ? ExcelCellDataKind.Text : ExcelCellDataKind.Blank
                : value.Kind;
            if (value.Kind == ExcelCellDataKind.Text
                || cachedType == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                || cachedType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                || cachedType == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                return new FormulaArgumentValue(null, value.Value?.ToString() ?? string.Empty,
                    sourceCellKind: sourceKind, isUnevaluatedFormulaCache: unevaluatedFormulaCache);
            }

            if (value.Value is bool boolean) {
                return new FormulaArgumentValue(boolean ? 1 : 0, boolean ? "1" : "0", isBoolean: true,
                    sourceCellKind: sourceKind, isUnevaluatedFormulaCache: unevaluatedFormulaCache);
            }

            if (TryParseFormulaErrorLiteral(value.CachedText ?? value.Value?.ToString() ?? string.Empty, out string errorCode)) {
                return FormulaArgumentValue.Error(errorCode);
            }

            if (value.Value is double d) {
                return new FormulaArgumentValue(d, value.CachedText, sourceCellKind: sourceKind, isUnevaluatedFormulaCache: unevaluatedFormulaCache);
            }

            if (double.TryParse(value.CachedText, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) {
                return new FormulaArgumentValue(parsed, value.CachedText, sourceCellKind: sourceKind, isUnevaluatedFormulaCache: unevaluatedFormulaCache);
            }

            return new FormulaArgumentValue(null, value.Value?.ToString(),
                sourceCellKind: sourceKind, isUnevaluatedFormulaCache: unevaluatedFormulaCache);
        }

        private static string? NormalizeFormulaCellReference(string? reference) {
            if (string.IsNullOrWhiteSpace(reference)) {
                return null;
            }

            var cellRef = A1.ParseCellRef(reference!.Trim().Replace("$", string.Empty));
            if (cellRef.Row <= 0
                || cellRef.Col <= 0
                || cellRef.Row > A1.MaxRows
                || cellRef.Col > A1.MaxColumns) {
                return null;
            }

            return A1.CellReference(cellRef.Row, cellRef.Col);
        }

        private string GetFormulaEvaluationCacheKey(string reference) {
            return Name + "!" + reference;
        }

        private static bool TryParseFormulaErrorLiteral(string token, out string errorCode) {
            string value = token.Trim();
            if (value.StartsWith("=", StringComparison.Ordinal)) {
                value = value.Substring(1).Trim();
            }

            switch (value.ToUpperInvariant()) {
                case "#NULL!":
                    errorCode = "#NULL!";
                    return true;
                case "#DIV/0!":
                    errorCode = "#DIV/0!";
                    return true;
                case "#VALUE!":
                    errorCode = "#VALUE!";
                    return true;
                case "#REF!":
                    errorCode = "#REF!";
                    return true;
                case "#NAME?":
                    errorCode = "#NAME?";
                    return true;
                case "#NUM!":
                    errorCode = "#NUM!";
                    return true;
                case "#N/A":
                    errorCode = "#N/A";
                    return true;
                case "#CALC!":
                case "#SPILL!":
                    errorCode = value.ToUpperInvariant();
                    return true;
                default:
                    errorCode = string.Empty;
                    return false;
            }
        }

    }
}
