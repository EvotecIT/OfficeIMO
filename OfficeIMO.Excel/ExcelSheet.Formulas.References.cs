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

        private bool TryResolveTextArgumentValues(IEnumerable<string> tokens, out List<string> values) {
            values = new List<string>();
            foreach (string token in tokens) {
                if (TryResolveFormulaRange(token, out var rangeValues)) {
                    foreach (var rangeValue in rangeValues) {
                        if (rangeValue.IsUnresolvedFormula) {
                            values.Clear();
                            return false;
                        }

                        values.Add(FormulaValueToText(rangeValue));
                    }

                    continue;
                }

                if (!TryResolveTextArgument(token, out string value)) {
                    values.Clear();
                    return false;
                }

                values.Add(value);
            }

            return true;
        }

        private bool TryResolveTextArgument(string token, out string text) {
            text = string.Empty;
            if (!TryResolveFormulaArgument(token, out FormulaArgumentValue value)) {
                return false;
            }

            if (value.IsUnresolvedFormula) {
                return false;
            }

            text = FormulaValueToText(value);
            return true;
        }

        private bool TryResolveBooleanArgument(string token, out bool value) {
            string trimmed = token.Trim();
            if (trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) {
                value = true;
                return true;
            }

            if (trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) {
                value = false;
                return true;
            }

            if (TryEvaluateFormulaOrNumeric(trimmed, out double numeric)) {
                value = Math.Abs(numeric) >= double.Epsilon;
                return true;
            }

            value = false;
            return false;
        }

        private static string FormulaValueToText(FormulaArgumentValue value) {
            return value.ErrorCode ?? value.Text ?? (value.Number.HasValue ? InvariantNumberText.Get(value.Number.Value) : string.Empty);
        }

        private bool TryResolveFormulaArgument(string token, out FormulaArgumentValue value) {
            string trimmed = token.Trim();
            if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"') {
                value = new FormulaArgumentValue(null, trimmed.Substring(1, trimmed.Length - 2).Replace("\"\"", "\""));
                return true;
            }

            if (trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) {
                value = new FormulaArgumentValue(1d, "1");
                return true;
            }

            if (trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) {
                value = new FormulaArgumentValue(0d, "0");
                return true;
            }

            if (TryParseFormulaErrorLiteral(trimmed, out string errorCode)) {
                value = FormulaArgumentValue.Error(errorCode);
                return true;
            }

            if (TryParseQualifiedFormulaCellReference(trimmed, out ExcelSheet sheet, out int row, out int column)) {
                value = sheet.ResolveCellArgument(row, column);
                return true;
            }

            if (TryResolveFormulaRangeReference(trimmed, out ExcelSheet rangeSheet, out int r1, out int c1, out int r2, out int c2)
                && r1 == r2
                && c1 == c2) {
                value = rangeSheet.ResolveCellArgument(r1, c1);
                return true;
            }

            if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)) {
                value = new FormulaArgumentValue(numeric, trimmed);
                return true;
            }

            if (TryEvaluateFormulaValue(trimmed, out value)) {
                return true;
            }

            value = default;
            return false;
        }

        private static bool IsExactLookupMode(string token) {
            string value = token.Trim();
            return value == "0" || value.Equals("FALSE", StringComparison.OrdinalIgnoreCase);
        }

        private static bool FormulaValuesEqual(FormulaArgumentValue left, FormulaArgumentValue right) {
            if (left.Number.HasValue && right.Number.HasValue) {
                return Math.Abs(left.Number.Value - right.Number.Value) < 0.0000001;
            }

            string leftText = left.Text ?? (left.Number.HasValue ? InvariantNumberText.Get(left.Number.Value) : string.Empty);
            string rightText = right.Text ?? (right.Number.HasValue ? InvariantNumberText.Get(right.Number.Value) : string.Empty);
            return string.Equals(leftText, rightText, StringComparison.OrdinalIgnoreCase);
        }

        private bool TryResolveFormulaArguments(string args, out List<FormulaArgumentValue> values) {
            values = new List<FormulaArgumentValue>();
            int remainingCellBudget = MaxResolvedFormulaRangeCells;
            foreach (string trimmed in SplitFormulaArguments(args)) {
                if (TryResolveFormulaRange(trimmed, out var rangeValues, ref remainingCellBudget)) {
                    values.AddRange(rangeValues);
                    continue;
                }

                if (TryParseQualifiedFormulaCellReference(trimmed, out ExcelSheet sheetReference, out int cellRow, out int cellColumn)) {
                    values.Add(sheetReference.ResolveCellArgument(cellRow, cellColumn));
                    continue;
                }

                if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric)) {
                    values.Add(new FormulaArgumentValue(numeric, trimmed));
                    continue;
                }

                if (TryResolveFormulaArgument(trimmed, out var argumentValue)) {
                    values.Add(argumentValue);
                    continue;
                }

                values.Clear();
                return false;
            }

            return true;
        }

        private static IReadOnlyList<string> SplitFormulaArguments(string args) {
            var tokens = new List<string>();
            var builder = new StringBuilder();
            int depth = 0;
            int bracketDepth = 0;
            bool inString = false;

            for (int index = 0; index < args.Length; index++) {
                char ch = args[index];
                if (ch == '"') {
                    builder.Append(ch);
                    if (inString && index + 1 < args.Length && args[index + 1] == '"') {
                        index++;
                        builder.Append(args[index]);
                        continue;
                    }

                    inString = !inString;
                    continue;
                }

                if (!inString && ch == '(') {
                    depth++;
                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == ')') {
                    depth--;
                    if (depth < 0) {
                        return Array.Empty<string>();
                    }

                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == '[') {
                    bracketDepth++;
                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == ']') {
                    bracketDepth--;
                    if (bracketDepth < 0) {
                        return Array.Empty<string>();
                    }

                    builder.Append(ch);
                    continue;
                }

                if (!inString && ch == ',' && depth == 0 && bracketDepth == 0) {
                    AddToken(tokens, builder);
                    continue;
                }

                builder.Append(ch);
            }

            if (depth != 0 || bracketDepth != 0 || inString) {
                return Array.Empty<string>();
            }

            AddToken(tokens, builder);
            return tokens;
        }

        private static void AddToken(List<string> tokens, StringBuilder builder) {
            string token = builder.ToString().Trim();
            if (token.Length > 0) {
                tokens.Add(token);
            }

            builder.Clear();
        }

        private static bool TryConvertFormulaAValues(IReadOnlyList<FormulaArgumentValue> values, out List<double> numbers) {
            numbers = new List<double>();
            foreach (var value in values) {
                if (value.IsUnresolvedFormula || value.IsError) {
                    numbers.Clear();
                    return false;
                }

                if (value.Number.HasValue) {
                    numbers.Add(value.Number.Value);
                    continue;
                }

                if (value.Text != null) {
                    numbers.Add(0d);
                }
            }

            return true;
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

            var argument = sheet.ResolveCellArgument(row, column);
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

        private bool TryResolveFormulaRangeReference(string token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
            if (TryParseQualifiedFormulaRange(token, out sheet, out r1, out c1, out r2, out c2)) {
                return true;
            }

            if (TryParseQualifiedFormulaCellReference(token, out sheet, out r1, out c1)) {
                r2 = r1;
                c2 = c1;
                return true;
            }

            if (TryResolveTableReferenceRange(token, out sheet, out r1, out c1, out r2, out c2)) {
                return true;
            }

            return TryResolveDefinedNameRange(token, out sheet, out r1, out c1, out r2, out c2);
        }

        private bool TryResolveTableReferenceRange(string token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
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
                            _formulaEvaluationDepthFrames = _formulaEvaluationDepthFrames
                        };
                    return TryResolveTableReferenceRange(table, sections, out r1, out c1, out r2, out c2);
                }
            }

            return false;
        }

        private static bool TryParseStructuredTableReference(string token, out string tableName, out List<string> sections) {
            string value = token.Trim();
            tableName = string.Empty;
            sections = new List<string>();
            if (value.Length == 0 || value.IndexOf('!') >= 0) {
                return false;
            }

            int bracketStart = value.IndexOf('[');
            tableName = bracketStart < 0 ? value : value.Substring(0, bracketStart);
            if (!IsFormulaDefinedNameToken(tableName)) {
                return false;
            }

            if (bracketStart < 0) {
                return true;
            }

            string specifier = value.Substring(bracketStart);
            if (specifier.Length >= 4 && specifier.StartsWith("[[", StringComparison.Ordinal) && specifier.EndsWith("]]", StringComparison.Ordinal)) {
                return TryParseStructuredTableSectionList(specifier.Substring(1, specifier.Length - 2), sections);
            }

            if (specifier.Length >= 2 && specifier[0] == '[' && specifier[specifier.Length - 1] == ']') {
                string section = specifier.Substring(1, specifier.Length - 2).Trim();
                if (section.Length == 0 || section.IndexOf('[') >= 0 || section.IndexOf(']') >= 0) {
                    return false;
                }

                sections.Add(section);
                return true;
            }

            return false;
        }

        private static bool TryParseStructuredTableSectionList(string value, List<string> sections) {
            int index = 0;
            while (index < value.Length) {
                if (value[index] != '[') {
                    return false;
                }

                int end = value.IndexOf(']', index + 1);
                if (end < 0) {
                    return false;
                }

                string section = value.Substring(index + 1, end - index - 1).Trim();
                if (section.Length == 0 || section.IndexOf('[') >= 0 || section.IndexOf(']') >= 0) {
                    return false;
                }

                sections.Add(section);
                index = end + 1;
                if (index == value.Length) {
                    return true;
                }

                if (value[index] != ',') {
                    return false;
                }

                index++;
            }

            return sections.Count > 0;
        }

        private static bool TryResolveTableReferenceRange(Table table, IReadOnlyList<string> sections, out int r1, out int c1, out int r2, out int c2) {
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

            string? item = null;
            string area = "#Data";
            if (sections.Count == 1) {
                if (IsStructuredTableAreaSpecifier(sections[0])) {
                    area = sections[0];
                } else {
                    item = sections[0];
                }
            } else if (sections.Count == 2) {
                if (!IsStructuredTableAreaSpecifier(sections[0])) {
                    return false;
                }

                area = sections[0];
                item = sections[1];
            } else if (sections.Count != 0) {
                return false;
            }

            if (!TryResolveTableAreaRows(area, tableR1, tableR2, headerRows, totalsRows, out r1, out r2)) {
                return false;
            }

            c1 = tableC1;
            c2 = tableC2;
            if (!string.IsNullOrWhiteSpace(item)) {
                int offset = ResolveTableColumnOffset(table, item!);
                if (offset < 0) {
                    return false;
                }

                c1 = tableC1 + offset;
                c2 = c1;
            }

            return r1 <= r2 && c1 <= c2;
        }

        private static bool TryResolveTableAreaRows(string area, int tableR1, int tableR2, uint headerRows, uint totalsRows, out int r1, out int r2) {
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
                || string.Equals(section, "#Totals", StringComparison.OrdinalIgnoreCase);
        }

        private static int ResolveTableColumnOffset(Table table, string columnName) {
            int index = 0;
            foreach (var tableColumn in table.TableColumns?.Elements<TableColumn>() ?? Enumerable.Empty<TableColumn>()) {
                if (string.Equals(tableColumn.Name?.Value, columnName, StringComparison.OrdinalIgnoreCase)) {
                    return index;
                }

                index++;
            }

            return -1;
        }

        private bool TryResolveDefinedNameRange(string token, out ExcelSheet sheet, out int r1, out int c1, out int r2, out int c2) {
            sheet = this;
            r1 = 0;
            c1 = 0;
            r2 = 0;
            c2 = 0;

            if (!TrySplitQualifiedReference(token, out string? sheetName, out string name)
                || !IsFormulaDefinedNameToken(name)) {
                return false;
            }

            var definedNames = WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return false;
            }

            var sheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            int? localSheetIndex = null;
            ExcelSheet defaultSheet = this;
            if (sheetName != null) {
                int index = sheets.FindIndex(candidate => string.Equals(candidate.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
                if (index < 0 || !TryGetFormulaReferenceSheet(sheetName, out defaultSheet)) {
                    return false;
                }

                localSheetIndex = index;
            } else {
                int index = sheets.FindIndex(candidate => string.Equals(candidate.Name?.Value, Name, StringComparison.OrdinalIgnoreCase));
                if (index >= 0) {
                    localSheetIndex = index;
                }
            }

            DefinedName? definedName = null;
            if (localSheetIndex.HasValue) {
                definedName = definedNames.Elements<DefinedName>()
                    .FirstOrDefault(candidate => candidate.LocalSheetId?.Value == (uint)localSheetIndex.Value
                        && string.Equals(candidate.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
            }

            if (definedName == null && sheetName == null) {
                definedName = definedNames.Elements<DefinedName>()
                    .FirstOrDefault(candidate => candidate.LocalSheetId == null
                        && string.Equals(candidate.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
            }

            if (definedName == null || IsBuiltInFormulaDefinedName(definedName.Name?.Value)) {
                return false;
            }

            string reference = (definedName.Text ?? string.Empty).Trim();
            if (reference.StartsWith("=", StringComparison.Ordinal)) {
                reference = reference.Substring(1).Trim();
            }

            if (reference.Length == 0
                || reference.IndexOf(',') >= 0
                || reference.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0) {
                return false;
            }

            if (definedName.LocalSheetId?.Value is uint scopedIndex
                && scopedIndex < (uint)sheets.Count
                && sheets[(int)scopedIndex].Name?.Value is string scopedSheetName
                && !string.Equals(scopedSheetName, Name, StringComparison.OrdinalIgnoreCase)) {
                _ = TryGetFormulaReferenceSheet(scopedSheetName, out defaultSheet);
            }

            if (TryParseQualifiedFormulaRange(reference, defaultSheet, out sheet, out r1, out c1, out r2, out c2)) {
                return true;
            }

            if (TryParseQualifiedFormulaCellReference(reference, defaultSheet, out sheet, out r1, out c1)) {
                r2 = r1;
                c2 = c1;
                return true;
            }

            return false;
        }

        private static bool IsFormulaDefinedNameToken(string token) {
            if (string.IsNullOrWhiteSpace(token)
                || string.Equals(token, "TRUE", StringComparison.OrdinalIgnoreCase)
                || string.Equals(token, "FALSE", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            char first = token[0];
            if (!char.IsLetter(first) && first != '_') {
                return false;
            }

            foreach (char character in token) {
                if (!char.IsLetterOrDigit(character) && character != '_' && character != '.') {
                    return false;
                }
            }

            return true;
        }

        private static bool IsBuiltInFormulaDefinedName(string? name) {
            return !string.IsNullOrWhiteSpace(name)
                && name!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase);
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
                _formulaEvaluationDepthFrames = _formulaEvaluationDepthFrames
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

        private FormulaArgumentValue ResolveCellArgument(int row, int column) {
            var cell = TryGetExistingCell(row, column);
            bool unresolvedFormula = false;
            if (cell?.CellFormula != null && _formulaEvaluationCache != null) {
                if (TryEvaluateFormulaCellValue(cell, out FormulaArgumentValue formulaResult)) {
                    return formulaResult;
                }

                if (_formulaEvaluationDepthFrames != null
                    && _formulaEvaluationDepthFrames.Count > 0
                    && _formulaEvaluationDepthFrames.Peek().DependencyGuardBlocked) {
                    return FormulaArgumentValue.UnresolvedFormula();
                }

                unresolvedFormula = true;
            }

            var value = GetCellValueSnapshot(row, column);
            if (unresolvedFormula && value.Value == null && string.IsNullOrEmpty(value.CachedText)) {
                return FormulaArgumentValue.UnresolvedFormula();
            }

            if (value.Kind == ExcelCellDataKind.Error) {
                return FormulaArgumentValue.Error(value.CachedText ?? value.Value?.ToString() ?? "#VALUE!");
            }

            if (TryParseFormulaErrorLiteral(value.CachedText ?? value.Value?.ToString() ?? string.Empty, out string errorCode)) {
                return FormulaArgumentValue.Error(errorCode);
            }

            if (value.Value is double d) {
                return new FormulaArgumentValue(d, value.CachedText);
            }

            if (double.TryParse(value.CachedText, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) {
                return new FormulaArgumentValue(parsed, value.CachedText);
            }

            return new FormulaArgumentValue(null, value.Value?.ToString());
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
                default:
                    errorCode = string.Empty;
                    return false;
            }
        }

    }
}
