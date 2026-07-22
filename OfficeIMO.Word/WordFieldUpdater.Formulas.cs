using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private static readonly Regex TableReferenceFunctionPattern = new Regex(
            @"\b(?<function>SUM|AVERAGE|MIN|MAX|PRODUCT|COUNT)\s*\(\s*(?<reference>ABOVE|BELOW|LEFT|RIGHT)\s*\)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex TableCellReferencePattern = new Regex(
            @"(?<![A-Z0-9_])(?<start>\$?[A-Z]{1,3}\$?[1-9][0-9]*)(?:\s*:\s*(?<end>\$?[A-Z]{1,3}\$?[1-9][0-9]*))?(?![A-Z0-9_])",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex FormulaNumericPicturePattern = new Regex(
            @"^(?<expression>.*?)\s+\\#\s*(?<format>""[^""]*""|\S+)(?<formats>(?:\s+\\\*\s*(?:""[^""]*""|\S+))*)\s*$",
            RegexOptions.CultureInvariant | RegexOptions.Singleline);

        private static readonly Regex FormulaTrailingFormatSwitchPattern = new Regex(
            @"^(?<expression>.*?)(?<formats>(?:\s+\\\*\s*(?:""[^""]*""|\S+))+)\s*$",
            RegexOptions.CultureInvariant | RegexOptions.Singleline);

        private static readonly Regex FormulaFormatSwitchPattern = new Regex(
            @"\\\*\s*(?:""(?<format>[^""]*)""|(?<format>\S+))",
            RegexOptions.CultureInvariant | RegexOptions.Singleline);

        private static bool TryEvaluateFormula(
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;

            string expression = parsed.Instructions.FirstOrDefault() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(expression)) {
                status = WordFieldUpdateStatus.ParseError;
                message = "Formula field is missing an expression.";
                return false;
            }

            string? numericPicture = null;
            if (!TryExtractFormulaNumericPicture(expression, out expression, out numericPicture, out IReadOnlyList<WordFieldFormat> formulaFormatSwitches, out string? numericPictureDiagnostic)) {
                message = numericPictureDiagnostic ?? "Formula numeric picture switch could not be parsed.";
                return false;
            }

            if (!TryResolveTableFormulaReferences(candidate, expression, out expression, out string? tableDiagnostic)) {
                message = tableDiagnostic ?? "Formula table reference could not be evaluated.";
                return false;
            }

            if (!TryResolveRnCnTableCellReferences(candidate, expression, out expression, out tableDiagnostic)) {
                message = tableDiagnostic ?? "Formula RnCn table cell reference could not be evaluated.";
                return false;
            }

            if (!TryResolveExplicitTableCellReferences(candidate, expression, out expression, out tableDiagnostic)) {
                message = tableDiagnostic ?? "Formula table cell reference could not be evaluated.";
                return false;
            }

            if (!FormulaExpressionParser.TryEvaluate(expression, out decimal result, out string? diagnostic)) {
                message = diagnostic ?? "Formula expression could not be evaluated.";
                return false;
            }

            WordFieldFormat? formulaFormat = GetLastMeaningfulFormat(formulaFormatSwitches);
            if (!string.IsNullOrWhiteSpace(numericPicture) && formulaFormat != null) {
                message = "Formula fields cannot combine numeric picture and general format switches for deterministic refresh.";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(numericPicture)) {
                if (!TryFormatFormulaValue(result, numericPicture, out value, out string? formatDiagnostic)) {
                    message = formatDiagnostic ?? "Formula result could not be formatted.";
                    return false;
                }
            } else if (formulaFormat != null) {
                if (!TryFormatFormulaGeneralNumericValue(result, formulaFormat.Value, out value, out string? formatDiagnostic)) {
                    message = formatDiagnostic ?? "Formula general numeric format switch could not be applied.";
                    return false;
                }
            } else if (!TryFormatFormulaValue(result, numericPicture, out value, out string? formatDiagnostic)) {
                message = formatDiagnostic ?? "Formula result could not be formatted.";
                return false;
            }

            status = WordFieldUpdateStatus.Updated;
            message = !string.IsNullOrWhiteSpace(numericPicture)
                ? "Updated from bounded OfficeIMO arithmetic/function evaluator with numeric picture formatting."
                : formulaFormat != null
                    ? "Updated from bounded OfficeIMO arithmetic/function evaluator with general numeric formatting."
                    : "Updated from bounded OfficeIMO arithmetic/function evaluator.";
            return true;
        }

        private static string FormatFormulaValue(decimal value) {
            decimal normalized = decimal.Truncate(value) == value ? decimal.Truncate(value) : value;
            return normalized.ToString("G29", CultureInfo.InvariantCulture);
        }

        private static bool TryFormatFormulaValue(decimal value, string? numericPicture, out string formattedValue, out string? diagnostic) {
            diagnostic = null;
            if (string.IsNullOrWhiteSpace(numericPicture)) {
                formattedValue = FormatFormulaValue(value);
                return true;
            }

            string format = TrimFormulaFormatQuotes(numericPicture ?? string.Empty);
            if (format.Length == 0) {
                formattedValue = string.Empty;
                diagnostic = "Formula numeric picture switch is empty.";
                return false;
            }

            if (!TrySelectNumericPictureSection(value, format, out decimal formatValue, out string selectedFormat, out diagnostic)) {
                formattedValue = string.Empty;
                return false;
            }

            if (!TryNormalizeNumericPictureFillSyntax(selectedFormat, out string normalizedFormat, out string? fillDiagnostic)) {
                formattedValue = string.Empty;
                diagnostic = fillDiagnostic ?? $"Formula numeric picture switch '{format}' contains unsupported fill formatting syntax.";
                return false;
            }

            if (normalizedFormat.Length == 0) {
                formattedValue = string.Empty;
                diagnostic = $"Formula numeric picture switch '{format}' contains only layout-dependent fill formatting syntax.";
                return false;
            }

            if (!TryValidateNumericPicture(normalizedFormat, out string? validationDiagnostic)) {
                formattedValue = string.Empty;
                diagnostic = validationDiagnostic ?? $"Formula numeric picture switch '{format}' contains unsupported locale-specific formatting syntax.";
                return false;
            }

            try {
                formattedValue = formatValue.ToString(normalizedFormat, CultureInfo.InvariantCulture);
                return true;
            } catch (FormatException) {
                formattedValue = string.Empty;
                diagnostic = $"Formula numeric picture switch '{format}' is not supported by the bounded OfficeIMO formatter.";
                return false;
            }
        }

        private static bool TryFormatFormulaGeneralNumericValue(decimal value, WordFieldFormat format, out string formattedValue, out string? diagnostic) {
            formattedValue = string.Empty;
            diagnostic = null;

            switch (format) {
                case WordFieldFormat.Arabic:
                    formattedValue = FormatFormulaValue(value);
                    return true;
                case WordFieldFormat.Roman:
                case WordFieldFormat.roman:
                case WordFieldFormat.Ordinal:
                case WordFieldFormat.Alphabetical:
                case WordFieldFormat.ALPHABETICAL:
                case WordFieldFormat.Hex:
                case WordFieldFormat.CardText:
                case WordFieldFormat.OrdText:
                case WordFieldFormat.DollarText:
                    if (value != decimal.Truncate(value) || value < int.MinValue || value > int.MaxValue) {
                        diagnostic = $"Formula format switch {format} requires an integer value in the deterministic refresh range.";
                        return false;
                    }

                    int integerValue = decimal.ToInt32(value);
                    if (RequiresNonNegativeNumber(format) && integerValue < 0) {
                        diagnostic = $"Formula format switch {format} requires a non-negative value for deterministic numeric refresh.";
                        return false;
                    }

                    formattedValue = FormatSequenceValue(integerValue, new[] { format });
                    return true;
                default:
                    diagnostic = $"Formula format switch {format} is not supported for deterministic numeric refresh.";
                    return false;
            }
        }

        private static bool TryExtractFormulaNumericPicture(
            string expression,
            out string expressionWithoutPicture,
            out string? numericPicture,
            out IReadOnlyList<WordFieldFormat> formatSwitches,
            out string? diagnostic) {
            expressionWithoutPicture = expression;
            numericPicture = null;
            formatSwitches = Array.Empty<WordFieldFormat>();
            diagnostic = null;

            Match match = FormulaNumericPicturePattern.Match(expression);
            if (!match.Success) {
                if (expression.IndexOf(@"\#", StringComparison.Ordinal) >= 0) {
                    diagnostic = "Formula numeric picture switch must appear at the end of the field instruction.";
                    return false;
                }

                Match trailingFormatMatch = FormulaTrailingFormatSwitchPattern.Match(expression);
                if (trailingFormatMatch.Success) {
                    if (!TryParseFormulaFormatSwitches(trailingFormatMatch.Groups["formats"].Value, out formatSwitches, out diagnostic)) {
                        return false;
                    }

                    expressionWithoutPicture = trailingFormatMatch.Groups["expression"].Value.Trim();
                }

                return true;
            }

            expressionWithoutPicture = match.Groups["expression"].Value.Trim();
            numericPicture = match.Groups["format"].Value.Trim();
            if (!TryParseFormulaFormatSwitches(match.Groups["formats"].Value, out formatSwitches, out diagnostic)) {
                return false;
            }

            return true;
        }

        private static bool TryParseFormulaFormatSwitches(string switchesText, out IReadOnlyList<WordFieldFormat> formatSwitches, out string? diagnostic) {
            var switches = new List<WordFieldFormat>();
            diagnostic = null;

            foreach (Match match in FormulaFormatSwitchPattern.Matches(switchesText)) {
                string formatSwitch = match.Groups["format"].Value.Trim();
                bool success = Enum.TryParse(formatSwitch, false, out WordFieldFormat fieldFormat) ||
                    Enum.TryParse(formatSwitch, true, out fieldFormat);
                if (!success) {
                    formatSwitches = Array.Empty<WordFieldFormat>();
                    diagnostic = $"Formula format switch \\* {formatSwitch} is not recognized by OfficeIMO.";
                    return false;
                }

                switches.Add(fieldFormat);
            }

            formatSwitches = switches;
            return true;
        }

        private static string TrimFormulaFormatQuotes(string value) {
            value = value.Trim();
            return value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"'
                ? value.Substring(1, value.Length - 2)
                : value;
        }

        private static bool TrySelectNumericPictureSection(
            decimal value,
            string format,
            out decimal formatValue,
            out string selectedFormat,
            out string? diagnostic) {
            formatValue = value;
            selectedFormat = string.Empty;
            diagnostic = null;

            string[] sectionTexts = SplitNumericPictureSections(format);
            if (sectionTexts.Length > 3) {
                diagnostic = $"Formula numeric picture switch '{format}' contains more than three sections.";
                return false;
            }

            var sections = new List<NumericPictureSection>();
            foreach (string sectionText in sectionTexts) {
                if (!TryParseNumericPictureSection(sectionText, out NumericPictureSection section, out diagnostic)) {
                    return false;
                }

                sections.Add(section);
            }

            bool hasExplicitCondition = sections.Any(section => section.Condition != null);
            int selectedIndex = 0;

            if (hasExplicitCondition) {
                selectedIndex = sections.FindIndex(section => section.Condition == null || section.Condition.Value.Matches(value));
                if (selectedIndex < 0) {
                    selectedIndex = sections.Count - 1;
                }
            } else if (sections.Count > 1 && value < 0) {
                selectedIndex = 1;
            } else if (sections.Count > 2 && value == 0) {
                selectedIndex = 2;
            }

            NumericPictureSection selected = sections[selectedIndex];
            selectedFormat = selected.Format;
            if (selectedFormat.Length == 0) {
                diagnostic = $"Formula numeric picture switch '{format}' selected an empty section.";
                return false;
            }

            if (value < 0 && (selectedIndex > 0 || selected.Condition != null)) {
                try {
                    formatValue = Math.Abs(value);
                } catch (OverflowException) {
                    diagnostic = "Formula numeric picture switch could not normalize the negative value.";
                    return false;
                }
            }

            return true;
        }

        private static bool TryParseNumericPictureSection(string sectionText, out NumericPictureSection section, out string? diagnostic) {
            section = default;
            diagnostic = null;

            string text = sectionText.TrimStart();
            NumericPictureCondition? condition = null;

            while (text.StartsWith("[", StringComparison.Ordinal)) {
                int end = text.IndexOf(']');
                if (end < 0) {
                    diagnostic = $"Formula numeric picture section '{sectionText}' contains an unterminated bracket token.";
                    return false;
                }

                string token = text.Substring(1, end - 1).Trim();
                if (TryParseNumericPictureCondition(token, out NumericPictureCondition parsedCondition)) {
                    if (condition != null) {
                        diagnostic = $"Formula numeric picture section '{sectionText}' contains more than one condition.";
                        return false;
                    }

                    condition = parsedCondition;
                } else if (!IsNumericPictureColorToken(token)) {
                    diagnostic = $"Formula numeric picture section '{sectionText}' contains unsupported bracket token '[{token}]'.";
                    return false;
                }

                text = text.Substring(end + 1).TrimStart();
            }

            section = new NumericPictureSection(text, condition);
            return true;
        }

        private static bool TryParseNumericPictureCondition(string token, out NumericPictureCondition condition) {
            condition = default;

            string[] operators = new[] { ">=", "<=", "<>", ">", "<", "=" };
            foreach (string comparisonOperator in operators) {
                if (!token.StartsWith(comparisonOperator, StringComparison.Ordinal)) {
                    continue;
                }

                string comparisonValueText = token.Substring(comparisonOperator.Length).Trim();
                if (!decimal.TryParse(comparisonValueText, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal comparisonValue)) {
                    return false;
                }

                condition = new NumericPictureCondition(comparisonOperator, comparisonValue);
                return true;
            }

            return false;
        }

        private static bool IsNumericPictureColorToken(string token) {
            switch (token.ToUpperInvariant()) {
                case "BLACK":
                case "BLUE":
                case "CYAN":
                case "GREEN":
                case "MAGENTA":
                case "RED":
                case "WHITE":
                case "YELLOW":
                    return true;
                default:
                    return Regex.IsMatch(token, @"^COLOR\s*[1-9][0-9]?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }
        }

        private static string[] SplitNumericPictureSections(string format) {
            var sections = new List<string>();
            var currentSection = new StringBuilder();
            bool escaped = false;

            foreach (char current in format) {
                if (escaped) {
                    currentSection.Append(current);
                    escaped = false;
                    continue;
                }

                if (current == '\\') {
                    currentSection.Append(current);
                    escaped = true;
                    continue;
                }

                if (current == ';') {
                    sections.Add(currentSection.ToString());
                    currentSection.Clear();
                    continue;
                }

                currentSection.Append(current);
            }

            sections.Add(currentSection.ToString());
            return sections.ToArray();
        }

        private static bool TryValidateNumericPicture(string format, out string? diagnostic) {
            diagnostic = null;
            bool escaped = false;

            foreach (char current in format) {
                if (escaped) {
                    if (char.IsControl(current)) {
                        diagnostic = "Formula numeric picture switch contains an escaped control character.";
                        return false;
                    }

                    escaped = false;
                    continue;
                }

                if (current == '\\') {
                    escaped = true;
                    continue;
                }

                if (char.IsLetterOrDigit(current) ||
                    current == '#' ||
                    current == '0' ||
                    current == '.' ||
                    current == ',' ||
                    current == '%' ||
                    current == '\u2030' ||
                    current == ';' ||
                    current == '$' ||
                    current == '(' ||
                    current == ')' ||
                    current == '+' ||
                    current == '-' ||
                    current == ' ') {
                    continue;
                }

                diagnostic = $"Formula numeric picture switch contains unsupported formatting character '{current}'.";
                return false;
            }

            if (escaped) {
                diagnostic = "Formula numeric picture switch contains a dangling escape character.";
                return false;
            }

            return true;
        }

        private static bool TryNormalizeNumericPictureFillSyntax(string format, out string normalizedFormat, out string? diagnostic) {
            diagnostic = null;
            var builder = new StringBuilder(format.Length);
            bool escaped = false;

            for (int i = 0; i < format.Length; i++) {
                char current = format[i];
                if (escaped) {
                    builder.Append(current);
                    escaped = false;
                    continue;
                }

                if (current == '\\') {
                    builder.Append(current);
                    escaped = true;
                    continue;
                }

                if (current == '*') {
                    if (i == format.Length - 1) {
                        normalizedFormat = string.Empty;
                        diagnostic = "Formula numeric picture switch contains an unsupported dangling fill formatting token.";
                        return false;
                    }

                    char fillCharacter = format[++i];
                    if (char.IsControl(fillCharacter)) {
                        normalizedFormat = string.Empty;
                        diagnostic = "Formula numeric picture switch contains an unsupported control character fill token.";
                        return false;
                    }

                    continue;
                }

                builder.Append(current);
            }

            if (escaped) {
                normalizedFormat = string.Empty;
                diagnostic = "Formula numeric picture switch contains a dangling escape character.";
                return false;
            }

            normalizedFormat = builder.ToString();
            return true;
        }

        private readonly struct NumericPictureSection {
            internal NumericPictureSection(string format, NumericPictureCondition? condition) {
                Format = format;
                Condition = condition;
            }

            internal string Format { get; }

            internal NumericPictureCondition? Condition { get; }
        }

        private readonly struct NumericPictureCondition {
            internal NumericPictureCondition(string comparisonOperator, decimal comparisonValue) {
                ComparisonOperator = comparisonOperator;
                ComparisonValue = comparisonValue;
            }

            private string ComparisonOperator { get; }

            private decimal ComparisonValue { get; }

            internal bool Matches(decimal value) {
                switch (ComparisonOperator) {
                    case ">=":
                        return value >= ComparisonValue;
                    case "<=":
                        return value <= ComparisonValue;
                    case "<>":
                        return value != ComparisonValue;
                    case ">":
                        return value > ComparisonValue;
                    case "<":
                        return value < ComparisonValue;
                    case "=":
                        return value == ComparisonValue;
                    default:
                        return false;
                }
            }
        }

        private static bool TryResolveTableFormulaReferences(
            MutableFieldCandidate candidate,
            string expression,
            out string resolvedExpression,
            out string? diagnostic) {
            resolvedExpression = expression;
            diagnostic = null;

            MatchCollection matches = TableReferenceFunctionPattern.Matches(expression);
            if (matches.Count == 0) {
                return true;
            }

            var builder = new StringBuilder();
            int startIndex = 0;

            foreach (Match match in matches) {
                string functionName = match.Groups["function"].Value;
                string referenceName = match.Groups["reference"].Value;

                if (!TryResolveTableReference(candidate, referenceName, out IReadOnlyList<decimal> values, out diagnostic)) {
                    return false;
                }

                if (values.Count == 0) {
                    diagnostic = $"Formula table reference {referenceName.ToUpperInvariant()} did not resolve any numeric cells.";
                    return false;
                }

                builder.Append(expression, startIndex, match.Index - startIndex);
                builder.Append(functionName);
                builder.Append('(');
                builder.Append(string.Join(", ", values.Select(FormatFormulaValue)));
                builder.Append(')');
                startIndex = match.Index + match.Length;
            }

            builder.Append(expression, startIndex, expression.Length - startIndex);
            resolvedExpression = builder.ToString();
            return true;
        }

        private static bool TryResolveExplicitTableCellReferences(
            MutableFieldCandidate candidate,
            string expression,
            out string resolvedExpression,
            out string? diagnostic) {
            resolvedExpression = expression;
            diagnostic = null;

            MatchCollection matches = TableCellReferencePattern.Matches(expression);
            if (matches.Count == 0) {
                return true;
            }

            if (!TryGetFieldTable(candidate, out Table? table, out diagnostic) || table == null) {
                return false;
            }

            List<TableRow> rows = table.Elements<TableRow>().ToList();
            var builder = new StringBuilder();
            int startIndex = 0;

            foreach (Match match in matches) {
                if (!TryParseTableAddress(match.Groups["start"].Value, out TableAddress startAddress)) {
                    diagnostic = $"Formula table cell reference {match.Groups["start"].Value} could not be parsed.";
                    return false;
                }

                TableAddress endAddress = startAddress;
                if (match.Groups["end"].Success && !TryParseTableAddress(match.Groups["end"].Value, out endAddress)) {
                    diagnostic = $"Formula table cell reference {match.Groups["end"].Value} could not be parsed.";
                    return false;
                }

                if (!TryResolveTableAddressRange(rows, startAddress, endAddress, out IReadOnlyList<decimal> values, out diagnostic)) {
                    return false;
                }

                if (values.Count == 0) {
                    diagnostic = $"Formula table cell reference {match.Value} did not resolve any numeric cells.";
                    return false;
                }

                builder.Append(expression, startIndex, match.Index - startIndex);
                builder.Append(string.Join(", ", values.Select(FormatFormulaValue)));
                startIndex = match.Index + match.Length;
            }

            builder.Append(expression, startIndex, expression.Length - startIndex);
            resolvedExpression = builder.ToString();
            return true;
        }

        private static bool TryResolveTableReference(
            MutableFieldCandidate candidate,
            string referenceName,
            out IReadOnlyList<decimal> values,
            out string? diagnostic) {
            values = Array.Empty<decimal>();
            diagnostic = null;

            if (!TryGetFieldTable(candidate, out Table? table, out diagnostic, out TableCell? currentCell, out TableRow? currentRow)
                || table == null
                || currentCell == null
                || currentRow == null) {
                diagnostic = $"Formula table reference {referenceName.ToUpperInvariant()} requires the field to be inside a table cell.";
                return false;
            }

            List<TableRow> rows = table.Elements<TableRow>().ToList();
            List<TableCellGridPlacement> placements = GetRowCellGridPlacements(currentRow);
            int rowIndex = rows.FindIndex(row => ReferenceEquals(row, currentRow));
            TableCellGridPlacement? currentPlacement = placements.FirstOrDefault(placement => ReferenceEquals(placement.Cell, currentCell));

            if (rowIndex < 0 || currentPlacement == null) {
                diagnostic = $"Formula table reference {referenceName.ToUpperInvariant()} could not locate the field cell.";
                return false;
            }

            IEnumerable<TableCell> sourceCells;
            switch (referenceName.ToUpperInvariant()) {
                case "ABOVE":
                    sourceCells = GetCellsInVisualColumn(rows.Take(rowIndex), currentPlacement.StartColumn);
                    break;
                case "BELOW":
                    sourceCells = GetCellsInVisualColumn(rows.Skip(rowIndex + 1), currentPlacement.StartColumn);
                    break;
                case "LEFT":
                    sourceCells = placements
                        .Where(placement => placement.EndColumn <= currentPlacement.StartColumn)
                        .Select(placement => placement.Cell);
                    break;
                case "RIGHT":
                    sourceCells = placements
                        .Where(placement => placement.StartColumn >= currentPlacement.EndColumn)
                        .Select(placement => placement.Cell);
                    break;
                default:
                    diagnostic = $"Formula table reference {referenceName} is not supported.";
                    return false;
            }

            sourceCells = sourceCells.Where(sourceCell => !IsVerticalMergeContinuation(sourceCell));

            var resolvedValues = new List<decimal>();
            foreach (TableCell sourceCell in sourceCells) {
                if (TryReadTableCellNumber(sourceCell, out decimal number, out bool hasText)) {
                    resolvedValues.Add(number);
                } else if (hasText) {
                    diagnostic = $"Formula table reference {referenceName.ToUpperInvariant()} found non-numeric cell text '{GetTableCellText(sourceCell)}'.";
                    return false;
                }
            }

            values = resolvedValues;
            return true;
        }

        private static bool TryGetFieldTable(MutableFieldCandidate candidate, out Table? table, out string? diagnostic) {
            return TryGetFieldTable(candidate, out table, out diagnostic, out _, out _);
        }

        private static bool TryGetFieldTable(
            MutableFieldCandidate candidate,
            out Table? table,
            out string? diagnostic,
            out TableCell? currentCell,
            out TableRow? currentRow) {
            currentCell = candidate.AnchorElement is TableCell cell
                ? cell
                : candidate.AnchorElement.Ancestors<TableCell>().FirstOrDefault();
            currentRow = currentCell?.Ancestors<TableRow>().FirstOrDefault();
            table = currentRow?.Ancestors<Table>().FirstOrDefault();
            diagnostic = table == null ? "Formula table cell references require the field to be inside a table cell." : null;
            return currentCell != null && currentRow != null && table != null;
        }

        private static bool TryResolveTableAddressRange(
            IReadOnlyList<TableRow> rows,
            TableAddress startAddress,
            TableAddress endAddress,
            out IReadOnlyList<decimal> values,
            out string? diagnostic) {
            values = Array.Empty<decimal>();
            diagnostic = null;

            int startRow = Math.Min(startAddress.RowIndex, endAddress.RowIndex);
            int endRow = Math.Max(startAddress.RowIndex, endAddress.RowIndex);
            int startColumn = Math.Min(startAddress.ColumnIndex, endAddress.ColumnIndex);
            int endColumn = Math.Max(startAddress.ColumnIndex, endAddress.ColumnIndex);

            if (startRow < 0 || endRow >= rows.Count) {
                diagnostic = $"Formula table cell reference {startAddress}:{endAddress} is outside the table row range.";
                return false;
            }

            var resolvedValues = new List<decimal>();
            var resolvedCells = new HashSet<TableCell>();
            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
                List<TableCellGridPlacement> placements = GetRowCellGridPlacements(rows[rowIndex]);
                for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++) {
                    TableCellGridPlacement? placement = placements.FirstOrDefault(item => item.ContainsColumn(columnIndex));
                    if (placement == null || !resolvedCells.Add(placement.Cell) || IsVerticalMergeContinuation(placement.Cell)) {
                        continue;
                    }

                    if (TryReadTableCellNumber(placement.Cell, out decimal number, out bool hasText)) {
                        resolvedValues.Add(number);
                    } else if (hasText) {
                        diagnostic = $"Formula table cell reference {startAddress}:{endAddress} found non-numeric cell text '{GetTableCellText(placement.Cell)}'.";
                        return false;
                    }
                }
            }

            values = resolvedValues;
            return true;
        }

        private static IEnumerable<TableCell> GetCellsInVisualColumn(IEnumerable<TableRow> rows, int columnIndex) {
            foreach (TableRow row in rows) {
                foreach (TableCellGridPlacement placement in GetRowCellGridPlacements(row)) {
                    if (placement.ContainsColumn(columnIndex)) {
                        yield return placement.Cell;
                        break;
                    }
                }
            }
        }

        private static List<TableCellGridPlacement> GetRowCellGridPlacements(TableRow row) {
            List<TableCell> cells = row.Elements<TableCell>().ToList();
            var placements = new List<TableCellGridPlacement>();
            int columnIndex = GetRowGridBefore(row);

            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCell cell = cells[cellIndex];
                if (cell.TableCellProperties?.HorizontalMerge?.Val?.Value == MergedCellValues.Continue) {
                    continue;
                }

                int span = GetGridSpan(cell);
                if (span == 1 && cell.TableCellProperties?.HorizontalMerge?.Val?.Value == MergedCellValues.Restart) {
                    span = CountHorizontalMergeSpan(cells, cellIndex);
                }

                placements.Add(new TableCellGridPlacement(cell, columnIndex, columnIndex + span));
                columnIndex += span;
            }

            return placements;
        }

        private static int GetRowGridBefore(TableRow row) {
            int? gridBefore = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value;
            int value = gridBefore.GetValueOrDefault();
            return value > 0 ? value : 0;
        }

        private static int CountHorizontalMergeSpan(IReadOnlyList<TableCell> cells, int restartIndex) {
            int span = 1;
            for (int index = restartIndex + 1; index < cells.Count; index++) {
                if (cells[index].TableCellProperties?.HorizontalMerge?.Val?.Value != MergedCellValues.Continue) {
                    break;
                }

                span++;
            }

            return span;
        }

        private static int GetGridSpan(TableCell cell) {
            int? span = cell.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value;
            int value = span.GetValueOrDefault();
            return value > 1 ? value : 1;
        }

        private static bool IsVerticalMergeContinuation(TableCell cell) {
            VerticalMerge? verticalMerge = cell.TableCellProperties?.VerticalMerge;
            return verticalMerge != null && verticalMerge.Val?.Value != MergedCellValues.Restart;
        }

        private static bool TryReadTableCellNumber(TableCell cell, out decimal value, out bool hasText) {
            string text = GetTableCellText(cell).Replace('\u00A0', ' ').Trim();
            hasText = text.Length > 0;
            if (text.EndsWith("%", StringComparison.Ordinal)) {
                string numericText = text.Substring(0, text.Length - 1).TrimEnd();
                if (decimal.TryParse(numericText, NumberStyles.Number, CultureInfo.InvariantCulture, out value)) {
                    value /= 100m;
                    return true;
                }

                value = 0m;
                return false;
            }

            return decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out value);
        }

        private static bool TryParseTableAddress(string text, out TableAddress address) {
            address = default;
            string normalized = text.Replace("$", string.Empty).Trim();
            int splitIndex = 0;
            while (splitIndex < normalized.Length && char.IsLetter(normalized[splitIndex])) {
                splitIndex++;
            }

            if (splitIndex == 0 || splitIndex == normalized.Length) {
                return false;
            }

            string columnText = normalized.Substring(0, splitIndex);
            string rowText = normalized.Substring(splitIndex);
            if (!int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out int oneBasedRow) || oneBasedRow <= 0) {
                return false;
            }

            int columnIndex = 0;
            for (int i = 0; i < columnText.Length; i++) {
                char current = char.ToUpperInvariant(columnText[i]);
                if (current < 'A' || current > 'Z') {
                    return false;
                }

                columnIndex = checked(columnIndex * 26 + (current - 'A' + 1));
            }

            address = new TableAddress(columnIndex - 1, oneBasedRow - 1);
            return true;
        }

        private static string GetTableCellText(TableCell cell) {
            return string.Concat(cell.Descendants<Text>().Select(text => text.Text));
        }

        private readonly struct TableAddress {
            internal TableAddress(int columnIndex, int rowIndex) {
                ColumnIndex = columnIndex;
                RowIndex = rowIndex;
            }

            internal int ColumnIndex { get; }

            internal int RowIndex { get; }

            public override string ToString() {
                return $"{FormatColumn(ColumnIndex)}{(RowIndex + 1).ToString(CultureInfo.InvariantCulture)}";
            }

            private static string FormatColumn(int columnIndex) {
                int value = columnIndex + 1;
                var builder = new StringBuilder();
                while (value > 0) {
                    value--;
                    builder.Insert(0, (char)('A' + value % 26));
                    value /= 26;
                }

                return builder.ToString();
            }
        }

        private sealed class TableCellGridPlacement {
            internal TableCellGridPlacement(TableCell cell, int startColumn, int endColumn) {
                Cell = cell;
                StartColumn = startColumn;
                EndColumn = endColumn;
            }

            internal TableCell Cell { get; }

            internal int StartColumn { get; }

            internal int EndColumn { get; }

            internal bool ContainsColumn(int columnIndex) {
                return columnIndex >= StartColumn && columnIndex < EndColumn;
            }
        }

        private sealed class FormulaExpressionParser {
            private readonly string _expression;
            private int _position;

            private FormulaExpressionParser(string expression) {
                _expression = expression;
            }

            internal static bool TryEvaluate(string expression, out decimal result, out string? diagnostic) {
                result = 0m;
                diagnostic = null;

                try {
                    var parser = new FormulaExpressionParser(expression);
                    result = parser.ParseComparison();
                    parser.SkipWhitespace();

                    if (!parser.IsAtEnd) {
                        diagnostic = $"Formula expression contains unsupported token '{parser.CurrentChar}'.";
                        return false;
                    }

                    return true;
                } catch (InvalidOperationException ex) {
                    diagnostic = ex.Message;
                    return false;
                } catch (DivideByZeroException) {
                    diagnostic = "Formula expression divides by zero.";
                    return false;
                } catch (OverflowException) {
                    diagnostic = "Formula expression is outside the supported decimal range.";
                    return false;
                }
            }

            private bool IsAtEnd => _position >= _expression.Length;

            private char CurrentChar => IsAtEnd ? '\0' : _expression[_position];

            private decimal ParseComparison() {
                decimal value = ParseExpression();

                while (true) {
                    SkipWhitespace();

                    if (TryConsume("<>")) {
                        value = value != ParseExpression() ? 1m : 0m;
                    } else if (TryConsume("<=")) {
                        value = value <= ParseExpression() ? 1m : 0m;
                    } else if (TryConsume(">=")) {
                        value = value >= ParseExpression() ? 1m : 0m;
                    } else if (TryConsume("=")) {
                        value = value == ParseExpression() ? 1m : 0m;
                    } else if (TryConsume("<")) {
                        value = value < ParseExpression() ? 1m : 0m;
                    } else if (TryConsume(">")) {
                        value = value > ParseExpression() ? 1m : 0m;
                    } else {
                        return value;
                    }
                }
            }

            private decimal ParseExpression() {
                decimal value = ParseTerm();

                while (true) {
                    SkipWhitespace();

                    if (TryConsume('+')) {
                        value += ParseTerm();
                    } else if (TryConsume('-')) {
                        value -= ParseTerm();
                    } else {
                        return value;
                    }
                }
            }

            private decimal ParseTerm() {
                decimal value = ParsePower();

                while (true) {
                    SkipWhitespace();

                    if (TryConsume('*')) {
                        value *= ParsePower();
                    } else if (TryConsume('/')) {
                        decimal divisor = ParsePower();
                        if (divisor == 0m) {
                            throw new DivideByZeroException();
                        }

                        value /= divisor;
                    } else {
                        return value;
                    }
                }
            }

            private decimal ParsePower() {
                decimal value = ParseUnary();
                SkipWhitespace();

                if (!TryConsume('^')) {
                    return value;
                }

                decimal exponent = ParsePower();
                if (decimal.Truncate(exponent) != exponent) {
                    throw new InvalidOperationException("Formula exponent must be an integer.");
                }

                return DecimalPower(value, exponent);
            }

            private decimal ParseUnary() {
                SkipWhitespace();

                if (TryConsume('+')) {
                    return ParseUnary();
                }

                if (TryConsume('-')) {
                    return -ParseUnary();
                }

                return ParsePostfix();
            }

            private decimal ParsePostfix() {
                decimal value = ParsePrimary();
                SkipWhitespace();

                while (TryConsume('%')) {
                    value /= 100m;
                    SkipWhitespace();
                }

                return value;
            }

            private decimal ParsePrimary() {
                SkipWhitespace();

                if (TryConsume('(')) {
                    decimal value = ParseComparison();
                    SkipWhitespace();

                    if (!TryConsume(')')) {
                        throw new InvalidOperationException("Formula expression has an unclosed parenthesis.");
                    }

                    return value;
                }

                if (char.IsLetter(CurrentChar)) {
                    return ParseFunction();
                }

                return ParseNumber();
            }

            private decimal ParseFunction() {
                string functionName = ParseIdentifier();
                SkipWhitespace();

                if (!TryConsume('(')) {
                    if (string.Equals(functionName, "TRUE", StringComparison.OrdinalIgnoreCase)) {
                        return 1m;
                    }

                    if (string.Equals(functionName, "FALSE", StringComparison.OrdinalIgnoreCase)) {
                        return 0m;
                    }

                    throw new InvalidOperationException($"Formula expression contains unsupported formula function or reference '{functionName}'.");
                }

                if (string.Equals(functionName, "DEFINED", StringComparison.OrdinalIgnoreCase)) {
                    return ParseDefinedFunction();
                }

                if (string.Equals(functionName, "IF", StringComparison.OrdinalIgnoreCase)) {
                    return ParseIfFunction();
                }

                if (string.Equals(functionName, "AND", StringComparison.OrdinalIgnoreCase)) {
                    return ParseAndFunction();
                }

                if (string.Equals(functionName, "OR", StringComparison.OrdinalIgnoreCase)) {
                    return ParseOrFunction();
                }

                var arguments = new List<decimal>();
                SkipWhitespace();
                if (TryConsume(')')) {
                    return EvaluateFunction(functionName, arguments);
                }

                while (true) {
                    arguments.Add(ParseComparison());
                    SkipWhitespace();

                    if (TryConsume(')')) {
                        break;
                    }

                    if (!TryConsume(',') && !TryConsume(';')) {
                        throw new InvalidOperationException($"Formula function {functionName} expected ',', ';', or ')' after an argument.");
                    }
                }

                return EvaluateFunction(functionName, arguments);
            }

            private decimal ParseIfFunction() {
                SkipWhitespace();
                if (CurrentChar == ')') {
                    throw new InvalidOperationException("Formula function IF requires three expression arguments.");
                }

                decimal condition = ParseComparison();
                SkipWhitespace();
                if (!TryConsumeArgumentSeparator()) {
                    throw new InvalidOperationException("Formula function IF expected ',', or ';' after the condition argument.");
                }

                if (IsTrue(condition)) {
                    decimal trueValue = ParseComparison();
                    SkipWhitespace();
                    if (!TryConsumeArgumentSeparator()) {
                        throw new InvalidOperationException("Formula function IF expected ',', or ';' after the true-result argument.");
                    }

                    if (!TrySkipArgumentToFunctionEnd()) {
                        throw new InvalidOperationException("Formula function IF expected ')' after the false-result argument.");
                    }

                    return trueValue;
                }

                if (!TrySkipArgumentToSeparator()) {
                    throw new InvalidOperationException("Formula function IF requires a false-result argument.");
                }

                decimal falseValue = ParseComparison();
                SkipWhitespace();
                if (!TryConsume(')')) {
                    throw new InvalidOperationException("Formula function IF expected ')' after the false-result argument.");
                }

                return falseValue;
            }

            private decimal ParseAndFunction() {
                SkipWhitespace();
                if (CurrentChar == ')') {
                    throw new InvalidOperationException("Formula function AND requires at least one expression argument.");
                }

                while (true) {
                    decimal value = ParseComparison();
                    SkipWhitespace();

                    if (!IsTrue(value)) {
                        if (TryConsume(')')) {
                            return 0m;
                        }

                        if (!TryConsumeArgumentSeparator() || !TrySkipRemainingArgumentsToFunctionEnd()) {
                            throw new InvalidOperationException("Formula function AND expected ',', ';', or ')' after an argument.");
                        }

                        return 0m;
                    }

                    if (TryConsume(')')) {
                        return 1m;
                    }

                    if (!TryConsumeArgumentSeparator()) {
                        throw new InvalidOperationException("Formula function AND expected ',', ';', or ')' after an argument.");
                    }
                }
            }

            private decimal ParseOrFunction() {
                SkipWhitespace();
                if (CurrentChar == ')') {
                    throw new InvalidOperationException("Formula function OR requires at least one expression argument.");
                }

                while (true) {
                    decimal value = ParseComparison();
                    SkipWhitespace();

                    if (IsTrue(value)) {
                        if (TryConsume(')')) {
                            return 1m;
                        }

                        if (!TryConsumeArgumentSeparator() || !TrySkipRemainingArgumentsToFunctionEnd()) {
                            throw new InvalidOperationException("Formula function OR expected ',', ';', or ')' after an argument.");
                        }

                        return 1m;
                    }

                    if (TryConsume(')')) {
                        return 0m;
                    }

                    if (!TryConsumeArgumentSeparator()) {
                        throw new InvalidOperationException("Formula function OR expected ',', ';', or ')' after an argument.");
                    }
                }
            }

            private decimal ParseDefinedFunction() {
                SkipWhitespace();
                if (CurrentChar == ')') {
                    throw new InvalidOperationException("Formula function DEFINED requires one expression argument.");
                }

                int argumentStart = _position;
                try {
                    _ = ParseComparison();
                    SkipWhitespace();

                    if (TryConsume(')')) {
                        return 1m;
                    }

                    if (CurrentChar == ',' || CurrentChar == ';') {
                        throw new InvalidOperationException("Formula function DEFINED requires one expression argument.");
                    }

                    throw new InvalidOperationException($"Formula function DEFINED expected ')' after its argument.");
                } catch (InvalidOperationException) {
                    _position = argumentStart;
                    if (TrySkipDefinedArgument()) {
                        return 0m;
                    }

                    throw;
                } catch (DivideByZeroException) {
                    _position = argumentStart;
                    if (TrySkipDefinedArgument()) {
                        return 0m;
                    }

                    throw;
                } catch (OverflowException) {
                    _position = argumentStart;
                    if (TrySkipDefinedArgument()) {
                        return 0m;
                    }

                    throw;
                }
            }

            private bool TrySkipDefinedArgument() {
                int depth = 0;
                while (!IsAtEnd) {
                    if (CurrentChar == '(') {
                        depth++;
                        _position++;
                        continue;
                    }

                    if (CurrentChar == ')') {
                        if (depth == 0) {
                            _position++;
                            return true;
                        }

                        depth--;
                        _position++;
                        continue;
                    }

                    if (depth == 0 && (CurrentChar == ',' || CurrentChar == ';')) {
                        return false;
                    }

                    _position++;
                }

                return false;
            }

            private bool TryConsumeArgumentSeparator() {
                return TryConsume(',') || TryConsume(';');
            }

            private bool TrySkipArgumentToSeparator() {
                int depth = 0;
                while (!IsAtEnd) {
                    if (CurrentChar == '(') {
                        depth++;
                        _position++;
                        continue;
                    }

                    if (CurrentChar == ')') {
                        if (depth == 0) {
                            return false;
                        }

                        depth--;
                        _position++;
                        continue;
                    }

                    if (depth == 0 && (CurrentChar == ',' || CurrentChar == ';')) {
                        _position++;
                        return true;
                    }

                    _position++;
                }

                return false;
            }

            private bool TrySkipArgumentToFunctionEnd() {
                int depth = 0;
                while (!IsAtEnd) {
                    if (CurrentChar == '(') {
                        depth++;
                        _position++;
                        continue;
                    }

                    if (CurrentChar == ')') {
                        if (depth == 0) {
                            _position++;
                            return true;
                        }

                        depth--;
                        _position++;
                        continue;
                    }

                    if (depth == 0 && (CurrentChar == ',' || CurrentChar == ';')) {
                        return false;
                    }

                    _position++;
                }

                return false;
            }

            private bool TrySkipRemainingArgumentsToFunctionEnd() {
                bool hasArgumentContent = false;
                int depth = 0;

                while (!IsAtEnd) {
                    if (CurrentChar == '(') {
                        hasArgumentContent = true;
                        depth++;
                        _position++;
                        continue;
                    }

                    if (CurrentChar == ')') {
                        if (depth == 0) {
                            if (!hasArgumentContent) {
                                return false;
                            }

                            _position++;
                            return true;
                        }

                        depth--;
                        _position++;
                        continue;
                    }

                    if (depth == 0 && (CurrentChar == ',' || CurrentChar == ';')) {
                        if (!hasArgumentContent) {
                            return false;
                        }

                        hasArgumentContent = false;
                        _position++;
                        continue;
                    }

                    if (!char.IsWhiteSpace(CurrentChar)) {
                        hasArgumentContent = true;
                    }

                    _position++;
                }

                return false;
            }

            private decimal ParseNumber() {
                SkipWhitespace();
                int start = _position;

                while (!IsAtEnd && (char.IsDigit(CurrentChar) || CurrentChar == '.')) {
                    _position++;
                }

                if (start == _position) {
                    throw new InvalidOperationException($"Formula expression expected a number near '{CurrentChar}'.");
                }

                string token = _expression.Substring(start, _position - start);
                if (!decimal.TryParse(token, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal value)) {
                    throw new InvalidOperationException($"Formula expression number '{token}' could not be parsed.");
                }

                return value;
            }

            private string ParseIdentifier() {
                int start = _position;
                while (!IsAtEnd && (char.IsLetter(CurrentChar) || CurrentChar == '_')) {
                    _position++;
                }

                return _expression.Substring(start, _position - start);
            }

            private bool TryConsume(char expected) {
                if (CurrentChar != expected) {
                    return false;
                }

                _position++;
                return true;
            }

            private bool TryConsume(string expected) {
                if (expected.Length == 0 || _position + expected.Length > _expression.Length) {
                    return false;
                }

                if (string.CompareOrdinal(_expression, _position, expected, 0, expected.Length) != 0) {
                    return false;
                }

                _position += expected.Length;
                return true;
            }

            private void SkipWhitespace() {
                while (!IsAtEnd && char.IsWhiteSpace(CurrentChar)) {
                    _position++;
                }
            }

            private static decimal DecimalPower(decimal value, decimal exponent) {
                int power = checked((int)exponent);
                if (power == 0) {
                    return 1m;
                }

                bool negativeExponent = power < 0;
                long remaining = Math.Abs((long)power);
                decimal result = 1m;
                decimal factor = value;
                while (remaining > 0) {
                    if ((remaining & 1L) != 0) result *= factor;
                    remaining >>= 1;
                    if (remaining > 0) factor *= factor;
                }

                return negativeExponent ? 1m / result : result;
            }

            private static decimal EvaluateFunction(string functionName, IReadOnlyList<decimal> arguments) {
                switch (functionName.ToUpperInvariant()) {
                    case "SUM":
                        RequireAtLeastOneArgument(functionName, arguments);
                        return arguments.Sum();
                    case "AVERAGE":
                        RequireAtLeastOneArgument(functionName, arguments);
                        return arguments.Sum() / arguments.Count;
                    case "MIN":
                        RequireAtLeastOneArgument(functionName, arguments);
                        return arguments.Min();
                    case "MAX":
                        RequireAtLeastOneArgument(functionName, arguments);
                        return arguments.Max();
                    case "PRODUCT":
                        RequireAtLeastOneArgument(functionName, arguments);
                        return arguments.Aggregate(1m, (current, value) => current * value);
                    case "COUNT":
                        RequireAtLeastOneArgument(functionName, arguments);
                        return arguments.Count;
                    case "IF":
                        RequireArgumentCount(functionName, arguments, 3);
                        return IsTrue(arguments[0]) ? arguments[1] : arguments[2];
                    case "AND":
                        RequireAtLeastOneArgument(functionName, arguments);
                        return arguments.All(IsTrue) ? 1m : 0m;
                    case "OR":
                        RequireAtLeastOneArgument(functionName, arguments);
                        return arguments.Any(IsTrue) ? 1m : 0m;
                    case "NOT":
                        RequireArgumentCount(functionName, arguments, 1);
                        return IsTrue(arguments[0]) ? 0m : 1m;
                    case "TRUE":
                        RequireArgumentCount(functionName, arguments, 0);
                        return 1m;
                    case "FALSE":
                        RequireArgumentCount(functionName, arguments, 0);
                        return 0m;
                    case "DEFINED":
                        RequireArgumentCount(functionName, arguments, 1);
                        return 1m;
                    case "MOD":
                        RequireArgumentCount(functionName, arguments, 2);
                        if (arguments[1] == 0m) {
                            throw new DivideByZeroException();
                        }

                        return arguments[0] % arguments[1];
                    case "SIGN":
                        RequireArgumentCount(functionName, arguments, 1);
                        return Math.Sign(arguments[0]);
                    case "ABS":
                        RequireArgumentCount(functionName, arguments, 1);
                        return Math.Abs(arguments[0]);
                    case "INT":
                        RequireArgumentCount(functionName, arguments, 1);
                        return Math.Floor(arguments[0]);
                    case "ROUND":
                        RequireArgumentCount(functionName, arguments, 2);
                        return Round(arguments[0], arguments[1]);
                    default:
                        throw new InvalidOperationException($"Formula function {functionName} is not supported by the bounded OfficeIMO evaluator.");
                }
            }

            private static bool IsTrue(decimal value) {
                return value != 0m;
            }

            private static void RequireAtLeastOneArgument(string functionName, IReadOnlyList<decimal> arguments) {
                if (arguments.Count == 0) {
                    throw new InvalidOperationException($"Formula function {functionName} requires at least one numeric argument.");
                }
            }

            private static void RequireArgumentCount(string functionName, IReadOnlyList<decimal> arguments, int expectedCount) {
                if (arguments.Count != expectedCount) {
                    throw new InvalidOperationException($"Formula function {functionName} requires {expectedCount.ToString(CultureInfo.InvariantCulture)} numeric argument(s).");
                }
            }

            private static decimal Round(decimal value, decimal digits) {
                if (decimal.Truncate(digits) != digits) {
                    throw new InvalidOperationException("Formula function ROUND requires an integer number of decimal places.");
                }

                int decimalPlaces = checked((int)digits);
                if (decimalPlaces < -28 || decimalPlaces > 28) {
                    throw new InvalidOperationException("Formula function ROUND supports decimal places from -28 to 28.");
                }

                if (decimalPlaces < 0) {
                    decimal factor = DecimalPower(10m, -decimalPlaces);
                    return Math.Round(value / factor, 0, MidpointRounding.AwayFromZero) * factor;
                }

                return Math.Round(value, decimalPlaces, MidpointRounding.AwayFromZero);
            }
        }
    }
}
