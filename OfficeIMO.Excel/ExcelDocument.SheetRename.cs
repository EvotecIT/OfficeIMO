using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ChartFormula = DocumentFormat.OpenXml.Drawing.Charts.Formula;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private void UpdateSheetNameReferences(string oldSheetName, string newSheetName) {
            if (string.IsNullOrWhiteSpace(oldSheetName) || string.IsNullOrWhiteSpace(newSheetName)) {
                return;
            }

            if (string.Equals(oldSheetName, newSheetName, StringComparison.Ordinal)) {
                return;
            }

            UpdateDefinedNameReferences(oldSheetName, newSheetName);
            UpdateWorksheetReferences(oldSheetName, newSheetName);
            UpdateTableReferences(oldSheetName, newSheetName);
            UpdateChartReferences(oldSheetName, newSheetName);
            UpdatePivotCacheReferences(oldSheetName, newSheetName);
        }

        private void UpdateDefinedNameReferences(string oldSheetName, string newSheetName) {
            var definedNames = WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return;
            }

            bool changed = false;
            foreach (var definedName in definedNames.Elements<DefinedName>()) {
                string? text = definedName.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                string updated = ReplaceSheetNameReferences(text, oldSheetName, newSheetName);
                if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                    definedName.Text = updated;
                    changed = true;
                }
            }

            if (changed) {
                WorkbookRoot.Save();
            }
        }

        private void UpdateWorksheetReferences(string oldSheetName, string newSheetName) {
            foreach (var worksheetPart in WorkbookPartRoot.WorksheetParts) {
                var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
                bool changed = false;

                foreach (var formula in worksheet.Descendants<CellFormula>()) {
                    string? text = formula.Text;
                    if (string.IsNullOrEmpty(text)) {
                        continue;
                    }

                    string updated = ReplaceSheetNameReferences(text, oldSheetName, newSheetName);
                    if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                        formula.Text = updated;
                        changed = true;
                    }
                }

                foreach (var formula in worksheet.Descendants<Formula>()) {
                    string? text = formula.Text;
                    if (string.IsNullOrEmpty(text)) {
                        continue;
                    }

                    string updated = ReplaceSheetNameReferences(text, oldSheetName, newSheetName);
                    if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                        formula.Text = updated;
                        changed = true;
                    }
                }

                foreach (var formula in worksheet.Descendants<Formula1>()) {
                    string? text = formula.Text;
                    if (string.IsNullOrEmpty(text)) {
                        continue;
                    }

                    string updated = ReplaceSheetNameReferences(text, oldSheetName, newSheetName);
                    if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                        formula.Text = updated;
                        changed = true;
                    }
                }

                foreach (var formula in worksheet.Descendants<Formula2>()) {
                    string? text = formula.Text;
                    if (string.IsNullOrEmpty(text)) {
                        continue;
                    }

                    string updated = ReplaceSheetNameReferences(text, oldSheetName, newSheetName);
                    if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                        formula.Text = updated;
                        changed = true;
                    }
                }

                foreach (var formula in worksheet.Descendants<OfficeFormula>()) {
                    string? text = formula.Text;
                    if (string.IsNullOrEmpty(text)) {
                        continue;
                    }

                    string updated = ReplaceSheetNameReferences(text, oldSheetName, newSheetName);
                    if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                        formula.Text = updated;
                        changed = true;
                    }
                }

                foreach (var hyperlink in worksheet.Descendants<Hyperlink>()) {
                    string? location = hyperlink.Location?.Value;
                    if (string.IsNullOrEmpty(location)) {
                        continue;
                    }

                    string updated = ReplaceSheetNameReferences(location!, oldSheetName, newSheetName);
                    if (!string.Equals(updated, location, StringComparison.Ordinal)) {
                        hyperlink.Location = updated;
                        TryRefreshInternalHyperlinkDisplayText(worksheetPart, hyperlink, oldSheetName, newSheetName, ref changed);
                        changed = true;
                    }
                }

                if (changed) {
                    worksheet.Save();
                }
            }
        }

        private void UpdateChartReferences(string oldSheetName, string newSheetName) {
            foreach (var worksheetPart in WorkbookPartRoot.WorksheetParts) {
                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart?.WorksheetDrawing == null) {
                    continue;
                }

                foreach (var chartPart in drawingsPart.ChartParts) {
                    var chartSpace = chartPart.ChartSpace;
                    if (chartSpace == null) {
                        continue;
                    }

                    bool changed = false;
                    foreach (var formula in chartSpace.Descendants<ChartFormula>()) {
                        string? text = formula.Text;
                        if (string.IsNullOrEmpty(text)) {
                            continue;
                        }

                        string updated = ReplaceSheetNameReferences(text, oldSheetName, newSheetName);
                        if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                            formula.Text = updated;
                            changed = true;
                        }
                    }

                    if (changed) {
                        chartSpace.Save();
                    }
                }
            }
        }

        private void UpdateTableReferences(string oldSheetName, string newSheetName) {
            foreach (var worksheetPart in WorkbookPartRoot.WorksheetParts) {
                foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts) {
                    var table = tableDefinitionPart.Table;
                    if (table == null) {
                        continue;
                    }

                    bool changed = false;

                    foreach (var formula in table.Descendants<CalculatedColumnFormula>()) {
                        string? text = formula.Text;
                        if (string.IsNullOrEmpty(text)) {
                            continue;
                        }

                        string updated = ReplaceSheetNameReferences(text, oldSheetName, newSheetName);
                        if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                            formula.Text = updated;
                            changed = true;
                        }
                    }

                    foreach (var formula in table.Descendants<TotalsRowFormula>()) {
                        string? text = formula.Text;
                        if (string.IsNullOrEmpty(text)) {
                            continue;
                        }

                        string updated = ReplaceSheetNameReferences(text, oldSheetName, newSheetName);
                        if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                            formula.Text = updated;
                            changed = true;
                        }
                    }

                    if (changed) {
                        table.Save();
                    }
                }
            }
        }
        private void UpdatePivotCacheReferences(string oldSheetName, string newSheetName) {
            foreach (var cacheDefinitionPart in WorkbookPartRoot.GetPartsOfType<PivotTableCacheDefinitionPart>()) {
                var worksheetSource = cacheDefinitionPart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                string? sourceSheet = worksheetSource?.Sheet?.Value;
                if (worksheetSource == null || !string.Equals(sourceSheet, oldSheetName, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                worksheetSource.Sheet = newSheetName;
                cacheDefinitionPart.PivotCacheDefinition!.Save();
            }
        }

        private void TryRefreshInternalHyperlinkDisplayText(
            WorksheetPart worksheetPart,
            Hyperlink hyperlink,
            string oldSheetName,
            string newSheetName,
            ref bool changed) {
            string? cellReference = hyperlink.Reference?.Value;
            if (string.IsNullOrEmpty(cellReference)) {
                return;
            }

            var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            var cell = worksheet.Descendants<Cell>()
                .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));
            if (cell == null) {
                return;
            }

            string currentText = GetCellText(cell);
            if (string.IsNullOrEmpty(currentText)) {
                return;
            }

            string? replacementText = null;
            if (string.Equals(currentText, oldSheetName, StringComparison.OrdinalIgnoreCase)) {
                replacementText = newSheetName;
            } else if (string.Equals(currentText, $"← {oldSheetName}", StringComparison.OrdinalIgnoreCase)) {
                replacementText = $"← {newSheetName}";
            }

            if (replacementText == null || string.Equals(currentText, replacementText, StringComparison.Ordinal)) {
                return;
            }

            SetCellText(cell, replacementText);
            changed = true;
        }

        internal static string ReplaceSheetNameReferences(string text, string oldSheetName, string newSheetName) {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(oldSheetName) || string.IsNullOrEmpty(newSheetName)) {
                return text;
            }

            return ReplaceSheetNameReferences(text, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
                [oldSheetName] = newSheetName
            });
        }

        internal static string ReplaceSheetNameReferences(string text, IReadOnlyDictionary<string, string> sheetNameMap) {
            if (string.IsNullOrEmpty(text) || sheetNameMap.Count == 0) {
                return text;
            }

            var builder = new StringBuilder(text.Length + 16);
            bool changed = false;
            bool inStringLiteral = false;

            for (int i = 0; i < text.Length;) {
                char ch = text[i];

                if (ch == '"') {
                    builder.Append(ch);
                    if (inStringLiteral && i + 1 < text.Length && text[i + 1] == '"') {
                        builder.Append('"');
                        i += 2;
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    i++;
                    continue;
                }

                if (!inStringLiteral && TryRewriteSheetReferenceAt(text, i, sheetNameMap, out string? replacement, out int consumed)) {
                    builder.Append(replacement);
                    i += consumed;
                    changed = true;
                    continue;
                }

                builder.Append(ch);
                i++;
            }

            return changed ? builder.ToString() : text;
        }

        private static bool TryRewriteSheetReferenceAt(
            string text,
            int startIndex,
            IReadOnlyDictionary<string, string> sheetNameMap,
            out string? replacement,
            out int consumed) {
            replacement = null;
            consumed = 0;

            if (!TryReadSheetToken(text, startIndex, out string? firstSheetName, out int afterFirstToken, out bool firstExternal)) {
                return false;
            }

            if (afterFirstToken < text.Length && text[afterFirstToken] == ':'
                && TryReadSheetToken(text, afterFirstToken + 1, out string? secondSheetName, out int afterSecondToken, out bool secondExternal)
                && afterSecondToken < text.Length
                && text[afterSecondToken] == '!'
                && !firstExternal
                && !secondExternal) {
                string rewrittenFirst = ResolveMappedSheetName(firstSheetName!, sheetNameMap, out bool firstChanged);
                string rewrittenSecond = ResolveMappedSheetName(secondSheetName!, sheetNameMap, out bool secondChanged);
                if (!firstChanged && !secondChanged) {
                    return false;
                }

                replacement = QuoteSheetRangeReference(rewrittenFirst, rewrittenSecond) + "!";
                consumed = afterSecondToken - startIndex + 1;
                return true;
            }

            if (afterFirstToken >= text.Length || text[afterFirstToken] != '!' || firstExternal) {
                return false;
            }

            if (TrySplitQuotedThreeDimensionalSheetRange(firstSheetName!, out string? firstRangeSheet, out string? secondRangeSheet)) {
                string rewrittenFirst = ResolveMappedSheetName(firstRangeSheet!, sheetNameMap, out bool firstChanged);
                string rewrittenSecond = ResolveMappedSheetName(secondRangeSheet!, sheetNameMap, out bool secondChanged);
                if (!firstChanged && !secondChanged) {
                    return false;
                }

                replacement = QuoteSheetRangeReference(rewrittenFirst, rewrittenSecond) + "!";
                consumed = afterFirstToken - startIndex + 1;
                return true;
            }

            string rewritten = ResolveMappedSheetName(firstSheetName!, sheetNameMap, out bool changed);
            if (!changed) {
                return false;
            }

            replacement = QuoteSheetNameReference(rewritten) + "!";
            consumed = afterFirstToken - startIndex + 1;
            return true;
        }

        private static bool TrySplitQuotedThreeDimensionalSheetRange(string sheetToken, out string? firstSheetName, out string? secondSheetName) {
            firstSheetName = null;
            secondSheetName = null;
            int separator = sheetToken.IndexOf(':');
            if (separator <= 0 || separator >= sheetToken.Length - 1) {
                return false;
            }

            firstSheetName = sheetToken.Substring(0, separator);
            secondSheetName = sheetToken.Substring(separator + 1);
            return !IsExternalSheetToken(firstSheetName) && !IsExternalSheetToken(secondSheetName);
        }

        private static bool TryReadSheetToken(
            string text,
            int startIndex,
            out string? sheetName,
            out int afterToken,
            out bool isExternal) {
            sheetName = null;
            afterToken = startIndex;
            isExternal = false;

            if (startIndex < 0 || startIndex >= text.Length) {
                return false;
            }

            if (text[startIndex] == '\'') {
                int closingQuote = FindQuotedTokenEnd(text, startIndex);
                if (closingQuote <= startIndex) {
                    return false;
                }

                string quotedName = text.Substring(startIndex + 1, closingQuote - startIndex - 1).Replace("''", "'");
                sheetName = quotedName;
                afterToken = closingQuote + 1;
                isExternal = IsExternalSheetToken(quotedName);
                return true;
            }

            if (!IsBareSheetReferenceBoundary(text, startIndex)) {
                return false;
            }

            int index = startIndex;
            while (index < text.Length && IsBareSheetNameCharacter(text[index])) {
                index++;
            }

            if (index == startIndex) {
                return false;
            }

            sheetName = text.Substring(startIndex, index - startIndex);
            afterToken = index;
            isExternal = IsExternalSheetToken(sheetName);
            return true;
        }

        private static string ResolveMappedSheetName(string sheetName, IReadOnlyDictionary<string, string> sheetNameMap, out bool changed) {
            if (sheetNameMap.TryGetValue(sheetName, out string? mapped) && !string.IsNullOrEmpty(mapped)) {
                changed = !string.Equals(sheetName, mapped, StringComparison.Ordinal);
                return mapped!;
            }

            changed = false;
            return sheetName;
        }

        private static string QuoteSheetNameReference(string sheetName) {
            return $"'{EscapeSheetName(sheetName)}'";
        }

        private static string QuoteSheetRangeReference(string firstSheetName, string secondSheetName) {
            return $"'{EscapeSheetName(firstSheetName)}:{EscapeSheetName(secondSheetName)}'";
        }

        private static bool IsExternalSheetToken(string sheetToken) {
            return !string.IsNullOrEmpty(sheetToken) && sheetToken.IndexOf(']') >= 0;
        }

        private string GetCellText(Cell cell) {
            if (cell.DataType?.Value == CellValues.SharedString) {
                var raw = cell.CellValue?.InnerText;
                if (!string.IsNullOrEmpty(raw) && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id)) {
                    var sst = SharedStringTablePart?.SharedStringTable;
                    if (sst != null) {
                        var item = sst.Elements<SharedStringItem>().ElementAtOrDefault(id);
                        if (item != null) {
                            if (item.Text != null) {
                                return item.Text.Text ?? string.Empty;
                            }

                            var text = new StringBuilder();
                            foreach (var runText in item.Descendants<Text>()) {
                                text.Append(runText.Text);
                            }
                            return text.ToString();
                        }
                    }
                }
                return string.Empty;
            }

            if (cell.DataType?.Value == CellValues.InlineString) {
                var inline = cell.InlineString;
                if (inline?.Text != null) {
                    return inline.Text.Text ?? string.Empty;
                }

                if (inline != null) {
                    var text = new StringBuilder();
                    foreach (var run in inline.Elements<Run>()) {
                        if (run.Text != null) {
                            text.Append(run.Text.Text);
                        }
                    }
                    return text.ToString();
                }
            }

            return cell.CellValue?.InnerText ?? string.Empty;
        }

        private void SetCellText(Cell cell, string text) {
            CoerceValueHelper.ValidateSharedStringLength(text, nameof(text));
            int index = GetSharedStringIndex(text);
            cell.InlineString = null;
            cell.CellValue = new CellValue(SharedStringIndexText.Get(index));
            cell.DataType = CellValues.SharedString;
        }

        private static int FindQuotedTokenEnd(string text, int startIndex) {
            for (int i = startIndex + 1; i < text.Length; i++) {
                if (text[i] != '\'') {
                    continue;
                }

                if (i + 1 < text.Length && text[i + 1] == '\'') {
                    i++;
                    continue;
                }

                return i;
            }

            return -1;
        }

        private static bool IsBareSheetReferenceBoundary(string text, int startIndex) {
            if (startIndex == 0) {
                return true;
            }

            char previous = text[startIndex - 1];
            if (char.IsLetterOrDigit(previous) || previous == '_' || previous == '.' || previous == '\'' || previous == ']') {
                return false;
            }

            return true;
        }

        private static bool IsBareSheetNameCharacter(char value) {
            return char.IsLetterOrDigit(value) || value == '_' || value == '.';
        }
    }
}
