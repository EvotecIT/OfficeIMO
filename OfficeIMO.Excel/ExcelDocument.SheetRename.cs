using System;
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
            var definedNames = _workBookPart.Workbook.DefinedNames;
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
                _workBookPart.Workbook.Save();
            }
        }

        private void UpdateWorksheetReferences(string oldSheetName, string newSheetName) {
            foreach (var worksheetPart in _workBookPart.WorksheetParts) {
                bool changed = false;

                foreach (var formula in worksheetPart.Worksheet.Descendants<CellFormula>()) {
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

                foreach (var formula in worksheetPart.Worksheet.Descendants<Formula>()) {
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

                foreach (var formula in worksheetPart.Worksheet.Descendants<Formula1>()) {
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

                foreach (var formula in worksheetPart.Worksheet.Descendants<Formula2>()) {
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

                foreach (var formula in worksheetPart.Worksheet.Descendants<OfficeFormula>()) {
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

                foreach (var hyperlink in worksheetPart.Worksheet.Descendants<Hyperlink>()) {
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
                    worksheetPart.Worksheet.Save();
                }
            }
        }

        private void UpdateChartReferences(string oldSheetName, string newSheetName) {
            foreach (var worksheetPart in _workBookPart.WorksheetParts) {
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
            foreach (var worksheetPart in _workBookPart.WorksheetParts) {
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
            foreach (var cacheDefinitionPart in _workBookPart.GetPartsOfType<PivotTableCacheDefinitionPart>()) {
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

            var cell = worksheetPart.Worksheet.Descendants<Cell>()
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

        private static string ReplaceSheetNameReferences(string text, string oldSheetName, string newSheetName) {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(oldSheetName) || string.IsNullOrEmpty(newSheetName)) {
                return text;
            }

            var builder = new StringBuilder(text.Length + 16);
            string replacement = $"'{EscapeSheetName(newSheetName)}'!";
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

                if (!inStringLiteral && ch == '\'') {
                    int closingQuote = FindQuotedTokenEnd(text, i);
                    if (closingQuote > i && closingQuote + 1 < text.Length && text[closingQuote + 1] == '!') {
                        string quotedName = text.Substring(i + 1, closingQuote - i - 1).Replace("''", "'");
                        if (!IsExternalSheetToken(quotedName) && string.Equals(quotedName, oldSheetName, StringComparison.OrdinalIgnoreCase)) {
                            builder.Append(replacement);
                            i = closingQuote + 2;
                            changed = true;
                            continue;
                        }
                    }
                }

                if (!inStringLiteral && IsBareSheetReferenceStart(text, i, oldSheetName)) {
                    builder.Append(replacement);
                    i += oldSheetName.Length + 1;
                    changed = true;
                    continue;
                }

                builder.Append(ch);
                i++;
            }

            return changed ? builder.ToString() : text;
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
            cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
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

        private static bool IsBareSheetReferenceStart(string text, int startIndex, string oldSheetName) {
            if (string.IsNullOrEmpty(oldSheetName)) {
                return false;
            }

            int endIndex = startIndex + oldSheetName.Length;
            if (endIndex >= text.Length) {
                return false;
            }

            if (!string.Equals(text.Substring(startIndex, oldSheetName.Length), oldSheetName, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (text[endIndex] != '!') {
                return false;
            }

            if (startIndex == 0) {
                return true;
            }

            char previous = text[startIndex - 1];
            if (char.IsLetterOrDigit(previous) || previous == '_' || previous == '.' || previous == '\'' || previous == ']') {
                return false;
            }

            return true;
        }
    }
}
