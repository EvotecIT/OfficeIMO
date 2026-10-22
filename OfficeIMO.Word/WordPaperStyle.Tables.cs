using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public sealed partial class WordPaperStyle {
    /// <summary>Applies horizontal hairlines, a muted header, and matched header/value alignment for numeric columns. Cell fills are retained.</summary>
    /// <param name="table">The table to style.</param>
    /// <param name="hasHeader">Whether the first row contains column labels.</param>
    /// <param name="allowRowSplit">Whether long evidence rows may continue on another page.</param>
    /// <param name="alignNumericColumns">Whether columns containing only numbers should align right, including their headers.</param>
    public void ApplyTable(WordTable table, bool hasHeader = true, bool allowRowSplit = false, bool alignNumericColumns = true) {
        if (table == null) throw new ArgumentNullException(nameof(table));
        table.Style = WordTableStyle.TableNormal;
        table.WidthType = WordTableWidthUnit.Pct;
        table.Width = 5000;
        table.RepeatHeaderRowAtTheTopOfEachPage = hasHeader;
        table.AllowRowToBreakAcrossPages = allowRowSplit;
        table.StyleDetails!.SetCustomBorders(topStyle: WordBorderStyle.Nil, bottomStyle: WordBorderStyle.Nil,
            leftStyle: WordBorderStyle.Nil, rightStyle: WordBorderStyle.Nil,
            insideHStyle: WordBorderStyle.Single, insideHSize: RuleSize,
            insideHColor: OfficeIMO.Drawing.OfficeColor.ParseHex(Border), insideVStyle: WordBorderStyle.Nil);
        table.StyleDetails.MarginDefaultTopWidth = TableCellVerticalPadding;
        table.StyleDetails.MarginDefaultBottomWidth = TableCellVerticalPadding;
        table.StyleDetails.MarginDefaultLeftWidth = TableCellHorizontalPadding;
        table.StyleDetails.MarginDefaultRightWidth = TableCellHorizontalPadding;
        int columnCount = table.Rows.Count == 0 ? 0 : table.Rows[0].Cells.Count;
        var numericColumns = new HashSet<int>();
        if (alignNumericColumns) {
            for (int column = 0; column < columnCount; column++) {
                string[] values = table.Rows.Skip(hasHeader ? 1 : 0).Where(row => row.Cells.Count > column)
                    .Select(row => string.Join(" ", row.Cells[column].Paragraphs.Select(p => p.Text))).Where(text => !string.IsNullOrWhiteSpace(text)).ToArray();
                if (values.Length > 0 && values.All(IsNumeric)) numericColumns.Add(column);
            }
        }
        for (int row = 0; row < table.Rows.Count; row++) {
            for (int column = 0; column < table.Rows[row].Cells.Count; column++) {
                WordTableCell cell = table.Rows[row].Cells[column];
                // Explicit cell borders can otherwise override the table's border contract.
                cell._tableCell.TableCellProperties?.GetFirstChild<TableCellBorders>()?.Remove();
                bool header = hasHeader && row == 0;
                if (header) cell.ShadingFillColorHex = Surface;
                foreach (WordParagraph paragraph in cell.Paragraphs) {
                    paragraph.FontFamily = FontFamily;
                    paragraph.FontSizePoints = header ? 9 : TableFontSizePoints;
                    paragraph.Bold = header;
                    paragraph.ColorHex = header ? Muted : Ink;
                    paragraph.LineSpacingBeforePoints = TableParagraphSpacingPoints;
                    paragraph.LineSpacingAfterPoints = TableParagraphSpacingPoints;
                    if (TableLineSpacingPoints.HasValue) {
                        paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
                        paragraph.LineSpacingPoints = TableLineSpacingPoints.Value;
                    }
                    paragraph.ParagraphAlignment = numericColumns.Contains(column) ? WordParagraphAlignment.Right : WordParagraphAlignment.Left;
                    if (!header && numericColumns.Contains(column)) {
                        paragraph.ParagraphAlignment = WordParagraphAlignment.Right;
                        foreach (Run run in paragraph._paragraph.Elements<Run>()) {
                            RunProperties properties = run.RunProperties ??= new RunProperties();
                            properties.RemoveAllChildren<DocumentFormat.OpenXml.Office2010.Word.NumberSpacing>();
                            properties.AddChild(new DocumentFormat.OpenXml.Office2010.Word.NumberSpacing {
                                Val = DocumentFormat.OpenXml.Office2010.Word.NumberSpacingValues.Tabular
                            }, true);
                        }
                    }
                }
            }
        }
    }

    /// <summary>Adds a two-column evidence table using the paper rules and preserving multiline values.</summary>
    /// <param name="document">The destination document.</param>
    /// <param name="rows">Labels and their evidence text, in display order.</param>
    /// <param name="keyHeader">First column label.</param>
    /// <param name="valueHeader">Second column label.</param>
    /// <param name="allowRowSplit">Whether long evidence may continue on another page.</param>
    /// <param name="hasHeader">Whether to include a header row; false is useful for cover and definition tables.</param>
    /// <returns>The styled evidence table.</returns>
    public WordTable AddKeyValueTable(WordDocument document, IEnumerable<KeyValuePair<string, string>> rows,
        string keyHeader = "Field", string valueHeader = "Value", bool allowRowSplit = false, bool hasHeader = true) {
        var values = rows.ToList();
        int offset = hasHeader ? 1 : 0;
        WordTable table = document.AddTable(Math.Max(1, values.Count + offset), 2, WordTableStyle.TableNormal);
        if (hasHeader) {
            SetCellText(table.Rows[0].Cells[0], keyHeader);
            SetCellText(table.Rows[0].Cells[1], valueHeader);
        }
        for (int row = 0; row < values.Count; row++) {
            SetCellText(table.Rows[row + offset].Cells[0], values[row].Key);
            SetCellText(table.Rows[row + offset].Cells[1], values[row].Value);
        }
        ApplyKeyValueTable(table, hasHeader: hasHeader, allowRowSplit: allowRowSplit);
        if (!hasHeader) {
            foreach (WordTableRow row in table.Rows) ApplyCellTone(row.Cells[0], Muted);
        }
        return table;
    }

    /// <summary>Styles a key/value or definition table with all labels and values aligned left, including numbers.</summary>
    public void ApplyKeyValueTable(WordTable table, bool hasHeader = true, bool allowRowSplit = false)
        => ApplyTable(table, hasHeader, allowRowSplit, alignNumericColumns: false);

    /// <summary>Styles a compact count table with content-sized columns and a right-aligned value column and header.</summary>
    public void ApplyCountTable(WordTable table, bool hasHeader = true) {
        ApplyTable(table, hasHeader);
        table.AutoFitToContents();
        foreach (WordTableRow row in table.Rows) {
            if (row.Cells.Count < 2) continue;
            foreach (WordParagraph paragraph in row.Cells[1].Paragraphs) paragraph.ParagraphAlignment = WordParagraphAlignment.Right;
        }
    }

    /// <summary>Adds a content-sized count table. Unknown values may be supplied as text alongside numeric counts.</summary>
    public WordTable AddCountTable(WordDocument document, IEnumerable<KeyValuePair<string, string>> rows,
        string keyHeader = "Metric", string valueHeader = "Count") {
        WordTable table = AddKeyValueTable(document, rows, keyHeader, valueHeader);
        ApplyCountTable(table);
        return table;
    }

}
