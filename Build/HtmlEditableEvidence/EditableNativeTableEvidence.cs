using System.Data;
using System.Globalization;
using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.OneNote;
using OfficeIMO.PowerPoint;
using OfficeIMO.Rtf;
using OfficeIMO.Word;

namespace OfficeIMO.Html.EditableEvidence;

internal sealed record NativeTableEvidence(string Location, string Kind, IReadOnlyList<string[]> Rows);

/// <summary>
/// Records cells from reopened native models, independently of their HTML export.
/// These observations do not accept a table merely because its text markers survived.
/// </summary>
internal static class EditableNativeTableEvidence {
    private const int MaximumCells = 100_000;

    internal static IReadOnlyList<NativeTableEvidence> FromWord(WordDocument document) =>
        document.TablesIncludingNestedTables.Select((table, index) => Snapshot($"table-{index + 1}", "word-table",
            table.Rows.Select(row => row.Cells.Select(cell => string.Join("\n", cell.Paragraphs.Select(p => p.Text)))))).ToArray();

    internal static IReadOnlyList<NativeTableEvidence> FromExcel(ExcelDocument document) =>
        document.Sheets.Select(sheet => {
            using DataTable cells = sheet.ToDataTable(headersInFirstRow: false,
                options: new ExcelReadOptions { MaxRangeCells = MaximumCells });
            return Snapshot(sheet.Name, "worksheet-grid", cells.Rows.Cast<DataRow>().Select(row =>
                row.ItemArray.Select(value => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)));
        }).ToArray();

    internal static IReadOnlyList<NativeTableEvidence> FromPowerPoint(PowerPointPresentation document) =>
        document.Slides.SelectMany((slide, slideIndex) => slide.Tables.Select((table, tableIndex) =>
            Snapshot($"slide-{slideIndex + 1}/table-{tableIndex + 1}", "powerpoint-table",
                Enumerable.Range(0, table.Rows).Select(row => Enumerable.Range(0, table.Columns)
                    .Select(column => table.GetCell(row, column).Text))))).ToArray();

    internal static IReadOnlyList<NativeTableEvidence> FromOneNote(OneNoteSection document) =>
        document.Pages.SelectMany((page, pageIndex) => OneNoteTables(page.DirectContent.Concat(page.Outlines))
            .Select((table, tableIndex) => Snapshot($"page-{pageIndex + 1}/table-{tableIndex + 1}", "onenote-table",
                table.Rows.Select(row => row.Cells.Select(cell => OneNoteText(cell.Content)))))).ToArray();

    internal static IReadOnlyList<NativeTableEvidence> FromRtf(RtfDocument document) =>
        RtfTables(document.Blocks).Select((table, index) => Snapshot($"table-{index + 1}", "rtf-table",
            table.Rows.Select(row => row.Cells.Select(cell =>
                string.Join("\n", cell.Paragraphs.Select(paragraph => paragraph.ToPlainText())))))).ToArray();

    internal static IReadOnlyList<NativeTableEvidence> FromMarkdown(MarkdownDoc document) =>
        document.Blocks.OfType<TableBlock>().Select((table, index) => Snapshot($"table-{index + 1}", "markdown-table",
            new[] { table.Headers.AsEnumerable() }.Concat(table.Rows.Select(row => row.AsEnumerable())))).ToArray();

    private static NativeTableEvidence Snapshot(string location, string kind, IEnumerable<IEnumerable<string>> rows) {
        var result = new List<string[]>();
        int cells = 0;
        foreach (IEnumerable<string> row in rows) {
            var values = new List<string>();
            foreach (string value in row) {
                if (++cells > MaximumCells) throw new InvalidDataException("Native table evidence exceeds the cell limit.");
                values.Add(value);
            }
            result.Add(values.ToArray());
        }
        return new NativeTableEvidence(location, kind, result);
    }

    private static IEnumerable<OneNoteTable> OneNoteTables(IEnumerable<OneNoteElement> elements) {
        foreach (OneNoteElement element in elements) {
            if (element is OneNoteTable table) {
                yield return table;
                foreach (OneNoteTable nested in OneNoteTables(table.Rows.SelectMany(row => row.Cells).SelectMany(cell => cell.Content)))
                    yield return nested;
            } else if (element is OneNoteOutline outline) {
                foreach (OneNoteTable nested in OneNoteTables(outline.Children)) yield return nested;
            } else if (element is OneNoteParagraph paragraph) {
                foreach (OneNoteTable nested in OneNoteTables(paragraph.Children)) yield return nested;
            }
        }
    }

    private static string OneNoteText(IEnumerable<OneNoteElement> elements) => string.Join("\n", elements.Select(element =>
        element is OneNoteParagraph paragraph ? string.Concat(paragraph.Runs.Select(run => run.Text))
            : element is OneNoteOutline outline ? OneNoteText(outline.Children) : string.Empty));

    private static IEnumerable<RtfTable> RtfTables(IEnumerable<IRtfBlock> blocks) {
        foreach (RtfTable table in blocks.OfType<RtfTable>()) {
            yield return table;
            foreach (RtfTable nested in RtfTables(table.Rows.SelectMany(row => row.Cells).SelectMany(cell => cell.Blocks)))
                yield return nested;
        }
    }
}
