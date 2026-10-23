using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    /// <summary>Reports Word's separated-cell spacing without substituting an RTF half-gap.</summary>
    private static void AddWordTableCellSpacingDiagnostics(IEnumerable<WordElement> elements, RtfConversionReport report) {
        int count = elements.OfType<WordTable>().Select(table => table._tableProperties)
            .Where(properties => properties != null).Distinct()
            .Count(properties => GetWordTableCellSpacingWidth(properties!.TableCellSpacing?.Width?.Value) != 0);
        if (count == 0) return;
        report.Add(RtfConversionSeverity.Warning, "WordRtfTableCellSpacingOmitted",
            "Word spacing between separated table cells is not represented by the RTF bridge and is omitted; it is not an RTF row half-gap.",
            RtfConversionAction.Omitted, feature: "table-cell-spacing", count: count);
    }

    private static int GetWordTableCellSpacingWidth(string? width) =>
        int.TryParse(width, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture,
            out int value) ? value : 0;

    /// <summary>Reports missing, mixed, negative or unrepresentable row gaps that cannot become one table margin.</summary>
    private static void AddRtfTableCellGapDiagnostics(RtfDocument document, RtfConversionReport report) {
        IEnumerable<IRtfBlock> blocks = document.Sections.Count > 0
            ? document.Sections.SelectMany(section => section.Blocks)
            : document.Blocks;
        int count = CountRtfTablesWithFlattenedCellGaps(blocks);
        if (count == 0) return;
        report.Add(RtfConversionSeverity.Warning, "RtfWordTableCellGapFlattened",
            "RTF tables with missing, mixed, negative or unrepresentable row half-gaps use zero table margins; row-specific gap geometry is not preserved.",
            RtfConversionAction.Flattened, feature: "table-cell-gap", count: count);
    }

    private static int CountRtfTablesWithFlattenedCellGaps(IEnumerable<IRtfBlock> blocks) {
        int count = 0;
        foreach (RtfTable table in blocks.OfType<RtfTable>()) {
            if (table.Rows.Count > 0 && !GetUniformTableCellGap(table).HasValue) count++;
            foreach (RtfTableRow row in table.Rows) {
                foreach (RtfTableCell cell in row.Cells) {
                    count += CountRtfTablesWithFlattenedCellGaps(cell.Blocks);
                }
            }
        }
        return count;
    }
}
