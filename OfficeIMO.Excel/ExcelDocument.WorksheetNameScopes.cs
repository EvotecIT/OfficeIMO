using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private void ReindexLocalNamesAfterSheetReorder(IReadOnlyList<Sheet> previous, IReadOnlyList<Sheet> current) {
        var positions = current.Select((sheet, index) => (sheet, index)).ToDictionary(item => item.sheet, item => item.index);
        foreach (var name in WorkbookRoot.DefinedNames?.Elements<DefinedName>() ?? Enumerable.Empty<DefinedName>()) {
            if (name.LocalSheetId?.Value is uint scope && scope < previous.Count)
                name.LocalSheetId = checked((uint)positions[previous[(int)scope]]);
        }
    }
}
