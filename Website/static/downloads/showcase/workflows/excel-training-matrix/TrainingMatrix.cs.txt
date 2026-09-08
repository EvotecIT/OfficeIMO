using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Combines validated completion states with per-person and team training counts.</summary>
internal static class TrainingMatrix {
    internal static void Create(string folder) {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "example.xlsx"));
        var sheet = document.AddWorksheet("Training");
        sheet.SetColumnWidth(1, 26);
        for (int column = 2; column <= 5; column++) sheet.SetColumnWidth(column, 21);
        sheet.MergeRange("A1:E1"); sheet.Cell(1, 1, "Team training coverage");
        sheet.CellFontSize(1, 1, 20); sheet.CellBold(1, 1, true); sheet.CellFontColor(1, 1, "17365D");
        sheet.MergeRange("A2:E2"); sheet.Cell(2, 1, "Three required modules / sample completion record");
        string[,] values = {
            { "Colleague", "Service basics", "Data handling", "Recovery", "Completed" },
            { "Maya Ellis", "Complete", "Complete", "Complete", "" },
            { "Leon Brooks", "Complete", "Scheduled", "Not started", "" },
            { "Sofia Chen", "Complete", "Complete", "Scheduled", "" },
            { "Amir Patel", "Scheduled", "Not started", "Not started", "" },
            { "Eva Morgan", "Complete", "Complete", "Complete", "" }
        };
        for (int row = 0; row < 6; row++) for (int column = 0; column < 5; column++) sheet.Cell(row + 4, column + 1, values[row, column]);
        for (int row = 5; row <= 9; row++) sheet.CellFormula(row, 5, $"COUNTIF(B{row}:D{row},\"Complete\")");
        sheet.AddTable("A4:E9", hasHeader: true, name: "Training", style: ExcelTableStyle.TableStyleMedium2);
        sheet.ValidationList("B5:D9", new[] { "Not started", "Scheduled", "Complete" });
        sheet.Cell(11, 1, "Completed modules"); sheet.CellFormula(11, 5, "SUM(E5:E9)");
        sheet.Cell(12, 1, "Required modules"); sheet.Cell(12, 5, 15);
        sheet.Cell(13, 1, "Completion rate"); sheet.CellFormula(13, 5, "E11/E12"); sheet.FormatCell(13, 5, "0%");
        sheet.CellBackground(13, 5, "E7F6ED"); sheet.CellBold(13, 5, true);
        document.Calculate();
        sheet.Range("A1:E15").ExportImage(OfficeImageExportFormat.Png).Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
    }
}
