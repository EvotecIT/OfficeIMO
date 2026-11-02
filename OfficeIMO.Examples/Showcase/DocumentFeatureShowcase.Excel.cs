using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Showcase;

internal static partial class DocumentFeatureShowcase {
    private static void CreateExcelExamples(string output) {
        CreateExcelForecast(output);
        CreateExcelInventory(output);
        CreateExcelChart(output);
        CreateExcelHeatmap(output);
        CreateExcelInputForm(output);
    }

    private static ExcelSheet PrepareSheet(ExcelDocument document, string title, string subtitle) {
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Range("A1:F1").Merge();
        sheet.Cell(1, 1).SetValue(title);
        sheet.Cell(1, 1).SetBold(true);
        sheet.Cell(1, 1).SetFontSize(global::OfficeIMO.OfficeLength.Points(20));
        sheet.Cell(1, 1).SetFontColor("17365D");
        sheet.Range("A2:F2").Merge();
        sheet.Cell(2, 1).SetValue(subtitle);
        sheet.Cell(2, 1).SetFontColor("526179");
        for (int column = 1; column <= 6; column++) {
            sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(column == 1 ? 24 : 17));
        }
        return sheet;
    }

    private static void SetHeaders(ExcelSheet sheet, params string[] headers) {
        for (int column = 0; column < headers.Length; column++) {
            sheet.Cell(4, column + 1).SetValue(headers[column]);
            sheet.Cell(4, column + 1).SetBold(true);
            sheet.Cell(4, column + 1).SetFillColor("17365D");
            sheet.Cell(4, column + 1).SetFontColor("FFFFFF");
        }
    }

    private static void SaveExcelExample(ExcelDocument document, ExcelSheet sheet, string folder, string range) {
        sheet.Range(range).ExportImage(OfficeImageExportFormat.Png)
            .Save(Path.Combine(folder, "preview.png"), OfficeImageExportFileConflictPolicy.Replace);
        document.Save();
        if (!document.DocumentIsValid) {
            throw new InvalidOperationException("Showcase workbook failed Open XML validation: " + folder);
        }
    }

    private static void CreateExcelForecast(string output) {
        string folder = CreateExampleFolder(output, "excel-formulas");
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "forecast.xlsx"));
        ExcelSheet sheet = PrepareSheet(document, "Six-month operating forecast", "Formula cells calculate profit and margin from revenue and cost.");
        SetHeaders(sheet, "Month", "Revenue", "Cost", "Profit", "Margin");
        string[] months = { "January", "February", "March", "April", "May", "June" };
        for (int index = 0; index < months.Length; index++) {
            int row = index + 5;
            sheet.Cell(row, 1).SetValue(months[index]);
            sheet.Cell(row, 2).SetValue(24000 + index * 3500).SetNumberFormat("#,##0");
            sheet.Cell(row, 3).SetValue(17500 + index * 1400).SetNumberFormat("#,##0");
            sheet.Cell(row, 4).SetFormula($"B{row}-C{row}");
            sheet.Cell(row, 5).SetFormula($"D{row}/B{row}");
            sheet.Cell(row, 4).SetFillColor("E8F4EC");
            sheet.Cell(row, 5).SetFillColor("E8F4EC");
            sheet.Cell(row, 4).SetNumberFormat("#,##0");
            sheet.Cell(row, 5).SetNumberFormat("0.0%");
        }
        document.Calculate();
        SaveExcelExample(document, sheet, folder, "A1:F13");
    }

    private static void CreateExcelInventory(string output) {
        string folder = CreateExampleFolder(output, "excel-tables");
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "inventory.xlsx"));
        ExcelSheet sheet = PrepareSheet(document, "Inventory you can filter", "A styled Excel table with filter controls and frozen headings.");
        SetHeaders(sheet, "Item", "Category", "In stock", "Reorder at", "Location");
        string[] items = { "Docking station", "USB-C cable", "Monitor arm", "Keyboard", "Webcam", "Headset" };
        for (int index = 0; index < items.Length; index++) {
            int row = index + 5;
            sheet.Cell(row, 1).SetValue(items[index]);
            sheet.Cell(row, 2).SetValue(index < 3 ? "Workspace" : "Accessories");
            sheet.Cell(row, 3).SetValue(new[] { 32, 140, 18, 65, 23, 42 }[index]);
            sheet.Cell(row, 4).SetValue(20);
            sheet.Cell(row, 5).SetValue(index % 2 == 0 ? "Main store" : "Annex");
        }
        sheet.AddTable("A4:E10", hasHeader: true, name: "Inventory", style: ExcelTableStyle.TableStyleMedium2);
        sheet.Freeze(topRows: 4, leftCols: 1);
        SaveExcelExample(document, sheet, folder, "A1:F13");
    }

    private static void CreateExcelChart(string output) {
        string folder = CreateExampleFolder(output, "excel-charts");
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "quarterly-sales.xlsx"));
        ExcelSheet sheet = PrepareSheet(document, "Quarterly sales against target", "Editable chart series stay inside the workbook.");
        SetHeaders(sheet, "Quarter", "Actual", "Target");
        for (int index = 0; index < 4; index++) {
            sheet.Cell(index + 5, 1).SetValue("Q" + (index + 1));
            sheet.Cell(index + 5, 2).SetValue(new[] { 42, 58, 67, 81 }[index]);
            sheet.Cell(index + 5, 3).SetValue(new[] { 45, 55, 65, 75 }[index]);
        }
        sheet.AddChart(OfficeChartType.ColumnClustered, "A4:C8", new OfficeIMO.Excel.ExcelChartOptions { Row = 10, Column = 1, Width = OfficeLength.Pixels(620), Height = OfficeLength.Pixels(300), Title = "Actual vs target" });
        SaveExcelExample(document, sheet, folder, "A1:F27");
    }

    private static void CreateExcelHeatmap(string output) {
        string folder = CreateExampleFolder(output, "excel-conditional-formatting");
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "capacity.xlsx"));
        ExcelSheet sheet = PrepareSheet(document, "Team capacity at a glance", "Conditional formatting applies a color scale to the numeric cells.");
        SetHeaders(sheet, "Team", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday");
        string[] teams = { "Engineering", "Operations", "Support", "Design", "Delivery" };
        for (int team = 0; team < teams.Length; team++) {
            sheet.Cell(team + 5, 1).SetValue(teams[team]);
            for (int day = 0; day < 5; day++) {
                sheet.Cell(team + 5, day + 2).SetValue(25 + ((team * 17 + day * 23) % 76));
            }
        }
        sheet.AddConditionalColorScale("B5:F9", OfficeColor.Parse("#FDE8E7"), OfficeColor.Parse("#16A34A"));
        SaveExcelExample(document, sheet, folder, "A1:F12");
    }

    private static void CreateExcelInputForm(string output) {
        string folder = CreateExampleFolder(output, "excel-validation");
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(folder, "request-register.xlsx"));
        ExcelSheet sheet = PrepareSheet(document, "A request register with guardrails", "The workbook validates priority, status, and whole-number estimates.");
        SetHeaders(sheet, "Request", "Owner", "Priority", "Status", "Estimate");
        string[] requests = { "Refresh runbook", "Review access", "Update dashboard", "Test recovery", "Triage backlog", "Plan workshop" };
        for (int index = 0; index < requests.Length; index++) {
            int row = index + 5;
            sheet.Cell(row, 1).SetValue(requests[index]);
            sheet.Cell(row, 2).SetValue(index % 2 == 0 ? "Operations" : "Engineering");
            sheet.Cell(row, 3).SetValue(index % 3 == 0 ? "High" : "Normal");
            sheet.Cell(row, 4).SetValue(index < 2 ? "Complete" : "Planned");
            sheet.Cell(row, 5).SetValue(index + 1);
        }
        sheet.ValidationList("C5:C10", new[] { "Low", "Normal", "High" });
        sheet.ValidationList("D5:D10", new[] { "Planned", "In progress", "Complete" });
        sheet.ValidationWholeNumber("E5:E10", ExcelDataValidationOperator.Between, 1, 20);
        sheet.AddTable("A4:E10", hasHeader: true, name: "Requests", style: ExcelTableStyle.TableStyleMedium2);
        SaveExcelExample(document, sheet, folder, "A1:F13");
    }
}
