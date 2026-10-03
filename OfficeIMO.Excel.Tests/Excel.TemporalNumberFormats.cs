using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelTemporalNumberFormatTests {
    [Theory]
    [InlineData("hh:mm:ss.0000", "13:04:09.9496")]
    [InlineData("hh:mm:ss.0000000", "13:04:09.9496000")]
    public void Pdf_retains_explicit_ISO_date_cell_ticks(string format, string expected) {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("ISO date");
        sheet.CellAt(1, 1).SetValue(0d).SetNumberFormat(format);
        sheet.SetColumnWidth(1, 35);
        using var saved = new MemoryStream();
        workbook.Save(saved);
        saved.Position = 0;
        using (var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(saved, true)) {
            var part = package.WorkbookPart!.WorksheetParts.First();
            var cell = part.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Date;
            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("2026-01-05T13:04:09.9496");
            part.Worksheet.Save();
        }
        saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        foreach (var layout in new[] { ExcelPdfWorksheetLayoutMode.WorksheetCanvas, ExcelPdfWorksheetLayoutMode.FlowTable }) {
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(reopened.ToPdfBytes(new ExcelToPdfOptions { IncludeSheetHeadings = false, WorksheetLayout = layout }));
            Assert.Contains(expected, string.Concat(pdf.GetPages().Select(page => page.Text)));
        }
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred, "hh:mm:ss.0", "13:04:09.9")]
    [InlineData(ExcelDateSystem.NineteenFour, "hh:mm:ss.0", "13:04:09.9")]
    [InlineData(ExcelDateSystem.NineteenHundred, "hh:mm:ss.0000", "13:04:09.9496")]
    [InlineData(ExcelDateSystem.NineteenFour, "hh:mm:ss.0000", "13:04:09.9496")]
    [InlineData(ExcelDateSystem.NineteenHundred, "hh:mm:ss.0000000", "13:04:09.9496002")]
    [InlineData(ExcelDateSystem.NineteenFour, "hh:mm:ss.0000000", "13:04:09.9496002")]
    public void Numeric_calendar_serial_preserves_available_submillisecond_precision(ExcelDateSystem system, string format, string expected) {
        double serial = 46027 + 47049.9496 / 86400 - (system == ExcelDateSystem.NineteenFour ? 1462 : 0);
        AssertSavedDisplay(serial, format, expected, system);
        using var workbook = ExcelDocument.Create();
        workbook.DateSystem = system;
        var sheet = workbook.AddWorksheet("Precision");
        sheet.CellAt(1, 1).SetValue(serial).SetNumberFormat(format);
        sheet.SetColumnWidth(1, 35);
        foreach (var layout in new[] { ExcelPdfWorksheetLayoutMode.WorksheetCanvas, ExcelPdfWorksheetLayoutMode.FlowTable }) {
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(workbook.ToPdfBytes(new ExcelToPdfOptions { IncludeSheetHeadings = false, WorksheetLayout = layout }));
            Assert.Contains(expected, string.Concat(pdf.GetPages().Select(page => page.Text)));
        }
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred, -0.25, "1899-12-30 18:00:00.0000")]
    [InlineData(ExcelDateSystem.NineteenFour, -1.25, "1903-12-30 18:00:00.0000")]
    [InlineData(ExcelDateSystem.NineteenFour, -1462.25, "1899-12-30 06:00:00.0000")]
    public void Fractional_display_preserves_the_existing_negative_calendar_epoch_contract(ExcelDateSystem system, double serial, string expected) {
        AssertSavedDisplay(serial, "yyyy-mm-dd hh:mm:ss.0000", expected, system);
    }

    [Theory]
    [InlineData("yyyy-mm-dd;yyyy-mm-dd;\"none\"", "yyyy-mm-dd")]
    [InlineData("[>=1]yyyy-mm-dd hh:mm;\"\"", "yyyy-mm-dd hh:mm")]
    public void Autofit_measures_the_section_selected_by_the_actual_date_serial(string sectionedFormat, string plainFormat) {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("Autofit");
        for (int column = 1; column <= 2; column++) {
            sheet.CellAt(1, column).SetValue(new DateTime(2026, 1, 5, 13, 4, 0))
                .SetNumberFormat(column == 1 ? plainFormat : sectionedFormat);
        }
        sheet.AutoFitColumns();
        var snapshot = sheet.Range("A1:B1").CreateVisualSnapshot();
        Assert.Equal(snapshot.Cells[0].Text, snapshot.Cells[1].Text);
        Assert.Equal(snapshot.Columns[0].Width, snapshot.Columns[1].Width);
    }

    [Theory]
    [InlineData("d.mm.yyyy hh:mm", "5.01.2026 13:04")]
    [InlineData("dd.mm.yyyy hh:mm:ss", "05.01.2026 13:04:09")]
    [InlineData("yyyy/mm/dd", "2026/01/05")]
    [InlineData("mm:dd", "01:05")]
    [InlineData("h\"h\" m\"m\"", "13h 4m")]
    [InlineData("mmmm d \"days\" yyyy", "January 5 days 2026")]
    [InlineData("mmmmm d, yyyy", "J 5, 2026")]
    [InlineData("yyyy-mm-dd \"[h]\" hh:mm", "2026-01-05 [h] 13:04")]
    [InlineData("yyyy-mm-dd\\h hh:mm", "2026-01-05h 13:04")]
    [InlineData("h:mm A/P", "1:04 P")]
    [InlineData("h:mm am/pm", "1:04 pm")]
    [InlineData("hh:mm:ss.00", "13:04:09.75")]
    [InlineData("hh:mm:ss.0", "13:04:09.8")]
    [InlineData("[>=1]yyyy-mm-dd;[h]:mm", "2026-01-05")]
    public void Saved_date_formats_preserve_components_and_literals(string format, string expected) {
        AssertSavedDisplay(new DateTime(2026, 1, 5, 13, 4, 9, 750).ToOADate(), format, expected);
    }

    [Theory]
    [InlineData(0.1, "[h]\"h\" m\"m\"", "2h 24m")]
    [InlineData(1.5, "[h]\"h\" m\"m\"", "36h 0m")]
    [InlineData(-0.1, "[h]\"h\" m\"m\"", "-2h 24m")]
    [InlineData(0.1, "[hh]\\h mm\\m", "02h 24m")]
    [InlineData(0.1, "[mm]\"m\" ss\"s\"", "144m 00s")]
    [InlineData(1.5, "[h] \"hours\"", "36 hours")]
    [InlineData(-1.5, "[h]:mm;[Red]([h]:mm)", "(36:00)")]
    [InlineData(-1.5, "[<0]\"negative \"[h]:mm;[h]:mm", "-negative 36:00")]
    [InlineData(0, "[h]:mm;[h]:mm;\"none\"", "none")]
    [InlineData(0, "[h]:mm;[h]:mm;", "")]
    [InlineData(1.5, "[>=1][h]\" hours\";[m]\" minutes\"", "36 hours")]
    [InlineData(0.1, "[>=1][h]\" hours\";[m]\" minutes\"", "144 minutes")]
    [InlineData(0.00078125, "[s].00", "67.50")]
    [InlineData(0.0006939814814814815, "[h]:mm:ss.0", "0:01:00.0")]
    [InlineData(1.5, "[h] mmm", "1.5")]
    public void Saved_elapsed_formats_preserve_units_padding_and_sections(double value, string format, string expected) {
        AssertSavedDisplay(value, format, expected);
    }

    [Fact]
    public void Native_Numbers_export_retains_cached_date_and_elapsed_display_after_save() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", "numbers-formulas-v14.5.xlsx");
        using var native = ExcelDocument.Load(path);
        AssertNativeDisplay(native);
        using var saved = new MemoryStream();
        native.Save(saved);
        saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        AssertNativeDisplay(reopened);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void Pdf_and_image_use_the_same_temporal_format_semantics(ExcelPdfWorksheetLayoutMode layout) {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("Formats");
        sheet.CellAt(1, 1).SetValue(new DateTime(2026, 1, 5, 13, 4, 0)).SetNumberFormat("d.mm.yyyy hh:mm");
        sheet.CellAt(2, 1).SetValue(0.1).SetNumberFormat("[h]\\h m\\m");
        sheet.CellAt(3, 1).SetValue(-1.5).SetNumberFormat("[h]:mm;[Red]([h]:mm)");
        sheet.CellAt(4, 1).SetValue(-0.1).SetNumberFormat("[h]\"h\" m\"m\"");
        sheet.SetColumnWidth(1, 35);
        byte[] pdf = workbook.ToPdfBytes(new ExcelToPdfOptions { IncludeSheetHeadings = false, WorksheetLayout = layout });
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf);
        string text = string.Concat(parsed.GetPages().Select(page => page.Text));
        Assert.Contains("5.01.2026 13:04", text);
        Assert.Contains("2h 24m", text);
        Assert.Contains("(36:00)", text);
        Assert.Contains("-2h 24m", text);
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred)]
    [InlineData(ExcelDateSystem.NineteenFour)]
    public void Temporal_display_obeys_calendar_epoch_without_shifting_elapsed_values(ExcelDateSystem system) {
        using var workbook = ExcelDocument.Create();
        workbook.DateSystem = system;
        var sheet = workbook.AddWorksheet("Formats");
        sheet.CellAt(1, 1).SetValue(new DateTime(2026, 1, 5, 13, 4, 59, 960)).SetNumberFormat("d.mm.yyyy hh:mm:ss.0");
        sheet.CellAt(2, 1).SetValue(0.1).SetNumberFormat("[h]\"h\" m\"m\"");
        sheet.SetColumnWidth(1, 40);
        using var saved = new MemoryStream();
        workbook.Save(saved);
        saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var cells = reopened.Sheets[0].Range("A1:A2").CreateVisualSnapshot().Cells;
        Assert.Equal("5.01.2026 13:05:00.0", cells.Single(cell => cell.Row == 1).Text);
        Assert.Equal("2h 24m", cells.Single(cell => cell.Row == 2).Text);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(reopened.ToPdfBytes(new ExcelToPdfOptions { IncludeSheetHeadings = false }));
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("5.01.2026 13:05:00.0", text);
        Assert.Contains("2h 24m", text);
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred, ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelDateSystem.NineteenHundred, ExcelPdfWorksheetLayoutMode.FlowTable)]
    [InlineData(ExcelDateSystem.NineteenFour, ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelDateSystem.NineteenFour, ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void Pdf_selects_conditional_sections_using_the_workbook_serial(ExcelDateSystem system, ExcelPdfWorksheetLayoutMode layout) {
        using var workbook = ExcelDocument.Create();
        workbook.DateSystem = system;
        var sheet = workbook.AddWorksheet("Epoch");
        // The same calendar date is serial 1463 in the 1900 system and 1 in 1904.
        sheet.CellAt(1, 1).SetValue(new DateTime(1904, 1, 2)).SetNumberFormat("[>=1000]yyyy-mm-dd\" large\";yyyy-mm-dd\" small\"");
        sheet.SetColumnWidth(1, 40);
        string expected = system == ExcelDateSystem.NineteenHundred ? "1904-01-02 large" : "1904-01-02 small";
        Assert.Equal(expected, Assert.Single(sheet.Range("A1:A1").CreateVisualSnapshot().Cells).Text);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(workbook.ToPdfBytes(new ExcelToPdfOptions { IncludeSheetHeadings = false, WorksheetLayout = layout }));
        Assert.Contains(expected, string.Concat(pdf.GetPages().Select(page => page.Text)));
    }

    private static void AssertSavedDisplay(double value, string format, string expected, ExcelDateSystem system = ExcelDateSystem.NineteenHundred) {
        using var workbook = ExcelDocument.Create();
        workbook.DateSystem = system;
        var sheet = workbook.AddWorksheet("Formats");
        sheet.CellAt(1, 1).SetValue(value).SetNumberFormat(format);
        using var saved = new MemoryStream();
        workbook.Save(saved);
        saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var range = reopened.Sheets[0].Range("A1:A1");
        Assert.Equal(expected, Assert.Single(range.CreateVisualSnapshot().Cells).Text);
        Assert.Equal(value, reopened.Sheets[0].CellAt(1, 1).GetValue().Value);
    }

    private static void AssertNativeDisplay(ExcelDocument workbook) {
        var cells = workbook.Sheets[0].Range("C4:C6").CreateVisualSnapshot().Cells;
        Assert.Equal("30.09.2026 12:42", cells.Single(cell => cell.Row == 4).Text);
        Assert.Equal("30.09.2026 15:06", cells.Single(cell => cell.Row == 5).Text);
        Assert.Equal("2h 24m", cells.Single(cell => cell.Row == 6).Text);
        Assert.Equal("NOW()", workbook.Sheets[0].CellAt(4, 3).GetValue().Formula);
        Assert.Equal(46295.529641203706, workbook.Sheets[0].CellAt(4, 3).GetValue().Value);
    }
}
