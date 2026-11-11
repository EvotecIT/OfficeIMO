using OfficeIMO;
using OfficeIMO.Excel;

internal static class HiddenDataSmoke {
    internal static void Run() {
        using var stream = new MemoryStream();
        using ExcelDocument document = ExcelDocument.Create(stream);
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Value");
        sheet.SetComment(1, 1, "Private note", "NativeAOT Author");
        OfficeHiddenDataReport report = document.RemoveHiddenData();
        if (!report.Changed || sheet.GetComments().Count != 0) {
            throw new InvalidOperationException("Excel hidden-data removal did not remove the note.");
        }
        Console.WriteLine("PASS | Excel hidden-data removal");
    }
}
