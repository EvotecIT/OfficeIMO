using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
namespace PSWriteOffice.Services.Excel;
/// <summary>Collects native ExcelText children without introducing a parallel run schema.</summary>
internal static class ExcelRichTextRunService {
    public static ExcelRichTextRun[] ReadContent(ScriptBlock content) => content.Invoke().Select(item => OfficeArgumentValue.Unwrap(item) is ExcelRichTextRun run ? run : throw new PSArgumentException("Excel rich-text content must emit ExcelText or native ExcelRichTextRun objects.")).ToArray();
    public static PSObject CreateRecord(ExcelRichTextRun run, int index, string address, int row, int column, string sheetName, string? path) {
        var record = PSObject.AsPSObject(run);
        record.Properties.Add(new PSNoteProperty("Index", index + 1));
        record.Properties.Add(new PSNoteProperty("Address", address));
        record.Properties.Add(new PSNoteProperty("Row", row));
        record.Properties.Add(new PSNoteProperty("Column", column));
        record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
        if (path != null) record.Properties.Add(new PSNoteProperty("Path", path));
        return record;
    }
}
