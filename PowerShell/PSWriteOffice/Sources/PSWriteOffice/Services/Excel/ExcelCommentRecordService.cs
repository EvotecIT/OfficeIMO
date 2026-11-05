using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelCommentRecordService
{
    public static PSObject CreateRecord(ExcelCommentInfo comment, string sheetName, string? path)
    {
        var record = PSObject.AsPSObject(comment);
        record.Properties.Add(new PSNoteProperty("Address", comment.CellReference));
        record.Properties.Add(new PSNoteProperty("RunCount", comment.RichTextRuns.Count));
        record.Properties.Add(new PSNoteProperty("SheetName", sheetName));
        record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
        if (!string.IsNullOrWhiteSpace(path))
        {
            record.Properties.Add(new PSNoteProperty("Path", path));
        }

        return record;
    }
}
