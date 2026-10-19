using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelCommentRecordService
{
    public static PSObject CreateRecord(ExcelCommentInfo comment, string sheetName, string? path)
    {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Address", comment.CellReference));
        record.Properties.Add(new PSNoteProperty("CellReference", comment.CellReference));
        record.Properties.Add(new PSNoteProperty("Row", comment.Row));
        record.Properties.Add(new PSNoteProperty("Column", comment.Column));
        record.Properties.Add(new PSNoteProperty("Author", comment.Author));
        record.Properties.Add(new PSNoteProperty("Text", comment.Text));
        record.Properties.Add(new PSNoteProperty("RunCount", comment.RichTextRuns.Count));
        record.Properties.Add(new PSNoteProperty("RichTextRuns", comment.RichTextRuns));
        record.Properties.Add(new PSNoteProperty("SheetName", sheetName));
        record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
        if (!string.IsNullOrWhiteSpace(path))
        {
            record.Properties.Add(new PSNoteProperty("Path", path));
            record.Properties.Add(new PSNoteProperty("InputPath", path));
        }

        return record;
    }
}
