using System.Management.Automation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelPageBreakRecordService
{
    public static PSObject Create(string kind, int index, string sheetName, string? path)
    {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Type", kind));
        record.Properties.Add(new PSNoteProperty("Kind", kind));
        record.Properties.Add(new PSNoteProperty("Index", index));
        record.Properties.Add(new PSNoteProperty("Position", index));
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
