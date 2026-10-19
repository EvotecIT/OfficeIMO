using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelRichTextRunService
{
    public static ExcelRichTextRun[] ToRuns(object[] runs)
    {
        return OfficeTextRunParser.ParseMany(runs).Select(ToRun).ToArray();
    }

    public static PSObject CreateRecord(
        ExcelRichTextRun run,
        int index,
        string address,
        int row,
        int column,
        string sheetName,
        string? path)
    {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Index", index));
        record.Properties.Add(new PSNoteProperty("Text", run.Text));
        record.Properties.Add(new PSNoteProperty("Bold", run.Bold));
        record.Properties.Add(new PSNoteProperty("Italic", run.Italic));
        record.Properties.Add(new PSNoteProperty("Underline", run.Underline));
        record.Properties.Add(new PSNoteProperty("FontColor", run.FontColor));
        record.Properties.Add(new PSNoteProperty("Color", run.FontColor));
        record.Properties.Add(new PSNoteProperty("FontName", run.FontName));
        record.Properties.Add(new PSNoteProperty("FontSize", run.FontSize));
        record.Properties.Add(new PSNoteProperty("Address", address));
        record.Properties.Add(new PSNoteProperty("Row", row));
        record.Properties.Add(new PSNoteProperty("Column", column));
        record.Properties.Add(new PSNoteProperty("SheetName", sheetName));
        record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
        if (!string.IsNullOrWhiteSpace(path))
        {
            record.Properties.Add(new PSNoteProperty("Path", path));
            record.Properties.Add(new PSNoteProperty("InputPath", path));
        }

        return record;
    }

    private static ExcelRichTextRun ToRun(OfficeTextRunSpec run)
    {
        return new ExcelRichTextRun(run.IsLineBreak ? "\n" : run.IsTab ? "\t" : run.Text)
        {
            Bold = run.Bold,
            Italic = run.Italic,
            Underline = run.Underline,
            FontColor = OfficeColorUtilities.ToExcelColorHex(run.Color),
            FontName = run.FontName,
            FontSize = run.FontSize
        };
    }
}
