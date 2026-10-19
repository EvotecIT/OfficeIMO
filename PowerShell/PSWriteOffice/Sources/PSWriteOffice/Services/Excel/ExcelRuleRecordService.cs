using System;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelRuleRecordService
{
    public static PSObject CreateConditionalFormattingRecord(
        ExcelConditionalFormattingInfo rule,
        string sheetName,
        string? path)
    {
        var record = CreateBaseRecord(sheetName, rule.Range, path);
        record.Properties.Add(new PSNoteProperty("Type", rule.Type));
        record.Properties.Add(new PSNoteProperty("Operator", rule.Operator));
        record.Properties.Add(new PSNoteProperty("Priority", rule.Priority));
        record.Properties.Add(new PSNoteProperty("StopIfTrue", rule.StopIfTrue));
        record.Properties.Add(new PSNoteProperty("Formulas", rule.Formulas));
        record.Properties.Add(new PSNoteProperty("ColorScaleColors", rule.ColorScaleColors));
        record.Properties.Add(new PSNoteProperty("DataBarColor", rule.DataBarColor));
        record.Properties.Add(new PSNoteProperty("IconSet", rule.IconSet));
        record.Properties.Add(new PSNoteProperty("IconSetShowValue", rule.IconSetShowValue));
        record.Properties.Add(new PSNoteProperty("IconSetReverse", rule.IconSetReverse));
        return record;
    }

    public static PSObject CreateDataValidationRecord(
        ExcelDataValidationInfo validation,
        string sheetName,
        string? path)
    {
        var record = CreateBaseRecord(sheetName, validation.Range, path);
        record.Properties.Add(new PSNoteProperty("Type", validation.Type));
        record.Properties.Add(new PSNoteProperty("Operator", validation.Operator));
        record.Properties.Add(new PSNoteProperty("AllowBlank", validation.AllowBlank));
        record.Properties.Add(new PSNoteProperty("Formula1", validation.Formula1));
        record.Properties.Add(new PSNoteProperty("Formula2", validation.Formula2));
        record.Properties.Add(new PSNoteProperty("PromptTitle", validation.PromptTitle));
        record.Properties.Add(new PSNoteProperty("Prompt", validation.Prompt));
        record.Properties.Add(new PSNoteProperty("ErrorTitle", validation.ErrorTitle));
        record.Properties.Add(new PSNoteProperty("Error", validation.Error));
        return record;
    }

    private static PSObject CreateBaseRecord(string sheetName, string range, string? path)
    {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("SheetName", sheetName));
        record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
        record.Properties.Add(new PSNoteProperty("Range", range));
        if (!string.IsNullOrWhiteSpace(path))
        {
            record.Properties.Add(new PSNoteProperty("Path", path));
            record.Properties.Add(new PSNoteProperty("InputPath", path));
        }

        return record;
    }
}
