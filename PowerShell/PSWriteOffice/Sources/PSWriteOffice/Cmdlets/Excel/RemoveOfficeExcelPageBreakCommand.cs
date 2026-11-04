using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Removes manual print page breaks from selected open worksheets.</summary>
/// <example><summary>Clear manual breaks from report sheets before saving.</summary>
/// <code>Open-OfficeExcel -Path ./Report.xlsx { Remove-OfficeExcelPageBreak -Sheet 'Report*' -All }</code></example>
[Cmdlet(VerbsCommon.Remove, "OfficeExcelPageBreak", DefaultParameterSetName = "Default")]
[OutputType(typeof(ExcelPageBreakInfo))]
public sealed class RemoveOfficeExcelPageBreakCommand : OfficeDocumentCmdlet<ExcelDocument> {
    /// <summary>Worksheet names, wildcard patterns, one-based positions or native worksheets.</summary>
    [Parameter(ParameterSetName = "Default"), OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }
    /// <summary>One-based rows whose manual breaks are removed.</summary>
    [Parameter, ValidateRange(1, 1048576)]
    public int[] Row { get; set; } = Array.Empty<int>();
    /// <summary>One-based columns whose manual breaks are removed.</summary>
    [Parameter, ValidateRange(1, 16384)]
    public int[] Column { get; set; } = Array.Empty<int>();
    /// <summary>Remove every manual row and column break on the selected worksheets.</summary>
    [Parameter] public SwitchParameter All { get; set; }
    /// <summary>Emit a snapshot for each break actually removed.</summary>
    [Parameter] public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!All && Row.Length == 0 && Column.Length == 0) {
            InvalidArgument(new PSArgumentException("Provide Row, Column, or All."), null);
            return;
        }
        ExcelDocument document = GetDocument();
        var sheets = ExcelSheetResolver.ResolveMany(document, Sheet, Document == null ? ExcelDslContext.Current?.CurrentSheet : null);
        foreach (ExcelSheet sheet in sheets) {
            var rows = All ? sheet.GetManualRowPageBreaks() : Row;
            var columns = All ? sheet.GetManualColumnPageBreaks() : Column;
            foreach (int row in rows)
                if (sheet.RemoveManualRowPageBreak(row, save: false) && PassThru)
                    WriteObject(ExcelPageBreakRecordService.Create("Row", row, sheet.Name, null));
            foreach (int column in columns)
                if (sheet.RemoveManualColumnPageBreak(column, save: false) && PassThru)
                    WriteObject(ExcelPageBreakRecordService.Create("Column", column, sheet.Name, null));
        }
    }
}
