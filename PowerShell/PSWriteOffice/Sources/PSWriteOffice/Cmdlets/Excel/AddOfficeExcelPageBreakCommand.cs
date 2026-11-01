using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds manual print page breaks to an open worksheet.</summary>
/// <example><summary>Separate report sections when printing.</summary>
/// <code>ExcelSheet 'Data' { ExcelPageBreak -Row 25,50 -Column 8 }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelPageBreak", DefaultParameterSetName = "Default")]
[Alias("ExcelPageBreak")]
[OutputType(typeof(ExcelPageBreakInfo))]
public sealed class AddOfficeExcelPageBreakCommand : OfficeDocumentCmdlet<ExcelDocument> {
    /// <summary>Worksheet name, one-based position or native worksheet; omission uses the current sheet.</summary>
    [Parameter(ParameterSetName = "Default"), OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }
    /// <summary>One-based rows after which printing starts a new page.</summary>
    [Parameter, ValidateRange(1, 1048576)]
    public int[] Row { get; set; } = Array.Empty<int>();
    /// <summary>One-based columns after which printing starts a new page.</summary>
    [Parameter, ValidateRange(1, 16384)]
    public int[] Column { get; set; } = Array.Empty<int>();
    /// <summary>Emit the worksheet's page-break snapshots after adding the selected breaks.</summary>
    [Parameter] public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Row.Length == 0 && Column.Length == 0) {
            InvalidArgument(new PSArgumentException("Provide at least one Row or Column."), null);
            return;
        }
        ExcelSheet sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        foreach (int row in Row) sheet.AddManualRowPageBreak(row, save: false);
        foreach (int column in Column) sheet.AddManualColumnPageBreak(column, save: false);
        if (!PassThru) return;
        foreach (int row in sheet.GetManualRowPageBreaks())
            WriteObject(ExcelPageBreakRecordService.Create("Row", row, sheet.Name, null));
        foreach (int column in sheet.GetManualColumnPageBreaks())
            WriteObject(ExcelPageBreakRecordService.Create("Column", column, sheet.Name, null));
    }
}
