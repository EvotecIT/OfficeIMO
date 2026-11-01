using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Lists manual print page breaks without changing a workbook.</summary>
/// <example><summary>List page breaks in literal workbook files.</summary>
/// <code>Get-ChildItem ./Reports -Filter *.xlsx | Get-OfficeExcelPageBreak -Sheet 'Report*'</code></example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelPageBreak", DefaultParameterSetName = "Document")]
[OutputType(typeof(ExcelPageBreakInfo))]
public sealed class GetOfficeExcelPageBreakCommand : ExcelQueryCmdlet {
    /// <summary>Include row breaks; neither switch or both switches includes both kinds.</summary>
    [Parameter] public SwitchParameter Row { get; set; }
    /// <summary>Include column breaks; neither switch or both switches includes both kinds.</summary>
    [Parameter] public SwitchParameter Column { get; set; }

    /// <inheritdoc />
    protected override IEnumerable<PSObject> Query(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets, string? path) {
        foreach (ExcelSheet sheet in sheets) {
            if (Row || !Column)
                foreach (int row in sheet.GetManualRowPageBreaks())
                    yield return PSObject.AsPSObject(ExcelPageBreakRecordService.Create("Row", row, sheet.Name, path));
            if (Column || !Row)
                foreach (int column in sheet.GetManualColumnPageBreaks())
                    yield return PSObject.AsPSObject(ExcelPageBreakRecordService.Create("Column", column, sheet.Name, path));
        }
    }
}
