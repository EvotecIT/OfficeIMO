using System.Linq;
using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Clears any AutoFilter on the current worksheet.</summary>
/// <example>
/// <summary>Remove the AutoFilter from a sheet</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
///     ExcelSheet 'Tickets' {
///         Clear-OfficeExcelAutoFilter
///     }
/// }</code>
/// <para>Removes the filter buttons and any filter conditions from the Tickets sheet.</para>
/// </example>
/// <example>
/// <summary>Clear filters before sharing a workbook</summary>
/// <code>$book = Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx
/// foreach ($sheet in 'Orders', 'Returns', 'Customers') {
///     Clear-OfficeExcelAutoFilter -Document $book -Sheet $sheet
/// }
/// Save-OfficeExcel -Document $book -Path .\Outbox\Q3-Sales.xlsx -PassThru | Close-OfficeExcel</code>
/// <para>Removes the filters left on the working sheets so recipients see every row.</para>
/// </example>
[Cmdlet(VerbsCommon.Clear, "OfficeExcelAutoFilter", DefaultParameterSetName = "Default")]
public sealed class ClearOfficeExcelAutoFilterCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ResolveSheet();
        ExcelAutoFilterInfo? worksheetFilter = sheet.GetAutoFilters().FirstOrDefault(filter => !filter.IsTableFilter);
        if (worksheetFilter != null) {
            sheet.Range(worksheetFilter.Range).AutoFilter().Clear();
        }
        WritePassThru(sheet);
    }

    private ExcelSheet ResolveSheet() => ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
}
