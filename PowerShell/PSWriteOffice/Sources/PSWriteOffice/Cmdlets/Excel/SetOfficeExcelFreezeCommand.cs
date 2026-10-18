using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Freezes panes on the current worksheet.</summary>
/// <example>
///   <summary>Freeze the top row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelFreeze -TopRows 1 }</code>
///   <para>Freezes the first row.</para>
/// </example>
/// <example>
///   <summary>Freeze the top row and first column.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelFreeze -TopRows 1 -LeftColumns 1 }</code>
///   <para>Freezes row 1 and column A.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelFreeze", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelFreeze")]
public sealed class SetOfficeExcelFreezeCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }
    /// <summary>Number of rows to freeze from the top.</summary>
    [Parameter]
    public int TopRows { get; set; }

    /// <summary>Number of columns to freeze from the left.</summary>
    [Parameter]
    public int LeftColumns { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (TopRows < 0 || LeftColumns < 0) {
            throw new PSArgumentException("TopRows and LeftColumns must be zero or greater.");
        }

        if (TopRows == 0 && LeftColumns == 0) {
            throw new PSArgumentException("Specify TopRows and/or LeftColumns to freeze.");
        }

        var sheet = ResolveSheet();
        sheet.Freeze(TopRows, LeftColumns);
        WritePassThru(sheet);
    }

    private ExcelSheet ResolveSheet() {
        if (ParameterSetName == ParameterSetDocument) {
            if (Document == null) {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }
}