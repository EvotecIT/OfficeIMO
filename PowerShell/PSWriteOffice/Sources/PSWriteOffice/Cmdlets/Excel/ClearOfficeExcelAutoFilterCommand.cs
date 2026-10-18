using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Clears any AutoFilter on the current worksheet.</summary>
/// <example>
///   <summary>Remove AutoFilter from the active sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Clear-OfficeExcelAutoFilter }</code>
///   <para>Removes filter dropdowns and criteria.</para>
/// </example>
[Cmdlet(VerbsCommon.Clear, "OfficeExcelAutoFilter", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelAutoFilterClear")]
public sealed class ClearOfficeExcelAutoFilterCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
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

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ResolveSheet();
        sheet.AutoFilterClear();
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