using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds or reuses a worksheet within the current Excel DSL scope.</summary>
/// <para>Creates the sheet when missing, pushes it onto the DSL stack, and executes the nested script block.</para>
/// <example>
///   <summary>Create a sheet named Data.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\report.xlsx { Add-OfficeExcelSheet -Name 'Data' { ExcelCell -Address 'A1' -Value 'Region' } }</code>
///   <para>Creates a workbook with a worksheet named Data and writes the header “Region”.</para>
/// </example>
/// <example>
///   <summary>Add a worksheet through the workbook pipeline.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$sheet = $workbook | Add-OfficeExcelSheet -Name 'Data' -PassThru
/// $sheet | Set-OfficeExcelCell -Address A1 -Value 'Region'</code>
///   <para>Returns the worksheet so subsequent commands can target it directly.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelSheet", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelSheet")]
[OutputType(typeof(ExcelSheet))]
public sealed class AddOfficeExcelSheetCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook that will receive the worksheet.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Name of the worksheet to create or reuse. When omitted the last sheet is reused or a default sheet is created.</summary>
    [Parameter(Position = 0)]
    public string? Name { get; set; }

    /// <summary>Controls how invalid sheet names are handled.</summary>
    [Parameter]
    public ExcelSheetNameValidationMode ValidationMode { get; set; } = ExcelSheetNameValidationMode.Sanitize;

    /// <summary>Code to execute inside the worksheet context.</summary>
    [Parameter(Position = 1, ParameterSetName = ParameterSetContext)]
    [Parameter(Position = 1, ParameterSetName = ParameterSetDocument)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit the <see cref="ExcelSheet"/> object after execution.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            var context = ExcelDslContext.Require(this);
            var sheet = context.Document.GetOrCreateSheet(Name, ValidationMode);

            using (context.Push(sheet))
            {
                Content?.InvokeReturnAsIs();
            }

            if (PassThru.IsPresent)
            {
                WriteObject(sheet);
            }
            return;
        }

        var currentContext = Content != null ? ExcelDslContext.Current : null;
        if (currentContext != null && !ReferenceEquals(currentContext.Document, Document))
        {
            throw new InvalidOperationException(
                "The explicit workbook target does not match the active Excel composition scope.");
        }

        var createdSheet = Document.GetOrCreateSheet(Name, ValidationMode);
        if (Content != null)
        {
            if (currentContext != null)
            {
                using (currentContext.Push(createdSheet))
                {
                    Content.InvokeReturnAsIs();
                }
            }
            else
            {
                using var context = ExcelDslContext.Enter(Document);
                using (context.Push(createdSheet))
                {
                    Content.InvokeReturnAsIs();
                }
            }
        }

        if (PassThru.IsPresent)
        {
            WriteObject(createdSheet);
        }
    }
}
