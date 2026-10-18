using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates or updates a named range.</summary>
/// <para>Defaults to the current sheet scope when used inside the Excel DSL.</para>
/// <example>
///   <summary>Define a named range inside a sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelNamedRange -Name 'Totals' -Range 'B2:B50' }</code>
///   <para>Creates a sheet-scoped name for the range.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelNamedRange", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelNamedRange")]
public sealed class SetOfficeExcelNamedRangeCommand : PSCmdlet
{
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

    /// <summary>Name of the defined range.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Range in A1 notation.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Mark the defined name as hidden.</summary>
    [Parameter]
    public SwitchParameter Hidden { get; set; }

    /// <summary>Validate or sanitize the defined name.</summary>
    [Parameter]
    public ExcelDefinedNameValidationMode ValidationMode { get; set; } = ExcelDefinedNameValidationMode.Sanitize;

    /// <summary>Force a workbook-global name even inside a sheet block.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    public SwitchParameter Global { get; set; }

    /// <summary>Save the workbook immediately after setting the name.</summary>
    [Parameter]
    public SwitchParameter Save { get; set; }

    /// <summary>Emit the name after creation.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDocument document;
        ExcelSheet? scope = null;

        if (ParameterSetName == ParameterSetDocument)
        {
            document = Document ?? throw new PSArgumentException("Provide an Excel document.");

            if (!string.IsNullOrWhiteSpace(Sheet))
            {
                scope = document[Sheet!];
            }
            else if (SheetIndex.HasValue)
            {
                if (SheetIndex.Value < 0 || SheetIndex.Value >= document.Sheets.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(SheetIndex), "SheetIndex is out of range.");
                }
                scope = document.Sheets[SheetIndex.Value];
            }
        }
        else
        {
            var context = ExcelDslContext.Require(this);
            document = context.Document;
            if (!Global.IsPresent)
            {
                scope = context.CurrentSheet;
            }
        }

        document.SetNamedRange(Name, Range, scope, save: false, hidden: Hidden.IsPresent, validationMode: ValidationMode);
        if (Save.IsPresent)
        {
            document.Save();
        }

        if (PassThru.IsPresent)
        {
            WriteObject(Name);
        }
    }
}
