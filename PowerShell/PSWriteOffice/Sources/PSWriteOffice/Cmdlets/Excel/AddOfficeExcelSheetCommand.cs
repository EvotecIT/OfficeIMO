using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds or reuses a worksheet within the current Excel DSL scope.</summary>
/// <para>Creates the sheet when missing, pushes it onto the DSL stack, and executes the nested script block.</para>
/// <example>
/// <summary>Add a sheet with content</summary>
/// <code>New-OfficeExcel -Path .\Inventory.xlsx {
///     ExcelSheet 'Servers' {
///         ExcelTable (Import-OfficeCsv -Path .\Data\Servers.csv)
///     }
/// }</code>
/// <para>Adds an Inventory sheet and writes a table of servers into it.</para>
/// </example>
/// <example>
/// <summary>Create one sheet per department</summary>
/// <code>$staff = Import-OfficeCsv -Path .\HR\Headcount-2026-10.csv
/// New-OfficeExcel -Path .\Reports\Headcount.xlsx {
///     foreach ($group in $staff | Group-Object Department) {
///         ExcelSheet $group.Name -ValidationMode Sanitize {
///             ExcelTable ($group.Group | Select-Object Name, Title, Location, StartDate)
///         }
///     }
/// }</code>
/// <para>Groups the headcount export by department and writes each group to its own sheet, with a cleaned-up name when a department name has characters Excel does not allow.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelSheet", DefaultParameterSetName = "Default")]
[Alias("ExcelSheet")]
[OutputType(typeof(ExcelSheet))]
public sealed class AddOfficeExcelSheetCommand : PSCmdlet
{

    /// <summary>Workbook that will receive the worksheet.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Name of the worksheet to create or reuse. When omitted the last sheet is reused or a default sheet is created.</summary>
    [Parameter(ParameterSetName = "Default", Position = 0)]
    public string? Name { get; set; }

    /// <summary>Controls how invalid sheet names are handled.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelSheetNameValidationMode ValidationMode { get; set; } = ExcelSheetNameValidationMode.Sanitize;

    /// <summary>Code to execute inside the worksheet context.</summary>
    [Parameter(ParameterSetName = "Default", Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit the <see cref="ExcelSheet"/> object after execution.</summary>
    [Parameter(ParameterSetName = "Default")]
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
