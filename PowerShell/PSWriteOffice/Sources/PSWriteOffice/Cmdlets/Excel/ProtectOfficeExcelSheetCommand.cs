using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Protects the current worksheet.</summary>
/// <example>
///   <summary>Protect the active sheet with default options.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Protect-OfficeExcelSheet }</code>
///   <para>Enables worksheet protection.</para>
/// </example>
[Cmdlet(VerbsSecurity.Protect, "OfficeExcelSheet", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelProtect")]
public sealed class ProtectOfficeExcelSheetCommand : PSCmdlet
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

    /// <summary>Allow common protected-table workflows: selecting cells, inserting rows, sorting, and filtering.</summary>
    [Parameter]
    public SwitchParameter AllowTableEditing { get; set; }

    /// <summary>Allow selecting locked cells.</summary>
    [Parameter]
    public bool AllowSelectLockedCells { get; set; } = true;

    /// <summary>Allow selecting unlocked cells.</summary>
    [Parameter]
    public bool AllowSelectUnlockedCells { get; set; } = true;

    /// <summary>Allow formatting cells.</summary>
    [Parameter]
    public bool AllowFormatCells { get; set; }

    /// <summary>Allow formatting columns.</summary>
    [Parameter]
    public bool AllowFormatColumns { get; set; }

    /// <summary>Allow formatting rows.</summary>
    [Parameter]
    public bool AllowFormatRows { get; set; }

    /// <summary>Allow inserting columns.</summary>
    [Parameter]
    public bool AllowInsertColumns { get; set; }

    /// <summary>Allow inserting rows.</summary>
    [Parameter]
    public bool AllowInsertRows { get; set; }

    /// <summary>Allow inserting hyperlinks.</summary>
    [Parameter]
    public bool AllowInsertHyperlinks { get; set; }

    /// <summary>Allow deleting columns.</summary>
    [Parameter]
    public bool AllowDeleteColumns { get; set; }

    /// <summary>Allow deleting rows.</summary>
    [Parameter]
    public bool AllowDeleteRows { get; set; }

    /// <summary>Allow sorting.</summary>
    [Parameter]
    public bool AllowSort { get; set; }

    /// <summary>Allow AutoFilter.</summary>
    [Parameter]
    public bool AllowAutoFilter { get; set; }

    /// <summary>Allow PivotTables.</summary>
    [Parameter]
    public bool AllowPivotTables { get; set; }

    /// <summary>Emit the worksheet after protection.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        var options = AllowTableEditing.IsPresent
            ? ExcelSheetProtectionOptions.TableEditing()
            : new ExcelSheetProtectionOptions
            {
                AllowSelectLockedCells = AllowSelectLockedCells,
                AllowSelectUnlockedCells = AllowSelectUnlockedCells
            };

        ApplyBoundOption(options, nameof(AllowSelectLockedCells), value => options.AllowSelectLockedCells = value);
        ApplyBoundOption(options, nameof(AllowSelectUnlockedCells), value => options.AllowSelectUnlockedCells = value);
        ApplyBoundOption(options, nameof(AllowFormatCells), value => options.AllowFormatCells = value);
        ApplyBoundOption(options, nameof(AllowFormatColumns), value => options.AllowFormatColumns = value);
        ApplyBoundOption(options, nameof(AllowFormatRows), value => options.AllowFormatRows = value);
        ApplyBoundOption(options, nameof(AllowInsertColumns), value => options.AllowInsertColumns = value);
        ApplyBoundOption(options, nameof(AllowInsertRows), value => options.AllowInsertRows = value);
        ApplyBoundOption(options, nameof(AllowInsertHyperlinks), value => options.AllowInsertHyperlinks = value);
        ApplyBoundOption(options, nameof(AllowDeleteColumns), value => options.AllowDeleteColumns = value);
        ApplyBoundOption(options, nameof(AllowDeleteRows), value => options.AllowDeleteRows = value);
        ApplyBoundOption(options, nameof(AllowSort), value => options.AllowSort = value);
        ApplyBoundOption(options, nameof(AllowAutoFilter), value => options.AllowAutoFilter = value);
        ApplyBoundOption(options, nameof(AllowPivotTables), value => options.AllowPivotTables = value);

        sheet.Protect(options);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }

    private void ApplyBoundOption(ExcelSheetProtectionOptions options, string parameterName, Action<bool> apply)
    {
        if (MyInvocation.BoundParameters.TryGetValue(parameterName, out var value) && value is bool enabled)
        {
            apply(enabled);
        }
    }
}
