using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Protects a worksheet and chooses what users may still do.</summary>
/// <example>
/// <summary>Protect a sheet</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
///     Protect-OfficeExcelSheet -Sheet Inputs
/// }</code>
/// <para>Locks the Inputs sheet so only unlocked cells can be edited.</para>
/// </example>
/// <example>
/// <summary>Protect a shared list but keep it usable</summary>
/// <code>Open-OfficeExcel -Path .\HR\Staff-Directory.xlsx {
///     ExcelSheet 'Directory' {
///         Protect-OfficeExcelSheet -AllowSort -AllowAutoFilter -AllowFormatColumns
///     }
/// }</code>
/// <para>Protects the staff list while still letting colleagues sort, filter and resize columns.</para>
/// </example>
/// <example>
/// <summary>Protect a sheet with a password</summary>
/// <code>$password = Read-Host -Prompt 'Sheet password' -AsSecureString
/// Open-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
///     Protect-OfficeExcelSheet -Sheet Inputs -Password $password -AllowSort -AllowAutoFilter
/// }</code>
/// <para>Locks the input sheet with a password read from the console, still allowing sorting and filtering.</para>
/// </example>
[Cmdlet(VerbsSecurity.Protect, "OfficeExcelSheet", DefaultParameterSetName = "Default")]
public sealed class ProtectOfficeExcelSheetCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Password used to open an encrypted workbook package.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Allow common protected-table workflows: selecting cells, inserting rows, sorting, and filtering.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowTableEditing { get; set; }

    /// <summary>Prevents selecting locked cells while the worksheet is protected.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoSelectLockedCells { get; set; }

    /// <summary>Prevents selecting unlocked cells while the worksheet is protected.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoSelectUnlockedCells { get; set; }

    /// <summary>Allow formatting cells.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowFormatCells { get; set; }

    /// <summary>Allow formatting columns.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowFormatColumns { get; set; }

    /// <summary>Allow formatting rows.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowFormatRows { get; set; }

    /// <summary>Allow inserting columns.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowInsertColumns { get; set; }

    /// <summary>Allow inserting rows.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowInsertRows { get; set; }

    /// <summary>Allow inserting hyperlinks.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowInsertHyperlinks { get; set; }

    /// <summary>Allow deleting columns.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowDeleteColumns { get; set; }

    /// <summary>Allow deleting rows.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowDeleteRows { get; set; }

    /// <summary>Allow sorting.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowSort { get; set; }

    /// <summary>Allow AutoFilter.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowAutoFilter { get; set; }

    /// <summary>Allow PivotTables.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowPivotTables { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        var options = AllowTableEditing ? ExcelSheetProtectionOptions.TableEditing() : new ExcelSheetProtectionOptions();
        if (MyInvocation.BoundParameters.ContainsKey(nameof(NoSelectLockedCells))) options.AllowSelectLockedCells = !NoSelectLockedCells;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(NoSelectUnlockedCells))) options.AllowSelectUnlockedCells = !NoSelectUnlockedCells;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowFormatCells))) options.AllowFormatCells = AllowFormatCells;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowFormatColumns))) options.AllowFormatColumns = AllowFormatColumns;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowFormatRows))) options.AllowFormatRows = AllowFormatRows;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowInsertColumns))) options.AllowInsertColumns = AllowInsertColumns;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowInsertRows))) options.AllowInsertRows = AllowInsertRows;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowInsertHyperlinks))) options.AllowInsertHyperlinks = AllowInsertHyperlinks;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowDeleteColumns))) options.AllowDeleteColumns = AllowDeleteColumns;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowDeleteRows))) options.AllowDeleteRows = AllowDeleteRows;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowSort))) options.AllowSort = AllowSort;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowAutoFilter))) options.AllowAutoFilter = AllowAutoFilter;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowPivotTables))) options.AllowPivotTables = AllowPivotTables;
        options.Password = Password.ToPlainText();
        sheet.Protect(options);
        if (PassThru) WriteObject(sheet);
    }

}
