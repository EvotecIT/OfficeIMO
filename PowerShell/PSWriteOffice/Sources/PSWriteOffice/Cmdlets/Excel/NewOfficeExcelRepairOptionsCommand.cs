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

/// <summary>Creates Excel repair settings for Repair-OfficeDocument.</summary>
/// <example>
/// <summary>Repair defined names and tables</summary>
/// <code>Repair-OfficeDocument -Path .\Reports\Broken.xlsx -OutputPath .\Reports\Repaired.xlsx -Options (New-OfficeExcelRepairOptions -Area DefinedNames, Tables)</code>
/// <para>Repairs only the defined names and tables of a damaged workbook and writes the result to a new file.</para>
/// </example>
/// <example>
/// <summary>Repair print settings across a folder</summary>
/// <code>$repair = New-OfficeExcelRepairOptions -Area SheetViews, PrintSettings
/// Get-ChildItem -Path \\files\Migrated -Filter *.xlsx |
///     Repair-OfficeDocument -Options $repair -PassThru</code>
/// <para>Fixes sheet views and print settings in every workbook of a migrated share, leaving formulas and drawings untouched.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelRepairOptions", DefaultParameterSetName = "Default")]
public sealed class NewOfficeExcelRepairOptionsCommand : PSCmdlet {
    /// <summary>Native workbook areas included in the operation; omitted selection uses the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelRepairArea[]? Area { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var selected = Area == null ? new HashSet<ExcelRepairArea>((ExcelRepairArea[])Enum.GetValues(typeof(ExcelRepairArea))) : new HashSet<ExcelRepairArea>(Area);
        WriteObject(new ExcelWorkbookRepairOptions {
            DefinedNames = selected.Contains(ExcelRepairArea.DefinedNames),
            Tables = selected.Contains(ExcelRepairArea.Tables),
            SheetViews = selected.Contains(ExcelRepairArea.SheetViews),
            PrintSettings = selected.Contains(ExcelRepairArea.PrintSettings),
            Protection = selected.Contains(ExcelRepairArea.Protection),
            AutoFilters = selected.Contains(ExcelRepairArea.AutoFilters),
            Drawings = selected.Contains(ExcelRepairArea.Drawings),
            LinksAndComments = selected.Contains(ExcelRepairArea.LinksAndComments),
            Calculation = selected.Contains(ExcelRepairArea.Calculation),
            WorkbookArtifacts = selected.Contains(ExcelRepairArea.WorkbookArtifacts),
            Save = false
        });
    }

}
