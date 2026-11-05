using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Compares two worksheets or ranges cell by cell.</summary>
/// <example>
/// <summary>Compare a sheet in two versions</summary>
/// <code>Compare-OfficeExcelRange -ReferencePath .\Finance\Budget-v1.xlsx -DifferencePath .\Finance\Budget-v2.xlsx -Sheet Budget |
///     Select-Object Kind, ReferenceCell, ReferenceValue, DifferenceValue</code>
/// <para>Lists the cells of the Budget sheet that differ between two versions of the workbook.</para>
/// </example>
/// <example>
/// <summary>Reconcile two lists in one workbook</summary>
/// <code>$compare = @{
///     ReferencePath   = '.\Finance\Reconciliation-October.xlsx'
///     Sheet           = 'Ledger'
///     Range           = 'A1:F400'
///     DifferenceSheet = 'Bank'
///     DifferenceRange = 'A1:F400'
/// }
/// Compare-OfficeExcelRange @compare -TrimStrings -IgnoreCase |
///     Export-OfficeExcel -Path .\Finance\Reconciliation-Differences.xlsx -Title 'Ledger and bank differences'</code>
/// <para>Compares the ledger with the bank statement range by range, ignoring case and stray spaces, and writes the differences to a review workbook.</para>
/// </example>
[Cmdlet(VerbsData.Compare, "OfficeExcelRange", DefaultParameterSetName = "Path")]
public sealed class CompareOfficeExcelRangeCommand : PSCmdlet {
    /// <summary>Path to the workbook containing the reference values.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    public string ReferencePath { get; set; } = null!;

    /// <summary>Workbook path to compare against.</summary>
    [Parameter(ParameterSetName = "Path", Position = 1)]
    public string? DifferencePath { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter()]
    public string? Range { get; set; }

    /// <summary>Worksheet selector in the difference workbook.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? DifferenceSheet { get; set; }

    /// <summary>A1 range containing the comparison values in the difference workbook.</summary>
    [Parameter()]
    public string? DifferenceRange { get; set; }

    /// <summary>Compare strings after trimming whitespace.</summary>
    [Parameter()]
    public SwitchParameter TrimStrings { get; set; }

    /// <summary>Compare strings case-insensitively.</summary>
    [Parameter()]
    public SwitchParameter IgnoreCase { get; set; }

    /// <summary>Treat null and empty string values as different.</summary>
    [Parameter()]
    public SwitchParameter StrictNullEmpty { get; set; }

    /// <summary>Caller-owned workbook containing the reference values.</summary>
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipeline = true)]
    public ExcelDocument ReferenceDocument { get; set; } = null!;

    /// <summary>Workbook document to compare against.</summary>
    [Parameter(ParameterSetName = "Document")]
    public ExcelDocument? DifferenceDocument { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var reference = ParameterSetName == "Path" ? ExcelWorkbookCommandService.OpenWorkbook(this, ReferencePath!, readOnly: true) : new ExcelWorkbookCommandScope(ReferenceDocument!, ownsDocument: false);
        using var difference = DifferencePath != null ? ExcelWorkbookCommandService.OpenWorkbook(this, DifferencePath, readOnly: true) : new ExcelWorkbookCommandScope(DifferenceDocument ?? reference.Document, ownsDocument: false);
        var referenceSheet = ExcelSheetResolver.Resolve(reference.Document, Sheet);
        var differenceSheet = DifferenceSheet == null ? difference.Document[referenceSheet.Name] : ExcelSheetResolver.Resolve(difference.Document, DifferenceSheet);
        var differences = reference.Document.CompareRanges(referenceSheet, Range ?? referenceSheet.UsedRange.Address, differenceSheet, DifferenceRange ?? differenceSheet.UsedRange.Address, new ExcelRangeCompareOptions { TrimStrings = TrimStrings, IgnoreCase = IgnoreCase, TreatNullAndEmptyStringAsEqual = !StrictNullEmpty });
        WriteObject(differences, enumerateCollection: true);
    }

}
