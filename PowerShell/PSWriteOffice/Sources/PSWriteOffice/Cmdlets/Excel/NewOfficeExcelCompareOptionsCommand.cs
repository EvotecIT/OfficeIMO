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

/// <summary>Creates workbook comparison settings for Compare-OfficeDocument.</summary>
/// <example>
/// <summary>Compare values only</summary>
/// <code>Compare-OfficeDocument -ReferencePath .\Finance\Budget-v1.xlsx -DifferencePath .\Finance\Budget-v2.xlsx -Options (New-OfficeExcelCompareOptions -Exclude CellStyles)</code>
/// <para>Compares two versions of a workbook while ignoring cell style changes.</para>
/// </example>
/// <example>
/// <summary>Review content changes between releases</summary>
/// <code>$compare = New-OfficeExcelCompareOptions -Exclude CellStyles, Comments
/// Compare-OfficeDocument -ReferencePath .\Pricing\Price-List-2026-09.xlsx -DifferencePath .\Pricing\Price-List-2026-10.xlsx -Options $compare -OutputPath .\Pricing\Price-Changes-October.xlsx</code>
/// <para>Compares last month's and this month's price list, ignoring styles and comments, and saves the comparison report.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelCompareOptions", DefaultParameterSetName = "Default")]
public sealed class NewOfficeExcelCompareOptionsCommand : PSCmdlet {
    /// <summary>Worksheet names or wildcard patterns omitted from the navigation list.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelCompareArea[]? Exclude { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var excluded = new HashSet<ExcelCompareArea>(Exclude ?? Array.Empty<ExcelCompareArea>());
        WriteObject(new ExcelWorkbookDiffOptions {
            CompareCells = !excluded.Contains(ExcelCompareArea.Cells),
            CompareCellStyles = !excluded.Contains(ExcelCompareArea.CellStyles),
            CompareNamedRanges = !excluded.Contains(ExcelCompareArea.NamedRanges),
            CompareTables = !excluded.Contains(ExcelCompareArea.Tables),
            CompareWorksheetMetadata = !excluded.Contains(ExcelCompareArea.WorksheetMetadata),
            CompareComments = !excluded.Contains(ExcelCompareArea.Comments)
        });
    }

}
