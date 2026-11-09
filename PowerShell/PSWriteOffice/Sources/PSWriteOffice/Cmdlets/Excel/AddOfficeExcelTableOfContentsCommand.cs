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

/// <summary>Adds or refreshes a table-of-contents sheet with links to every sheet.</summary>
/// <example>
/// <summary>Add a contents sheet</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Board-Pack.xlsx {
///     ExcelTableOfContents
/// }</code>
/// <para>Adds a first sheet with a link to every other sheet.</para>
/// </example>
/// <example>
/// <summary>Build a navigable report pack</summary>
/// <code>$orders = Import-OfficeCsv -Path .\Data\Q3-Orders.csv
/// New-OfficeExcel -Path .\Reports\Q3-Regions.xlsx {
///     foreach ($group in $orders | Group-Object Region) {
///         ExcelSheet $group.Name {
///             ExcelTable $group.Group -Address A3
///         }
///     }
///     ExcelTableOfContents -Name 'Contents' -BackLink -BackLinkAddress A1 -BackLinkText 'Back to contents'
/// }</code>
/// <para>Writes one sheet per region, then adds a contents sheet with links and a back link on each regional sheet.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelTableOfContents", DefaultParameterSetName = "Default")]
[Alias("ExcelTableOfContents")]
public sealed class AddOfficeExcelTableOfContentsCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Named range name.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Name { get; set; }

    /// <summary>Keeps the table-of-contents worksheet at its current position.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoMoveFirst { get; set; }

    /// <summary>Disable internal hyperlinks in the TOC sheet.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoHyperlinks { get; set; }

    /// <summary>Include named ranges in the TOC.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeNamedRanges { get; set; }

    /// <summary>Include hidden named ranges when listing named ranges.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeHiddenNamedRanges { get; set; }

    /// <summary>Disable formatted TOC styling.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoStyle { get; set; }

    /// <summary>Adds a link from each listed worksheet back to the table of contents.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter BackLink { get; set; }

    /// <summary>A1 destination for worksheet back links; omission chooses an unused cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? BackLinkAddress { get; set; }

    /// <summary>Text used for back links.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? BackLinkText { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        string name = Name ?? "TOC";
        document.AddTableOfContents(sheetName: name, placeFirst: !NoMoveFirst, withHyperlinks: !NoHyperlinks, includeNamedRanges: IncludeNamedRanges, includeHiddenNamedRanges: IncludeHiddenNamedRanges, styled: !NoStyle);
        if (BackLink) {
            if (BackLinkAddress != null) {
                var (row, column) = ExcelHostExtensions.ResolveCellAddress(null, null, BackLinkAddress);
                document.AddBackLinksToToc(name, row, column, BackLinkText ?? "← TOC");
            } else foreach (var sheet in document.Sheets) {
                if (sheet.Name == name) continue;
                var (_, _, endRow, _) = ExcelA1.ParseRange(sheet.UsedRange.Address);
                sheet.SetInternalLink(Math.Max(1, endRow + 2), 1, document[name], "A1", BackLinkText ?? "← TOC");
            }
        }
        if (PassThru) WriteObject(document);
    }

}
