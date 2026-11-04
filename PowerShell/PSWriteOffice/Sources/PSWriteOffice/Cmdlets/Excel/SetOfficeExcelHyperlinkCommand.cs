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

/// <summary>Links a cell, or every cell of a range or header column, to a URL or a place in the workbook.</summary>
/// <example>
/// <summary>Link a cell to a web page</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Status.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelHyperlink A2 'https://status.example.com' -Display 'Service status'
///     }
/// }</code>
/// <para>Turns A2 into a link to the status page with readable text.</para>
/// </example>
/// <example>
/// <summary>Link every ticket number to the helpdesk</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
///     ExcelSheet 'Tickets' {
///         ExcelHyperlink -Header Ticket -Table Tickets -UrlScript { param($text) "https://helpdesk.example.com/tickets/$text" }
///     }
/// }</code>
/// <para>Links each value in the Ticket column of the Tickets table to its page in the helpdesk system.</para>
/// </example>
/// <example>
/// <summary>Link a summary to the detail sheets</summary>
/// <code>$servers = 'SQL01', 'SQL02', 'WEB01'
/// Open-OfficeExcel -Path .\Reports\Disk-Capacity.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelHyperlink -Header Server -Internal
///     }
///     foreach ($server in $servers) {
///         ExcelSheet $server {
///             ExcelHyperlink H1 -TargetSheet Summary -TargetAddress A1 -Display 'Back to summary'
///         }
///     }
/// }</code>
/// <para>Links each server name in the Summary sheet to the sheet of the same name, and adds a link back to the summary on every server sheet.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelHyperlink", DefaultParameterSetName = "Url")]
[Alias("ExcelHyperlink")]
public sealed class SetOfficeExcelHyperlinkCommand : PSCmdlet {
    /// <summary>A1 address of the target cell.</summary>
    [Parameter(ParameterSetName = "Url", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Internal", Mandatory = true, Position = 0)]
    public string Address { get; set; } = null!;

    /// <summary>External URL to link to.</summary>
    [Parameter(ParameterSetName = "Url", Mandatory = true, Position = 1)]
    public string Url { get; set; } = null!;

    /// <summary>Optional display text.</summary>
    [Parameter(ParameterSetName = "Url")]
    [Parameter(ParameterSetName = "Internal")]
    public string? Display { get; set; }

    /// <summary>Selects original text, full URL, host name or smart hyperlink display.</summary>
    [Parameter(ParameterSetName = "Url")]
    public ExcelHyperlinkDisplay? DisplayMode { get; set; }

    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Skip hyperlink styling (blue + underline).</summary>
    [Parameter()]
    public SwitchParameter NoStyle { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Target worksheet name for internal links.</summary>
    [Parameter(ParameterSetName = "Internal")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? TargetSheet { get; set; }

    /// <summary>Target A1 address for internal links.</summary>
    [Parameter(ParameterSetName = "Internal", Mandatory = true)]
    [Parameter(ParameterSetName = "InternalRange")]
    [Parameter(ParameterSetName = "InternalHeader")]
    public string TargetAddress { get; set; } = null!;

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "UrlRange", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "UrlHeader")]
    [Parameter(ParameterSetName = "InternalRange", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "InternalHeader")]
    public string Range { get; set; } = null!;

    /// <summary>Maps the cell text to a URL.</summary>
    [Parameter(ParameterSetName = "UrlRange", Mandatory = true)]
    [Parameter(ParameterSetName = "UrlHeader", Mandatory = true)]
    public ScriptBlock UrlScript { get; set; } = null!;

    /// <summary>Optional mapping from cell text to display text.</summary>
    [Parameter(ParameterSetName = "UrlRange")]
    [Parameter(ParameterSetName = "UrlHeader")]
    [Parameter(ParameterSetName = "InternalRange")]
    [Parameter(ParameterSetName = "InternalHeader")]
    public ScriptBlock? DisplayScript { get; set; }

    /// <summary>Header text identifying the target column.</summary>
    [Parameter(ParameterSetName = "UrlHeader", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "InternalHeader", Mandatory = true, Position = 0)]
    public string Header { get; set; } = null!;

    /// <summary>Native Excel table used to resolve the source range or header.</summary>
    [Parameter(ParameterSetName = "UrlHeader")]
    [Parameter(ParameterSetName = "InternalHeader")]
    public string? Table { get; set; }

    /// <summary>Script that maps each source cell text to a destination worksheet name.</summary>
    [Parameter(ParameterSetName = "InternalRange")]
    [Parameter(ParameterSetName = "InternalHeader")]
    public ScriptBlock? TargetSheetScript { get; set; }

    /// <summary>Uses worksheet destinations instead of external URLs for header-based links.</summary>
    [Parameter(ParameterSetName = "InternalRange", Mandatory = true)]
    [Parameter(ParameterSetName = "InternalHeader", Mandatory = true)]
    public SwitchParameter Internal { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        bool styled = !NoStyle;
        Func<string, string>? display = DisplayScript == null ? null : text => ExcelTextTransformService.Apply(DisplayScript, text);
        Func<string, string> destination = TargetSheetScript == null ? text => text : text => ExcelTextTransformService.Apply(TargetSheetScript, text);
        if (ParameterSetName.EndsWith("Header", StringComparison.Ordinal)) {
            var options = new ExcelLinkByHeaderOptions { Display = display, Styled = styled, Table = Table, Range = Range };
            var result = ParameterSetName.StartsWith("Url", StringComparison.Ordinal)
                ? sheet.LinkByHeader(Header!, ExcelLinkTarget.Url(text => ExcelTextTransformService.Apply(UrlScript!, text)), options)
                : sheet.LinkByHeader(Header!, ExcelLinkTarget.Sheet(destination, TargetAddress ?? "A1"), options);
            if (!result.Succeeded) throw new PSArgumentException(result.Failure);
        } else if (ParameterSetName.EndsWith("Range", StringComparison.Ordinal)) {
            var (startRow, startColumn, endRow, endColumn) = ExcelA1.ParseRange(Range!);
            for (int row = startRow; row <= endRow; row++)
                for (int column = startColumn; column <= endColumn; column++) {
                    string text = sheet.Cell(row, column).GetFormattedText() ?? string.Empty;
                    if (text.Length == 0) continue;
                    if (ParameterSetName.StartsWith("Url", StringComparison.Ordinal)) {
                        string url = ExcelTextTransformService.Apply(UrlScript!, text);
                        if (display == null) sheet.SetHyperlinkSmart(row, column, url, null, styled);
                        else sheet.SetHyperlink(row, column, url, display(text), styled);
                    } else sheet.SetInternalLink(row, column, sheet.Document.Sheets.Single(item => item.Name == destination(text)), TargetAddress ?? "A1", display?.Invoke(text), styled);
                }
        } else {
            var (row, column) = ExcelHostExtensions.ResolveCellAddress(null, null, Address);
            if (ParameterSetName == "Internal") {
                if (TargetSheet == null) sheet.SetInternalLink(row, column, TargetAddress!, Display, styled);
                else sheet.SetInternalLink(row, column, TargetSheet.Resolve(sheet.Document.Sheets, s => s.Name), TargetAddress!, Display, styled);
            } else switch (DisplayMode ?? ExcelHyperlinkDisplay.Text) {
                case ExcelHyperlinkDisplay.Host: sheet.SetHyperlinkHost(row, column, Url!, styled); break;
                case ExcelHyperlinkDisplay.Smart: sheet.SetHyperlinkSmart(row, column, Url!, Display, styled); break;
                case ExcelHyperlinkDisplay.Url: sheet.SetHyperlink(row, column, Url!, Url, styled); break;
                default: sheet.SetHyperlink(row, column, Url!, Display, styled); break;
            }
        }
        if (PassThru) WriteObject(sheet);
    }

}
