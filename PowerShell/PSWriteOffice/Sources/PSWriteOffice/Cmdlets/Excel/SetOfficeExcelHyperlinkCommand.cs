using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets a hyperlink on a worksheet cell.</summary>
/// <example>
///   <summary>Set an external hyperlink.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelHyperlink -Address 'A1' -Url 'https://example.org' -Display 'Example' }</code>
///   <para>Creates a styled hyperlink in A1.</para>
/// </example>
/// <example>
///   <summary>Set an internal hyperlink.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelHyperlink -Row 2 -Column 1 -TargetSheet 'Summary' -TargetAddress 'A1' -Display 'Go to Summary' }</code>
///   <para>Links A2 to Summary!A1.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelHyperlink", DefaultParameterSetName = ParameterSetContextExternal)]
[Alias("ExcelHyperlink")]
public sealed class SetOfficeExcelHyperlinkCommand : PSCmdlet
{
    private const string ParameterSetContextExternal = "ContextExternal";
    private const string ParameterSetContextInternal = "ContextInternal";
    private const string ParameterSetDocumentExternal = "DocumentExternal";
    private const string ParameterSetDocumentInternal = "DocumentInternal";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentExternal)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentInternal)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentExternal)]
    [Parameter(ParameterSetName = ParameterSetDocumentInternal)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentExternal)]
    [Parameter(ParameterSetName = ParameterSetDocumentInternal)]
    public int? SheetIndex { get; set; }

    /// <summary>1-based row index.</summary>
    [Parameter]
    public int? Row { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter]
    public int? Column { get; set; }

    /// <summary>A1-style cell address (e.g., A1, C5).</summary>
    [Parameter]
    public string? Address { get; set; }

    /// <summary>External URL to link to.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetContextExternal)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentExternal)]
    public string Url { get; set; } = string.Empty;

    /// <summary>Internal location to link to (e.g., "'Summary'!A1").</summary>
    [Parameter(ParameterSetName = ParameterSetContextInternal)]
    [Parameter(ParameterSetName = ParameterSetDocumentInternal)]
    public string? Location { get; set; }

    /// <summary>Target worksheet name for internal links.</summary>
    [Parameter(ParameterSetName = ParameterSetContextInternal)]
    [Parameter(ParameterSetName = ParameterSetDocumentInternal)]
    public string? TargetSheet { get; set; }

    /// <summary>Target A1 address for internal links.</summary>
    [Parameter(ParameterSetName = ParameterSetContextInternal)]
    [Parameter(ParameterSetName = ParameterSetDocumentInternal)]
    public string? TargetAddress { get; set; }

    /// <summary>Optional display text.</summary>
    [Parameter]
    public string? Display { get; set; }

    /// <summary>Skip hyperlink styling (blue + underline).</summary>
    [Parameter]
    public SwitchParameter NoStyle { get; set; }

    /// <summary>Emit the worksheet after setting the link.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
        bool style = !NoStyle.IsPresent;

        if (ParameterSetName == ParameterSetContextExternal || ParameterSetName == ParameterSetDocumentExternal)
        {
            if (string.IsNullOrWhiteSpace(Url))
            {
                throw new PSArgumentException("Provide a URL.");
            }

            sheet.SetHyperlink(row, column, Url, Display, style);
        }
        else
        {
            var location = ResolveLocation();
            sheet.SetInternalLink(row, column, location, Display, style);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private string ResolveLocation()
    {
        if (!string.IsNullOrWhiteSpace(Location))
        {
            return Location!;
        }

        if (string.IsNullOrWhiteSpace(TargetSheet))
        {
            throw new PSArgumentException("Provide -Location or -TargetSheet for internal links.");
        }

        if (string.IsNullOrWhiteSpace(TargetAddress))
        {
            throw new PSArgumentException("Provide -TargetAddress for internal links.");
        }

        var escapedSheet = TargetSheet!.Replace("'", "''");
        return $"'{escapedSheet}'!{TargetAddress}";
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocumentExternal || ParameterSetName == ParameterSetDocumentInternal)
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
}
