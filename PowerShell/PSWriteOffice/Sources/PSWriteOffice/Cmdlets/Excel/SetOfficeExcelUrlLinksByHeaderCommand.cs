using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Converts cells under a header into external URL hyperlinks.</summary>
/// <example>
///   <summary>Link the RFC column in a table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Summary' { Set-OfficeExcelUrlLinksByHeader -Header 'RFC' -TableName 'Links' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } }</code>
///   <para>Uses the RFC column values to create external hyperlinks.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelUrlLinksByHeader", DefaultParameterSetName = ParameterSetContextUsedRange)]
[Alias("ExcelUrlLinksByHeader")]
[OutputType(typeof(ExcelSheet))]
public sealed class SetOfficeExcelUrlLinksByHeaderCommand : PSCmdlet
{
    private const string ParameterSetContextUsedRange = "ContextUsedRange";
    private const string ParameterSetDocumentUsedRange = "DocumentUsedRange";
    private const string ParameterSetContextTable = "ContextTable";
    private const string ParameterSetDocumentTable = "DocumentTable";
    private const string ParameterSetContextRange = "ContextRange";
    private const string ParameterSetDocumentRange = "DocumentRange";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentUsedRange)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentTable)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentRange)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentUsedRange)]
    [Parameter(ParameterSetName = ParameterSetDocumentTable)]
    [Parameter(ParameterSetName = ParameterSetDocumentRange)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentUsedRange)]
    [Parameter(ParameterSetName = ParameterSetDocumentTable)]
    [Parameter(ParameterSetName = ParameterSetDocumentRange)]
    public int? SheetIndex { get; set; }

    /// <summary>Header text to locate.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Header { get; set; } = string.Empty;

    /// <summary>Restrict linking to a named table.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetContextTable)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentTable)]
    public string TableName { get; set; } = string.Empty;

    /// <summary>Restrict linking to a specific A1 range whose first row contains headers.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetContextRange)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentRange)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Maps the cell text to a URL.</summary>
    [Parameter(Mandatory = true)]
    public ScriptBlock UrlScript { get; set; } = null!;

    /// <summary>Optional mapping from cell text to display text.</summary>
    [Parameter]
    public ScriptBlock? TitleScript { get; set; }

    /// <summary>Skip hyperlink styling (blue + underline).</summary>
    [Parameter]
    public SwitchParameter NoStyle { get; set; }

    /// <summary>Emit the worksheet after creating links.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        var urlForText = new Func<string, string>(text => ExcelTextTransformService.Apply(UrlScript, text));
        var titleForText = TitleScript == null ? null : new Func<string, string>(text => ExcelTextTransformService.Apply(TitleScript, text));

        switch (ParameterSetName)
        {
            case ParameterSetContextTable:
            case ParameterSetDocumentTable:
                var tableRange = TryGetContextTableRange(sheet);
                if (tableRange == null || tableRange.Length == 0)
                {
                    sheet.LinkByHeaderToUrlsInTable(TableName, Header, urlForText, titleForText, !NoStyle.IsPresent);
                }
                else
                {
                    sheet.LinkByHeaderToUrlsInRange(tableRange, Header, urlForText, titleForText, !NoStyle.IsPresent);
                }

                break;
            case ParameterSetContextRange:
            case ParameterSetDocumentRange:
                sheet.LinkByHeaderToUrlsInRange(Range, Header, urlForText, titleForText, !NoStyle.IsPresent);
                break;
            default:
                sheet.LinkByHeaderToUrls(Header, urlForCellText: urlForText, titleForCellText: titleForText, styled: !NoStyle.IsPresent);
                break;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocumentUsedRange ||
            ParameterSetName == ParameterSetDocumentTable ||
            ParameterSetName == ParameterSetDocumentRange)
        {
            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        return ExcelDslContext.Require(this).RequireSheet();
    }

    private string? TryGetContextTableRange(ExcelSheet sheet)
    {
        if (ParameterSetName != ParameterSetContextTable)
        {
            return null;
        }

        try
        {
            var context = ExcelDslContext.Require(this);
            return context.TryGetTableRange(sheet, TableName, out var range) ? range : null;
        }
        catch (InvalidOperationException)
        {
            return null;
        }
    }
}
