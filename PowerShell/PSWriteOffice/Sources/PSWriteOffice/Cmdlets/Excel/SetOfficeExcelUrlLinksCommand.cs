using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Converts cells in a range into external URL hyperlinks.</summary>
/// <example>
///   <summary>Link a summary range to RFC pages.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Summary' { Set-OfficeExcelUrlLinks -Range 'D2:D10' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } }</code>
///   <para>Turns each non-empty cell in D2:D10 into an external hyperlink.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelUrlLinks", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelUrlLinks")]
[OutputType(typeof(ExcelSheet))]
public sealed class SetOfficeExcelUrlLinksCommand : PSCmdlet
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

    /// <summary>A1 range containing values to convert into external links.</summary>
    [Parameter(Mandatory = true, Position = 0)]
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
        var (rowFrom, columnFrom, rowTo, columnTo) = sheet.GetRangeBounds(Range);
        bool styled = !NoStyle.IsPresent;

        for (int row = rowFrom; row <= rowTo; row++)
        {
            for (int column = columnFrom; column <= columnTo; column++)
            {
                if (!sheet.TryGetCellText(row, column, out var text) || string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                string url = ExcelTextTransformService.Apply(UrlScript, text);
                if (string.IsNullOrWhiteSpace(url))
                {
                    continue;
                }

                if (TitleScript == null)
                {
                    sheet.SetHyperlinkSmart(row, column, url, null, styled);
                }
                else
                {
                    string title = ExcelTextTransformService.Apply(TitleScript, text);
                    sheet.SetHyperlink(row, column, url, title, styled);
                }
            }
        }

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        return ExcelDslContext.Require(this).RequireSheet();
    }
}
