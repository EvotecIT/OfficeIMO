using System;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets worksheet header and footer text and optional images.</summary>
/// <para>Uses OfficeIMO.Excel header/footer APIs and supports DSL or document usage.</para>
/// <example>
///   <summary>Set a centered header with a page footer.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelHeaderFooter -HeaderCenter 'Demo' -FooterRight 'Page &amp;P of &amp;N' }</code>
///   <para>Applies header and footer text to the worksheet.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelHeaderFooter", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelHeaderFooter")]
public sealed class SetOfficeExcelHeaderFooterCommand : AsyncPSCmdlet
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

    /// <summary>Left header text.</summary>
    [Parameter]
    public string? HeaderLeft { get; set; }

    /// <summary>Center header text.</summary>
    [Parameter]
    public string? HeaderCenter { get; set; }

    /// <summary>Right header text.</summary>
    [Parameter]
    public string? HeaderRight { get; set; }

    /// <summary>Left footer text.</summary>
    [Parameter]
    public string? FooterLeft { get; set; }

    /// <summary>Center footer text.</summary>
    [Parameter]
    public string? FooterCenter { get; set; }

    /// <summary>Right footer text.</summary>
    [Parameter]
    public string? FooterRight { get; set; }

    /// <summary>Use a different header/footer on the first page.</summary>
    [Parameter]
    public SwitchParameter DifferentFirstPage { get; set; }

    /// <summary>Use different headers/footers on odd/even pages.</summary>
    [Parameter]
    public SwitchParameter DifferentOddEven { get; set; }

    /// <summary>Align header/footer with margins (default: true).</summary>
    [Parameter]
    public bool AlignWithMargins { get; set; } = true;

    /// <summary>Scale header/footer with document (default: true).</summary>
    [Parameter]
    public bool ScaleWithDocument { get; set; } = true;

    /// <summary>Header image file path.</summary>
    [Parameter]
    public string? HeaderImagePath { get; set; }

    /// <summary>Header image URL.</summary>
    [Parameter]
    public string? HeaderImageUrl { get; set; }

    /// <summary>Header image position.</summary>
    [Parameter]
    public ExcelHeaderFooterPosition HeaderImagePosition { get; set; } = ExcelHeaderFooterPosition.Right;

    /// <summary>Footer image file path.</summary>
    [Parameter]
    public string? FooterImagePath { get; set; }

    /// <summary>Footer image URL.</summary>
    [Parameter]
    public string? FooterImageUrl { get; set; }

    /// <summary>Footer image position.</summary>
    [Parameter]
    public ExcelHeaderFooterPosition FooterImagePosition { get; set; } = ExcelHeaderFooterPosition.Right;

    /// <summary>Image width in points.</summary>
    [Parameter]
    public double? ImageWidthPoints { get; set; }

    /// <summary>Image height in points.</summary>
    [Parameter]
    public double? ImageHeightPoints { get; set; }

    /// <summary>Emit the worksheet after updating.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        var sheet = ResolveSheet();

        sheet.SetHeaderFooter(
            headerLeft: HeaderLeft,
            headerCenter: HeaderCenter,
            headerRight: HeaderRight,
            footerLeft: FooterLeft,
            footerCenter: FooterCenter,
            footerRight: FooterRight,
            differentFirstPage: DifferentFirstPage.IsPresent,
            differentOddEven: DifferentOddEven.IsPresent,
            alignWithMargins: AlignWithMargins,
            scaleWithDoc: ScaleWithDocument);

        if (!string.IsNullOrWhiteSpace(HeaderImagePath))
        {
            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(HeaderImagePath);
            var bytes = File.ReadAllBytes(resolved);
            sheet.SetHeaderImage(HeaderImagePosition, bytes, GetContentType(resolved), ImageWidthPoints, ImageHeightPoints);
        }
        else if (!string.IsNullOrWhiteSpace(HeaderImageUrl))
        {
            await sheet.SetHeaderImageFromUrlAsync(HeaderImagePosition, HeaderImageUrl!, ImageWidthPoints, ImageHeightPoints, CancelToken)
                .ConfigureAwait(false);
        }

        if (!string.IsNullOrWhiteSpace(FooterImagePath))
        {
            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(FooterImagePath);
            var bytes = File.ReadAllBytes(resolved);
            sheet.SetFooterImage(FooterImagePosition, bytes, GetContentType(resolved), ImageWidthPoints, ImageHeightPoints);
        }
        else if (!string.IsNullOrWhiteSpace(FooterImageUrl))
        {
            await sheet.SetFooterImageFromUrlAsync(FooterImagePosition, FooterImageUrl!, ImageWidthPoints, ImageHeightPoints, CancelToken)
                .ConfigureAwait(false);
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
            if (Document == null)
            {
                throw new PSArgumentException("Provide an Excel document.");
            }

            if (!string.IsNullOrWhiteSpace(Sheet))
            {
                return Document[Sheet!];
            }

            if (SheetIndex.HasValue)
            {
                if (SheetIndex.Value < 0 || SheetIndex.Value >= Document.Sheets.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(SheetIndex), "SheetIndex is out of range.");
                }
                return Document.Sheets[SheetIndex.Value];
            }

            if (Document.Sheets.Count == 0)
            {
                throw new InvalidOperationException("Workbook contains no worksheets.");
            }

            return Document.Sheets[0];
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }

    private static string GetContentType(string path)
    {
        var ext = Path.GetExtension(path)?.ToLowerInvariant();
        return ext switch
        {
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".tif" or ".tiff" => "image/tiff",
            _ => "image/png"
        };
    }
}
