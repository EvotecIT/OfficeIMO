using System;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds an image anchored to a worksheet cell or range.</summary>
/// <example>
///   <summary>Insert a scaled image from disk at B2.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelImage -Address 'B2' -Path .\logo.png -ScalePercent 20 -Name Logo -AltText 'Company logo' }</code>
///   <para>Anchors the image to cell B2 and sizes it to 20 percent of the original image dimensions.</para>
/// </example>
/// <example>
///   <summary>Pin an image to a worksheet range.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelImage -Range 'A1:C15' -Path .\logo.png -Name HeaderLogo -Placement MoveAndSize }</code>
///   <para>Uses Excel's two-cell anchor so the picture moves and resizes with the cells in A1:C15.</para>
/// </example>
/// <example>
///   <summary>Insert and rotate an image from a URL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelImage -Row 1 -Column 1 -Url 'https://example.org/logo.png' -WidthPixels 96 -HeightPixels 32 -RotationDegrees 12 }</code>
///   <para>Downloads, sizes, rotates, and anchors the image to cell A1.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelImage", DefaultParameterSetName = ParameterSetContextPath)]
[Alias("ExcelImage")]
public sealed class AddOfficeExcelImageCommand : AsyncPSCmdlet
{
    private const string ParameterSetContextPath = "ContextPath";
    private const string ParameterSetContextUrl = "ContextUrl";
    private const string ParameterSetDocumentPath = "DocumentPath";
    private const string ParameterSetDocumentUrl = "DocumentUrl";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentPath)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentUrl)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentPath)]
    [Parameter(ParameterSetName = ParameterSetDocumentUrl)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocumentPath)]
    [Parameter(ParameterSetName = ParameterSetDocumentUrl)]
    public int? SheetIndex { get; set; }

    /// <summary>Image file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetContextPath)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentPath)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Image URL to download.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetContextUrl)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentUrl)]
    public string Url { get; set; } = string.Empty;

    /// <summary>1-based row index.</summary>
    [Parameter]
    public int? Row { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter]
    public int? Column { get; set; }

    /// <summary>A1-style cell address (e.g., A1, C5).</summary>
    [Parameter]
    [Alias("Cell")]
    public string? Address { get; set; }

    /// <summary>A1-style range (for example, A1:C15) for a two-cell anchor that can move and resize with cells.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Image width in pixels.</summary>
    [Parameter]
    [Alias("Width")]
    public int WidthPixels { get; set; } = 96;

    /// <summary>Image height in pixels.</summary>
    [Parameter]
    [Alias("Height")]
    public int HeightPixels { get; set; } = 32;

    /// <summary>Percentage of the original image size. Cannot be combined with WidthPixels or HeightPixels.</summary>
    [Parameter]
    public double? ScalePercent { get; set; }

    /// <summary>Horizontal offset in pixels from the cell origin.</summary>
    [Parameter]
    public int OffsetXPixels { get; set; }

    /// <summary>Vertical offset in pixels from the cell origin.</summary>
    [Parameter]
    public int OffsetYPixels { get; set; }

    /// <summary>Horizontal offset in pixels for the range end marker when using Range.</summary>
    [Parameter]
    public int EndOffsetXPixels { get; set; }

    /// <summary>Vertical offset in pixels for the range end marker when using Range.</summary>
    [Parameter]
    public int EndOffsetYPixels { get; set; }

    /// <summary>Optional drawing name used by Excel's selection pane.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional alternative text description for accessibility.</summary>
    [Parameter]
    public string? AltText { get; set; }

    /// <summary>Optional alternative text title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Marks the image as decorative by clearing alternative text metadata.</summary>
    [Parameter]
    public SwitchParameter Decorative { get; set; }

    /// <summary>Do not lock the image aspect ratio in Excel.</summary>
    [Parameter]
    public SwitchParameter NoLockAspectRatio { get; set; }

    /// <summary>Lock the image aspect ratio in Excel. This is the default unless NoLockAspectRatio is used.</summary>
    [Parameter]
    public SwitchParameter LockAspectRatio { get; set; }

    /// <summary>How a range-anchored image behaves when cells move or resize.</summary>
    [Parameter]
    public ExcelImagePlacement Placement { get; set; } = ExcelImagePlacement.MoveAndSize;

    /// <summary>Clockwise image rotation in degrees.</summary>
    [Parameter]
    public double RotationDegrees { get; set; }

    /// <summary>Emit the worksheet after inserting the image.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        ValidateImageOptions();

        var sheet = ResolveSheet();

        if (ParameterSetName == ParameterSetContextPath || ParameterSetName == ParameterSetDocumentPath)
        {
            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            AddLocalImage(sheet, resolved);
        }
        else
        {
            if (TryGetLocalFilePath(Url, out var localPath))
            {
                AddLocalImage(sheet, localPath);
            }
            else
            {
                await AddRemoteImageAsync(sheet, Url).ConfigureAwait(false);
            }
        }

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocumentPath || ParameterSetName == ParameterSetDocumentUrl)
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

    private void AddLocalImage(ExcelSheet sheet, string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Image file '{path}' was not found.", path);
        }

        ExcelImage image;
        if (!string.IsNullOrWhiteSpace(Range))
        {
            image = sheet.AddImageFromFileToRange(Range!, path, OffsetXPixels, OffsetYPixels, EndOffsetXPixels, EndOffsetYPixels,
                Name, Decorative.IsPresent ? null : AltText, Title, ResolveLockAspectRatio(), Placement, RotationDegrees);
        }
        else
        {
            var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
            var (width, height) = ResolveCellImageSize();
            image = sheet.AddImageFromFile(row, column, path, width, height, ScalePercent, OffsetXPixels, OffsetYPixels,
                Name, Decorative.IsPresent ? null : AltText, Title, ResolveLockAspectRatio(), RotationDegrees);
        }

        if (Decorative.IsPresent)
        {
            image.Decorative();
        }
    }

    private async Task AddRemoteImageAsync(ExcelSheet sheet, string url)
    {
        ExcelImage? image;
        if (!string.IsNullOrWhiteSpace(Range))
        {
            image = await sheet.AddImageFromUrlToRangeAsync(Range!, url, OffsetXPixels, OffsetYPixels, EndOffsetXPixels, EndOffsetYPixels,
                Name, Decorative.IsPresent ? null : AltText, Title, ResolveLockAspectRatio(), Placement, RotationDegrees, CancelToken)
                .ConfigureAwait(false);
        }
        else
        {
            var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
            var (width, height) = ResolveCellImageSize();
            image = await sheet.AddImageFromUrlAsync(row, column, url, width, height, ScalePercent, OffsetXPixels, OffsetYPixels,
                Name, Decorative.IsPresent ? null : AltText, Title, ResolveLockAspectRatio(), RotationDegrees, CancelToken)
                .ConfigureAwait(false);
        }

        if (Decorative.IsPresent)
        {
            image?.Decorative();
        }
    }

    private void ValidateImageOptions()
    {
        bool widthBound = MyInvocation.BoundParameters.ContainsKey(nameof(WidthPixels));
        bool heightBound = MyInvocation.BoundParameters.ContainsKey(nameof(HeightPixels));
        bool rangeBound = !string.IsNullOrWhiteSpace(Range);

        if (WidthPixels <= 0 || HeightPixels <= 0)
        {
            throw new PSArgumentException("WidthPixels and HeightPixels must be greater than zero.");
        }

        if (ScalePercent.HasValue && (ScalePercent.Value <= 0 || double.IsNaN(ScalePercent.Value) || double.IsInfinity(ScalePercent.Value)))
        {
            throw new PSArgumentException("ScalePercent must be a positive finite number.");
        }

        if (ScalePercent.HasValue && (widthBound || heightBound))
        {
            throw new PSArgumentException("ScalePercent cannot be combined with WidthPixels or HeightPixels.");
        }

        if (rangeBound && (Row.HasValue || Column.HasValue || !string.IsNullOrWhiteSpace(Address)))
        {
            throw new PSArgumentException("Use either Range or Row/Column/Address, not both.");
        }

        if (rangeBound && (ScalePercent.HasValue || widthBound || heightBound))
        {
            throw new PSArgumentException("Range determines the image size. Do not combine Range with ScalePercent, WidthPixels, or HeightPixels.");
        }

        if (Decorative.IsPresent && (!string.IsNullOrWhiteSpace(AltText) || !string.IsNullOrWhiteSpace(Title)))
        {
            throw new PSArgumentException("Decorative images cannot also define AltText or Title.");
        }
    }

    private (int? Width, int? Height) ResolveCellImageSize()
    {
        if (ScalePercent.HasValue)
        {
            return (null, null);
        }

        bool widthBound = MyInvocation.BoundParameters.ContainsKey(nameof(WidthPixels));
        bool heightBound = MyInvocation.BoundParameters.ContainsKey(nameof(HeightPixels));
        return (widthBound ? WidthPixels : 96, heightBound ? HeightPixels : 32);
    }

    private bool ResolveLockAspectRatio()
    {
        if (NoLockAspectRatio.IsPresent)
        {
            return false;
        }

        return !MyInvocation.BoundParameters.ContainsKey(nameof(LockAspectRatio)) || LockAspectRatio.IsPresent;
    }

    private static bool TryGetLocalFilePath(string url, out string path)
    {
        path = string.Empty;
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || !uri.IsFile)
        {
            return false;
        }

        path = uri.LocalPath;
        return !string.IsNullOrWhiteSpace(path);
    }
}
