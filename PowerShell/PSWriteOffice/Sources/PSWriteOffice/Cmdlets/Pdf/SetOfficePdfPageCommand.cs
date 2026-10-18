using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets page-level PDF properties and writes a new PDF.</summary>
/// <example>
///   <summary>Rotate selected PDF pages.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Set-OfficePdfPage -Path .\Examples\Documents\Scanned.pdf -PageRange '2,4' -Rotation 90 -OutputPath .\Examples\Documents\Scanned-Rotated.pdf
///     Get-OfficePdfInfo -Path .\Examples\Documents\Scanned-Rotated.pdf | Select-Object PageCount
/// )
/// $proof</code>
///   <para>Rotates selected pages and writes a new PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfPage", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class SetOfficePdfPageCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Page ranges such as 1-3,5. Omit to affect all pages.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Rotation in degrees. Supported values are 0, 90, 180, and 270.</summary>
    [Parameter]
    [ValidateSet("0", "90", "180", "270")]
    public int Rotation { get; set; }

    /// <summary>Page boundary box to set. Supported values are MediaBox, CropBox, BleedBox, TrimBox, and ArtBox.</summary>
    [Parameter]
    public PdfPageBoundaryBox? BoxName { get; set; }

    /// <summary>Left coordinate for the page boundary box.</summary>
    [Parameter]
    public double? Left { get; set; }

    /// <summary>Bottom coordinate for the page boundary box.</summary>
    [Parameter]
    public double? Bottom { get; set; }

    /// <summary>Right coordinate for the page boundary box.</summary>
    [Parameter]
    public double? Right { get; set; }

    /// <summary>Top coordinate for the page boundary box.</summary>
    [Parameter]
    public double? Top { get; set; }

    /// <summary>Resize selected pages to a known OfficeIMO page size such as A4, Letter, or Custom.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficePdfPageSizeArgumentCompleter))]
    public string? PageSize { get; set; }

    /// <summary>Custom page width in points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Custom page height in points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Use the landscape orientation of the selected page size.</summary>
    [Parameter]
    public SwitchParameter Landscape { get; set; }

    /// <summary>How source page content is fitted into the resized output page.</summary>
    [Parameter]
    public PdfPageResizeMode ResizeMode { get; set; } = PdfPageResizeMode.Fit;

    /// <summary>Margin, in points, reserved around resized page content.</summary>
    [Parameter]
    public double? ResizeMargin { get; set; }

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed page-modification restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        int[] pages = string.IsNullOrWhiteSpace(PageRange)
            ? Array.Empty<int>()
            : PdfPageRange.ParseMany(PageRange!).SelectMany(ExpandPageRange).Distinct().ToArray();
        var resizeOptions = PdfCommandUtilities.CreatePageResizeOptions(
            PageSize,
            Width,
            Height,
            Landscape.IsPresent,
            ResizeMode,
            ResizeMargin,
            MyInvocation.BoundParameters.ContainsKey(nameof(PageSize)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Width)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Height)) ||
            Landscape.IsPresent ||
            MyInvocation.BoundParameters.ContainsKey(nameof(ResizeMode)) ||
            ResizeMargin.HasValue);
        ValidateOperation(resizeOptions);

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write updated PDF pages")) {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        string inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        var readOptions = PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent);

        if (resizeOptions != null) {
            PdfDocument.Load(inputPath, readOptions).Pages.Resize(resizeOptions, pages).Save(outputPath).RequireSuccess();
            WritePassThru(new FileInfo(outputPath));
            return;
        }

        if (BoxName is PdfPageBoundaryBox boxName &&
            Left is double left &&
            Bottom is double bottom &&
            Right is double right &&
            Top is double top) {
            PdfDocument
                .Load(inputPath, readOptions)
                .Pages.SetPageBox(
                    boxName,
                    left,
                    bottom,
                    right,
                    top,
                    pages)
                .Save(outputPath)
                .RequireSuccess();
            WritePassThru(new FileInfo(outputPath));
            return;
        }

        var document = PdfDocument.Load(inputPath, readOptions);
        var result = string.IsNullOrWhiteSpace(PageRange)
            ? document.Pages.Rotate(Rotation)
            : document.Pages.Rotate(Rotation, PageRange!);
        result.Save(outputPath).RequireSuccess();
        WritePassThru(new FileInfo(outputPath));
    }

    private void ValidateOperation(PdfPageResizeOptions? resizeOptions) {
        if (resizeOptions != null) {
            if (BoxName.HasValue || MyInvocation.BoundParameters.ContainsKey(nameof(Rotation))) {
                throw new PSArgumentException("Use page resize, rotation, or box editing as separate Set-OfficePdfPage operations.");
            }

            return;
        }

        if (BoxName.HasValue) {
            if (!Left.HasValue || !Bottom.HasValue || !Right.HasValue || !Top.HasValue) {
                throw new PSArgumentException("-BoxName requires -Left, -Bottom, -Right, and -Top.");
            }

            return;
        }

        if (!MyInvocation.BoundParameters.ContainsKey(nameof(Rotation))) {
            throw new PSArgumentException("Provide -Rotation, -BoxName with coordinates, or page resize options.");
        }
    }

    private static int[] ExpandPageRange(PdfPageRange range) {
        return Enumerable.Range(range.FirstPage, range.PageCount).ToArray();
    }
}
