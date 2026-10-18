using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Overlays or underlays one source PDF page on selected pages of another PDF.</summary>
/// <example>
///   <summary>Place a letterhead page behind every report page.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficePdfPageOverlay -Path .\Report.pdf -SourcePath .\Letterhead.pdf `
///     -SourcePageNumber 1 -Underlay -Opacity 0.9 -OutputPath .\BrandedReport.pdf</code>
///   <para>Imports the first source page once and places it behind each target page.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfPageOverlay", SupportsShouldProcess = true)]
[Alias("PdfPageOverlay")]
[OutputType(typeof(FileInfo))]
public sealed class AddOfficePdfPageOverlayCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Target PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>PDF containing the page to import.</summary>
    [Parameter(Mandatory = true)]
    public string SourcePath { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>One-based page number imported from the source PDF.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int SourcePageNumber { get; set; } = 1;

    /// <summary>Target page selector such as 1-3,odd,last. Omit to apply to every target page.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>How the source page fits the target page or explicit rectangle.</summary>
    [Parameter]
    public PdfPageOverlayFit Fit { get; set; } = PdfPageOverlayFit.Contain;

    /// <summary>Horizontal placement inside the target page or rectangle.</summary>
    [Parameter]
    public PdfAlign HorizontalAlign { get; set; } = PdfAlign.Center;

    /// <summary>Vertical placement inside the target page or rectangle.</summary>
    [Parameter]
    public PdfVerticalAlign VerticalAlign { get; set; } = PdfVerticalAlign.Middle;

    /// <summary>Optional target rectangle X coordinate in PDF points.</summary>
    [Parameter]
    public double? X { get; set; }

    /// <summary>Optional target rectangle Y coordinate in PDF points.</summary>
    [Parameter]
    public double? Y { get; set; }

    /// <summary>Optional target rectangle width in PDF points.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Optional target rectangle height in PDF points.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Opacity of the imported source page.</summary>
    [Parameter]
    [ValidateRange(0D, 1D)]
    public double Opacity { get; set; } = 1D;

    /// <summary>Place the imported page behind existing target-page content.</summary>
    [Parameter]
    public SwitchParameter Underlay { get; set; }

    /// <summary>Optional bounded read settings for the target PDF.</summary>
    [Parameter]
    public PdfLoadOptions? ReadOptions { get; set; }

    /// <summary>Password used to authenticate the target PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After target authentication, explicitly ignore owner-imposed target restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Password used to authenticate the imported source PDF.</summary>
    [Parameter]
    public string? SourcePassword { get; set; }

    /// <summary>After source authentication, explicitly ignore owner-imposed source restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnoreSourcePermissionRestrictions { get; set; }

    /// <summary>Optional bounded read settings for the imported source PDF.</summary>
    [Parameter]
    public PdfLoadOptions? SourceReadOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write PDF page overlay")) {
            return;
        }

        var targetReadOptions = PdfCommandUtilities.CreateReadOptions(
            ReadOptions,
            Password,
            IgnorePermissionRestrictions.IsPresent);
        var sourceReadOptions = PdfCommandUtilities.CreateReadOptions(
            SourceReadOptions,
            SourcePassword,
            IgnoreSourcePermissionRestrictions.IsPresent);
        var options = new PdfPageOverlayOptions {
            SourcePageNumber = SourcePageNumber,
            Fit = Fit,
            HorizontalAlignment = HorizontalAlign,
            VerticalAlignment = VerticalAlign,
            X = X,
            Y = Y,
            Width = Width,
            Height = Height,
            Opacity = Opacity,
            SourceReadOptions = sourceReadOptions
        };
        if (!string.IsNullOrWhiteSpace(PageRange)) {
            options.UseTargetPages(PageRange!);
        }

        var document = PdfCommandUtilities.LoadDocument(
            PdfCommandUtilities.ResolvePath(this, Path),
            targetReadOptions);
        var sourcePath = PdfCommandUtilities.ResolvePath(this, SourcePath);
        using var sourceStream = PdfCommandUtilities.OpenBoundedReadStream(sourcePath, sourceReadOptions);
        var result = Underlay.IsPresent
            ? document.Stamp.UnderlayPage(sourceStream, options, targetReadOptions)
            : document.Stamp.OverlayPage(sourceStream, options, targetReadOptions);

        PdfCommandUtilities.EnsureDirectory(outputPath);
        result.Save(outputPath).RequireSuccess();
        WritePassThru(new FileInfo(outputPath));
    }
}
