using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown.Pdf;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Creates discoverable Markdown-to-PDF conversion options for Export-OfficeDocumentPdf.</summary>
/// <example>
///   <summary>Allow local report images and apply PDF metadata.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeMarkdownPdfOptions -Title 'Service report' -Author 'Evotec' -IncludeLocalImages -BaseDirectory .\Assets
/// Export-OfficeDocumentPdf -InputPath .\Report.md -Path .\Report.pdf -MarkdownOptions $options</code>
///   <para>Builds a typed options object through ordinary PowerShell parameters; no hashtable or .NET construction is required.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeMarkdownPdfOptions")]
[OutputType(typeof(MarkdownToPdfOptions))]
public sealed class NewOfficeMarkdownPdfOptionsCommand : PSCmdlet, IMarkdownPdfOptionSource {
    /// <summary>Existing Markdown PDF options to clone and override.</summary>
    [Parameter(ValueFromPipeline = true)]
    public MarkdownToPdfOptions? Options { get; set; }

    /// <summary>Underlying low-level OfficeIMO PDF options.</summary>
    [Parameter]
    public OfficeIMO.Pdf.PdfOptions? PdfOptions { get; set; }

    /// <summary>Built-in visual theme.</summary>
    [Parameter]
    public OfficeVisualThemeKind? Theme { get; set; }

    /// <summary>Default font family.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>PDF title metadata.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>PDF author metadata.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>PDF subject metadata.</summary>
    [Parameter]
    public string? Subject { get; set; }

    /// <summary>PDF keywords metadata.</summary>
    [Parameter]
    public string? Keywords { get; set; }

    /// <summary>Base directory used to resolve local Markdown images.</summary>
    [Parameter]
    public string? BaseDirectory { get; set; }

    /// <summary>Apply the built-in Word-like Markdown PDF baseline theme.</summary>
    [Parameter]
    public SwitchParameter ApplyWordLikeTheme { get; set; }

    /// <summary>Embed supported local image files.</summary>
    [Parameter]
    public SwitchParameter IncludeLocalImages { get; set; }

    /// <summary>Embed supported data URI images.</summary>
    [Parameter]
    public SwitchParameter IncludeDataUriImages { get; set; }

    /// <summary>Require local images to resolve under BaseDirectory.</summary>
    [Parameter]
    public SwitchParameter RestrictLocalImagesToBaseDirectory { get; set; }

    /// <summary>Maximum decoded bytes for one data URI image.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaximumDataUriImageBytes { get; set; }

    /// <summary>Fallback image width in PDF points.</summary>
    [Parameter]
    [ValidateRange(double.Epsilon, double.MaxValue)]
    public double? DefaultImageWidth { get; set; }

    /// <summary>Fallback image height in PDF points.</summary>
    [Parameter]
    [ValidateRange(double.Epsilon, double.MaxValue)]
    public double? DefaultImageHeight { get; set; }

    /// <summary>Controls how YAML front matter appears in the PDF body.</summary>
    [Parameter]
    public MarkdownPdfFrontMatterRenderMode? FrontMatterRenderMode { get; set; }

    /// <summary>Use front matter values to select a visual theme.</summary>
    [Parameter]
    public SwitchParameter UseFrontMatterVisualTheme { get; set; }

    /// <summary>Use front matter values as PDF metadata.</summary>
    [Parameter]
    public SwitchParameter UseFrontMatterMetadata { get; set; }

    /// <summary>Use the first Markdown heading as the PDF title when no title is supplied.</summary>
    [Parameter]
    public SwitchParameter UseFirstHeadingAsTitle { get; set; }

    /// <summary>Create PDF outlines from Markdown headings.</summary>
    [Parameter]
    public SwitchParameter CreateOutlineFromHeadings { get; set; }

    MarkdownToPdfOptions? IMarkdownPdfOptionSource.MarkdownPdfOptions => Options;
    OfficeIMO.Pdf.PdfOptions? IMarkdownPdfOptionSource.PdfOptions => PdfOptions;
    OfficeVisualThemeKind? IMarkdownPdfOptionSource.PdfTheme => Theme;
    string? IMarkdownPdfOptionSource.PdfFontFamily => FontFamily;
    string? IMarkdownPdfOptionSource.PdfTitle => Title;
    string? IMarkdownPdfOptionSource.PdfAuthor => Author;
    string? IMarkdownPdfOptionSource.PdfSubject => Subject;
    string? IMarkdownPdfOptionSource.PdfKeywords => Keywords;
    string? IMarkdownPdfOptionSource.PdfBaseDirectory => BaseDirectory;
    bool? IMarkdownPdfOptionSource.PdfApplyWordLikeTheme => GetBoundSwitch(nameof(ApplyWordLikeTheme), ApplyWordLikeTheme);
    bool? IMarkdownPdfOptionSource.PdfIncludeLocalImages => GetBoundSwitch(nameof(IncludeLocalImages), IncludeLocalImages);
    bool? IMarkdownPdfOptionSource.PdfIncludeDataUriImages => GetBoundSwitch(nameof(IncludeDataUriImages), IncludeDataUriImages);
    bool? IMarkdownPdfOptionSource.PdfRestrictLocalImagesToBaseDirectory => GetBoundSwitch(nameof(RestrictLocalImagesToBaseDirectory), RestrictLocalImagesToBaseDirectory);
    int? IMarkdownPdfOptionSource.PdfMaximumDataUriImageBytes => MaximumDataUriImageBytes;
    double? IMarkdownPdfOptionSource.PdfDefaultImageWidth => DefaultImageWidth;
    double? IMarkdownPdfOptionSource.PdfDefaultImageHeight => DefaultImageHeight;
    MarkdownPdfFrontMatterRenderMode? IMarkdownPdfOptionSource.PdfFrontMatterRenderMode => FrontMatterRenderMode;
    bool? IMarkdownPdfOptionSource.PdfUseFrontMatterVisualTheme => GetBoundSwitch(nameof(UseFrontMatterVisualTheme), UseFrontMatterVisualTheme);
    bool? IMarkdownPdfOptionSource.PdfUseFrontMatterMetadata => GetBoundSwitch(nameof(UseFrontMatterMetadata), UseFrontMatterMetadata);
    bool? IMarkdownPdfOptionSource.PdfUseFirstHeadingAsTitle => GetBoundSwitch(nameof(UseFirstHeadingAsTitle), UseFirstHeadingAsTitle);
    bool? IMarkdownPdfOptionSource.PdfCreateOutlineFromHeadings => GetBoundSwitch(nameof(CreateOutlineFromHeadings), CreateOutlineFromHeadings);
    string? IMarkdownPdfOptionSource.PdfWarningVariable => null;
    string? IMarkdownPdfOptionSource.PdfConversionReportVariable => null;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WriteObject(MarkdownOptionUtilities.BuildPdfOptions(this, this, fallbackBaseDirectory: null));
    }

    private bool? GetBoundSwitch(string name, SwitchParameter value) =>
        MyInvocation.BoundParameters.ContainsKey(name) ? value.IsPresent : null;
}
