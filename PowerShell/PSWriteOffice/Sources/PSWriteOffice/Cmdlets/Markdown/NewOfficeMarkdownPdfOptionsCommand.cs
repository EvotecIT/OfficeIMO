using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Markdown;
using PSWriteOffice.Foundation;
using OfficeIMO.Markdown.Pdf;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Creates native Markdown-to-PDF conversion options for Convert-OfficeDocument.</summary>
/// <example>
///   <summary>Allow local report images and apply PDF metadata.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeMarkdownPdfOptions -Title 'Service report' -Author 'Evotec' -IncludeLocalImages -BaseDirectory .\Assets
/// Convert-OfficeDocument -Path .\Report.md -OutputPath .\Report.pdf -Options $options</code>
///   <para>Builds a typed options object through ordinary PowerShell parameters; no hashtable or .NET construction is required.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeMarkdownPdfOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(MarkdownPdfExportOptions))]
public sealed class NewOfficeMarkdownPdfOptionsCommand : PSCmdlet {
    /// <summary>Existing Markdown PDF options to clone and override.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Default")]
    public MarkdownPdfExportOptions? BaseOptions { get; set; }

    /// <summary>Underlying low-level OfficeIMO PDF options.</summary>
    [Parameter]
    public OfficeIMO.Pdf.PdfCreateOptions? PdfOptions { get; set; }

    /// <summary>Built-in theme name, shared theme or theme package.</summary>
    [Parameter, OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Default font family.</summary>
    [Parameter]
    public string? Font { get; set; }

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

    /// <summary>Skip the native default Markdown PDF theme.</summary>
    [Parameter]
    public SwitchParameter NoDefaultTheme { get; set; }

    /// <summary>Embed supported local image files.</summary>
    [Parameter]
    public SwitchParameter IncludeLocalImages { get; set; }

    /// <summary>Embed supported data URI images.</summary>
    [Parameter]
    public SwitchParameter IncludeDataUriImages { get; set; }

    /// <summary>Require local images to resolve under BaseDirectory.</summary>
    [Parameter]
    public SwitchParameter NoLocalImageRestriction { get; set; }

    /// <summary>Maximum decoded bytes for one data URI image.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaximumDataUriImageBytes { get; set; }

    /// <summary>Fallback image width in PDF points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? DefaultImageWidth { get; set; }

    /// <summary>Fallback image height in PDF points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? DefaultImageHeight { get; set; }

    /// <summary>Controls how YAML front matter appears in the PDF body.</summary>
    [Parameter]
    public MarkdownPdfFrontMatterRenderMode? FrontMatterRenderMode { get; set; }

    /// <summary>Use front matter values to select a visual theme.</summary>
    [Parameter]
    public SwitchParameter NoFrontMatterTheme { get; set; }

    /// <summary>Use front matter values as PDF metadata.</summary>
    [Parameter]
    public SwitchParameter NoFrontMatterMetadata { get; set; }

    /// <summary>Use the first Markdown heading as the PDF title when no title is supplied.</summary>
    [Parameter]
    public SwitchParameter NoFirstHeadingTitle { get; set; }

    /// <summary>Create PDF outlines from Markdown headings.</summary>
    [Parameter]
    public SwitchParameter NoOutline { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = BaseOptions?.Clone() ?? new MarkdownPdfExportOptions();
        if (PdfOptions != null) options.PdfOptions = PdfOptions.Clone();
        if (Theme != null) options.Theme = MarkdownVisualTheme.FromOfficeTheme(Theme.Resolve(OfficeTheme.FromDocument));
        if (Font != null) options.FontFamily = Font;
        if (Title != null) options.Title = Title;
        if (Author != null) options.Author = Author;
        if (Subject != null) options.Subject = Subject;
        if (Keywords != null) options.Keywords = Keywords;
        if (BaseDirectory != null) options.BaseDirectory = SessionState.Path.GetUnresolvedProviderPathFromPSPath(BaseDirectory);
        if (GetBoundSwitch(nameof(NoDefaultTheme), NoDefaultTheme) is bool noTheme) options.ApplyDefaultTheme = !noTheme;
        if (GetBoundSwitch(nameof(IncludeLocalImages), IncludeLocalImages) is bool local) options.ResourcePolicy.AllowLocalFileAccess = local;
        if (GetBoundSwitch(nameof(IncludeDataUriImages), IncludeDataUriImages) is bool data) options.ResourcePolicy.AllowDataUris = data;
        if (GetBoundSwitch(nameof(NoLocalImageRestriction), NoLocalImageRestriction) is bool noRestriction) options.RestrictLocalImagesToBaseDirectory = !noRestriction;
        if (MaximumDataUriImageBytes.HasValue) options.MaximumDataUriImageBytes = MaximumDataUriImageBytes.Value;
        if (DefaultImageWidth.HasValue) options.DefaultImageWidth = DefaultImageWidth.Value;
        if (DefaultImageHeight.HasValue) options.DefaultImageHeight = DefaultImageHeight.Value;
        if (FrontMatterRenderMode.HasValue) options.FrontMatterRenderMode = FrontMatterRenderMode.Value;
        if (GetBoundSwitch(nameof(NoFrontMatterTheme), NoFrontMatterTheme) is bool noFrontTheme) options.UseFrontMatterTheme = !noFrontTheme;
        if (GetBoundSwitch(nameof(NoFrontMatterMetadata), NoFrontMatterMetadata) is bool noMetadata) options.UseFrontMatterMetadata = !noMetadata;
        if (GetBoundSwitch(nameof(NoFirstHeadingTitle), NoFirstHeadingTitle) is bool noHeading) options.UseFirstHeadingAsTitle = !noHeading;
        if (GetBoundSwitch(nameof(NoOutline), NoOutline) is bool noOutline) options.CreateOutlineFromHeadings = !noOutline;
        WriteObject(options);
    }

    private bool? GetBoundSwitch(string name, SwitchParameter value) =>
        MyInvocation.BoundParameters.ContainsKey(name) ? value.IsPresent : null;
}
