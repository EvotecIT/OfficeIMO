using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates native Markdown-to-Word settings, including an engine-owned template path.</summary>
/// <example><summary>Insert Markdown into a template.</summary><code>$options = New-OfficeMarkdownWordOptions -Template .\Templates\Runbook.dotx -BookmarkName Body
/// Convert-OfficeDocument -Path .\Runbook.md -To Word -OutputPath .\Runbook.docx -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeMarkdownWordOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(MarkdownWordExportOptions))]
public sealed class NewOfficeMarkdownWordOptionsCommand : OfficeCmdlet {
    /// <summary>Template override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Template { get; set; }
    /// <summary>BookmarkName override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? BookmarkName { get; set; }
    /// <summary>ContentControlTag override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? ContentControlTag { get; set; }
    /// <summary>ContentControlAlias override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? ContentControlAlias { get; set; }
    /// <summary>KeepPlaceholder override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter KeepPlaceholder { get; set; }
    /// <summary>NoFrontMatter override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoFrontMatter { get; set; }
    /// <summary>Font override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Font { get; set; }
    /// <summary>BaseUri override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? BaseUri { get; set; }
    /// <summary>AllowLocalImages override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowLocalImages { get; set; }
    /// <summary>AllowedImageDirectory override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string[]? AllowedImageDirectory { get; set; }
    /// <summary>Built-in theme, shared theme object or theme template.</summary>
    [Parameter(ParameterSetName = "Default"), OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }
    /// <summary>Allow the native bounded HTTP image loader. Omission keeps remote images disabled.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowRemoteImages { get; set; }
    /// <summary>ReaderOptions override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeOptionsTransformation(typeof(MarkdownLoadOptions))]
    public MarkdownLoadOptions? ReaderOptions { get; set; }
    /// <summary>Profile override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public MarkdownLoadOptions.MarkdownDialectProfile? Profile { get; set; }
    /// <summary>NormalizeInput override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public MarkdownInputNormalizationPreset? NormalizeInput { get; set; }
    /// <summary>NoDataUriImages override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoDataUriImages { get; set; }
    /// <summary>MaxDataUriImageBytes override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public long? MaxDataUriImageBytes { get; set; }
    /// <summary>PreferNarrativeSingleLineDefinitions override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PreferNarrativeSingleLineDefinitions { get; set; }
    /// <summary>FitImagesToPageContentWidth override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter FitImagesToPageContentWidth { get; set; }
    /// <summary>FitImagesToContextWidth override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter FitImagesToContextWidth { get; set; }
    /// <summary>MaxImageWidth override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MaxImageWidth { get; set; }
    /// <summary>MaxImageHeight override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MaxImageHeight { get; set; }
    /// <summary>MaxImageWidthPercentOfContent override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public double? MaxImageWidthPercentOfContent { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ReaderOptions != null && Profile.HasValue) throw new PSArgumentException("Specify either -ReaderOptions or -Profile.");
        if (string.IsNullOrWhiteSpace(Template) && (BookmarkName != null || ContentControlTag != null || ContentControlAlias != null))
            throw new PSArgumentException("Template insertion parameters require -Template.");
        var options = new MarkdownWordExportOptions();
        if (Theme != null) options.Theme = MarkdownVisualTheme.FromOfficeTheme(Theme.Resolve(OfficeTheme.FromDocument));
        if (BoundSwitch(nameof(AllowRemoteImages)) is bool remote) options.AllowRemoteImages = remote;
        if (Template != null) options.TemplatePath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Template);
        if (BookmarkName != null) options.BookmarkName = BookmarkName;
        if (ContentControlTag != null) options.ContentControlTag = ContentControlTag;
        if (ContentControlAlias != null) options.ContentControlAlias = ContentControlAlias;
        if (BoundSwitch(nameof(KeepPlaceholder)) is bool keep) options.ReplacePlaceholder = !keep;
        if (BoundSwitch(nameof(NoFrontMatter)) is bool noFrontMatter) options.RenderFrontMatter = !noFrontMatter;
        if (Font != null) options.FontFamily = Font;
        if (BaseUri != null) options.BaseUri = BaseUri;
        if (BoundSwitch(nameof(AllowLocalImages)) is bool local) options.AllowLocalImages = local;
        foreach (var directory in AllowedImageDirectory ?? Array.Empty<string>()) options.AllowedImageDirectories.Add(SessionState.Path.GetUnresolvedProviderPathFromPSPath(directory));
        var reader = ReaderOptions?.Clone() ?? (Profile.HasValue ? MarkdownLoadOptions.CreateProfile(Profile.Value) : null);
        if (NormalizeInput.HasValue) {
            reader ??= MarkdownLoadOptions.CreateOfficeIMOProfile();
            reader.InputNormalization.ApplyPreset(NormalizeInput.Value);
        }
        options.ReaderOptions = reader;
        if (BoundSwitch(nameof(NoDataUriImages)) is bool noData) options.AllowDataUriImages = !noData;
        if (MaxDataUriImageBytes.HasValue) options.MaxDataUriImageBytes = MaxDataUriImageBytes.Value;
        if (BoundSwitch(nameof(PreferNarrativeSingleLineDefinitions)) is bool narrative) options.PreferNarrativeSingleLineDefinitions = narrative;
        if (BoundSwitch(nameof(FitImagesToPageContentWidth)) is bool fitPage) options.FitImagesToPageContentWidth = fitPage;
        if (BoundSwitch(nameof(FitImagesToContextWidth)) is bool fitContext) options.FitImagesToContextWidth = fitContext;
        if (MaxImageWidth.HasValue) options.MaxImageWidth = MaxImageWidth.Value;
        if (MaxImageHeight.HasValue) options.MaxImageHeight = MaxImageHeight.Value;
        if (MaxImageWidthPercentOfContent.HasValue) options.MaxImageWidthPercentOfContent = MaxImageWidthPercentOfContent.Value;
        WriteObject(options);
    }
}
