using System.Management.Automation;
using OfficeIMO.Markdown;
using OfficeIMO.PowerPoint;
using OfficeIMO.Workflows;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Creates native Markdown slide boundaries, speaker-note, front-matter and image policies.</summary>
/// <example><summary>Start slides at level-two headings.</summary><code>$options = New-OfficeMarkdownPowerPointOptions -SlideBreak Heading -TitleHeadingLevel 2
/// Convert-OfficeDocument -Path ./Kickoff.md -To PowerPoint -OutputPath ./Kickoff.pptx -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeMarkdownPowerPointOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(MarkdownToPowerPointOptions))]
public sealed class NewOfficeMarkdownPowerPointOptionsCommand : OfficeCmdlet {
    /// <summary>Native boundaries that begin a slide.</summary>
    [Parameter(ParameterSetName = "Default")]
    public MarkdownSlideBreak? SlideBreak { get; set; }
    /// <summary>Heading level used for slide titles.</summary>
    [Parameter, ValidateRange(1, 6)]
    public int? TitleHeadingLevel { get; set; }
    /// <summary>Keep blockquotes visible instead of placing them in speaker notes.</summary>
    [Parameter]
    public SwitchParameter NoBlockquoteNotes { get; set; }
    /// <summary>Ignore title, subtitle and author front matter.</summary>
    [Parameter]
    public SwitchParameter NoFrontMatter { get; set; }
    /// <summary>Native presentation canvas preset.</summary>
    [Parameter]
    public PowerPointSlideSizePreset? SlideSize { get; set; }
    /// <summary>Base directory for relative images.</summary>
    [Parameter]
    public string? BaseDirectory { get; set; }
    /// <summary>Allow native local image loading within the configured roots.</summary>
    [Parameter]
    public SwitchParameter AllowLocalImages { get; set; }
    /// <summary>Allowed local image directories.</summary>
    [Parameter]
    public string[]? AllowedImageDirectory { get; set; }
    /// <summary>Allow the native bounded HTTP image loader.</summary>
    [Parameter]
    public SwitchParameter AllowRemoteImages { get; set; }
    /// <summary>Disable data URI images.</summary>
    [Parameter]
    public SwitchParameter NoDataUriImages { get; set; }
    /// <summary>Maximum decoded data URI image bytes.</summary>
    [Parameter, ValidateRange(0, long.MaxValue)]
    public long? MaximumDataUriImageBytes { get; set; }
    /// <summary>Native Markdown dialect profile; mutually exclusive with ReaderOptions.</summary>
    [Parameter]
    public MarkdownLoadOptions.MarkdownDialectProfile? Profile { get; set; }
    /// <summary>Native parser, normalization and input-limit policy, captured independently.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(MarkdownLoadOptions))]
    public MarkdownLoadOptions? ReaderOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ReaderOptions != null && Profile.HasValue) throw new PSArgumentException("Specify either ReaderOptions or Profile.");
        var options = new MarkdownToPowerPointOptions {
            ReaderOptions = ReaderOptions?.Clone() ?? (Profile.HasValue ? MarkdownLoadOptions.CreateProfile(Profile.Value) : null)
        };
        if (SlideBreak.HasValue) options.SlideBreak = SlideBreak.Value;
        if (TitleHeadingLevel.HasValue) options.TitleHeadingLevel = TitleHeadingLevel.Value;
        if (BoundSwitch(nameof(NoBlockquoteNotes)) is bool noNotes) options.BlockquotesAsNotes = !noNotes;
        if (BoundSwitch(nameof(NoFrontMatter)) is bool noFrontMatter) options.RenderFrontMatter = !noFrontMatter;
        if (SlideSize.HasValue) options.SlideSize = SlideSize.Value;
        if (BaseDirectory != null) options.BaseDirectory = SessionState.Path.GetUnresolvedProviderPathFromPSPath(BaseDirectory);
        if (BoundSwitch(nameof(AllowLocalImages)) is bool local) options.AllowLocalImages = local;
        foreach (string directory in AllowedImageDirectory ?? Array.Empty<string>())
            options.AllowedImageDirectories.Add(SessionState.Path.GetUnresolvedProviderPathFromPSPath(directory));
        if (BoundSwitch(nameof(AllowRemoteImages)) is bool remote) options.AllowRemoteImages = remote;
        if (BoundSwitch(nameof(NoDataUriImages)) is bool noData) options.AllowDataUriImages = !noData;
        if (MaximumDataUriImageBytes.HasValue) options.MaximumDataUriImageBytes = MaximumDataUriImageBytes.Value;
        WriteObject(options);
    }
}
