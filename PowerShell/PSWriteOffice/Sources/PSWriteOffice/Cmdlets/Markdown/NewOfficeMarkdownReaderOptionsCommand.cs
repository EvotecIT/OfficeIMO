using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Creates native Markdown dialect, normalization, URL and input-limit settings.</summary>
/// <example><summary>Bound input from uploaded files.</summary>
/// <code>$options = New-OfficeMarkdownReaderOptions -MaxInputCharacters 500000 -DisallowFileUrls
/// Open-OfficeMarkdown -Path ./Uploads/*.md -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeMarkdownReaderOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(MarkdownLoadOptions))]
public sealed class NewOfficeMarkdownReaderOptionsCommand : OfficeCmdlet, IMarkdownReaderOptionSource {
    /// <summary>Native Markdown dialect profile.</summary>
    [Parameter]
    public MarkdownLoadOptions.MarkdownDialectProfile? Profile { get; set; }
    /// <summary>Base URI for relative links and images.</summary>
    [Parameter]
    public string? BaseUri { get; set; }
    /// <summary>Maximum input characters accepted by the native reader.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxInputCharacters { get; set; }
    /// <summary>Input normalization preset.</summary>
    [Parameter]
    public MarkdownInputNormalizationPreset? NormalizeInput { get; set; }
    /// <summary>Block file URLs.</summary>
    [Parameter]
    public SwitchParameter DisallowFileUrls { get; set; }
    /// <summary>Allow data URLs.</summary>
    [Parameter]
    public SwitchParameter AllowDataUrls { get; set; }
    /// <summary>Allow mailto links.</summary>
    [Parameter]
    public SwitchParameter AllowMailtoUrls { get; set; }
    /// <summary>Allow protocol-relative links.</summary>
    [Parameter]
    public SwitchParameter AllowProtocolRelativeUrls { get; set; }
    /// <summary>Restrict links to the native allow-list.</summary>
    [Parameter]
    public SwitchParameter RestrictUrlSchemes { get; set; }
    /// <summary>Allowed URL schemes; supplying this enables the allow-list.</summary>
    [Parameter]
    public string[]? AllowedUrlScheme { get; set; }
    MarkdownLoadOptions? IMarkdownReaderOptionSource.ReaderOptions => null;
    bool? IMarkdownReaderOptionSource.DisallowFileUrls => BoundSwitch(nameof(DisallowFileUrls));
    bool? IMarkdownReaderOptionSource.AllowDataUrls => BoundSwitch(nameof(AllowDataUrls));
    bool? IMarkdownReaderOptionSource.AllowMailtoUrls => BoundSwitch(nameof(AllowMailtoUrls));
    bool? IMarkdownReaderOptionSource.AllowProtocolRelativeUrls => BoundSwitch(nameof(AllowProtocolRelativeUrls));
    bool? IMarkdownReaderOptionSource.RestrictUrlSchemes => BoundSwitch(nameof(RestrictUrlSchemes));
    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(MarkdownOptionUtilities.BuildReaderOptions(this) ?? MarkdownLoadOptions.CreateOfficeIMOProfile());
}
