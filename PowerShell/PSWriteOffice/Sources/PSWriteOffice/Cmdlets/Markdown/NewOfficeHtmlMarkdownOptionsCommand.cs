using System;
using System.Management.Automation;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Creates native HTML-to-Markdown import and writer settings without converting or writing files.</summary>
/// <example><summary>Convert HTML using a portable Markdown profile.</summary>
/// <code>$options = New-OfficeHtmlMarkdownOptions -Portable -LineEnding LF
/// Convert-OfficeDocument -Path ./Report.html -To Markdown -OutputPath ./Report.md -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeHtmlMarkdownOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(MarkdownHtmlImportOptions))]
public sealed class NewOfficeHtmlMarkdownOptionsCommand : OfficeCmdlet, IMarkdownWriteOptionSource {
    /// <summary>Native settings to copy before applying explicit changes.</summary>
    [Parameter(ParameterSetName = "Default"), OfficeOptionsTransformation(typeof(MarkdownHtmlImportOptions)), ValidateNotNull]
    public MarkdownHtmlImportOptions? BaseOptions { get; set; }
    /// <summary>Start with the native portable profile instead of the OfficeIMO profile.</summary>
    [Parameter]
    public SwitchParameter Portable { get; set; }
    /// <summary>Absolute base URI for relative links and images.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? BaseUri { get; set; }
    /// <summary>Include the complete HTML document instead of only its body.</summary>
    [Parameter]
    public SwitchParameter IncludeDocumentChrome { get; set; }
    /// <summary>Retain script and style elements in the generated Markdown.</summary>
    [Parameter]
    public SwitchParameter PreserveScriptsAndStyles { get; set; }
    /// <summary>Convert unsupported block elements to their child content instead of retaining raw HTML.</summary>
    [Parameter]
    public SwitchParameter DropUnsupportedBlocks { get; set; }
    /// <summary>Convert unsupported inline elements instead of retaining raw HTML.</summary>
    [Parameter]
    public SwitchParameter DropUnsupportedInlineHtml { get; set; }
    /// <summary>Maximum accepted HTML input characters; null removes a copied limit.</summary>
    [Parameter]
    public int? MaxInputCharacters { get; set; }
    /// <summary>How data URI images are represented in Markdown.</summary>
    [Parameter]
    public MarkdownHtmlBase64ImageHandling? Base64ImageHandling { get; set; }
    /// <summary>Literal image extraction directory; creating options does not create it.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Base64ImageOutputDirectory { get; set; }
    /// <summary>Preserve or suppress repeated listing-card metadata.</summary>
    [Parameter]
    public MarkdownHtmlListingCardMetadataMode? ListingCardMetadataMode { get; set; }
    /// <summary>Maximum columns after expanding HTML table spans.</summary>
    [Parameter, ValidateRange(1, 4096)]
    public int? MaxTableExpandedColumns { get; set; }
    /// <summary>Native Markdown writer settings to copy.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(MarkdownSaveOptions)), ValidateNotNull]
    public MarkdownSaveOptions? WriteOptions { get; set; }
    /// <summary>Start with a named Markdown writer profile.</summary>
    [Parameter]
    public OfficeMarkdownWriteProfile? WriteProfile { get; set; }
    /// <summary>How the Markdown writer serializes images.</summary>
    [Parameter]
    public MarkdownImageRenderingMode? ImageRenderingMode { get; set; }
    /// <summary>Output line endings.</summary>
    [Parameter]
    public OfficeLineEnding? LineEnding { get; set; }
    /// <summary>Unordered list marker.</summary>
    [Parameter, ValidateSet("-", "*", "+")]
    public string? UnorderedListMarker { get; set; }
    string? IMarkdownWriteOptionSource.LineEnding => LineEnding?.ToString();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (BaseOptions != null && Portable) throw new PSArgumentException("Choose BaseOptions or Portable.");
        if (MaxInputCharacters is <= 0) throw new PSArgumentException("MaxInputCharacters must be positive or null.");
        MarkdownHtmlImportOptions options = BaseOptions?.Clone() ?? (Portable
            ? MarkdownHtmlImportOptions.CreatePortableProfile() : MarkdownHtmlImportOptions.CreateOfficeIMOProfile());
        if (BoundSwitch(nameof(IncludeDocumentChrome)) is bool chrome) options.UseBodyContentsOnly = !chrome;
        if (BoundSwitch(nameof(PreserveScriptsAndStyles)) is bool scripts) options.RemoveScriptsAndStyles = !scripts;
        if (BoundSwitch(nameof(DropUnsupportedBlocks)) is bool blocks) options.PreserveUnsupportedBlocks = !blocks;
        if (BoundSwitch(nameof(DropUnsupportedInlineHtml)) is bool inline) options.PreserveUnsupportedInlineHtml = !inline;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaxInputCharacters))) options.MaxInputCharacters = MaxInputCharacters;
        if (BaseUri != null) options.BaseUri = new Uri(BaseUri, UriKind.Absolute);
        if (Base64ImageHandling.HasValue) options.Base64Images = Base64ImageHandling.Value;
        if (Base64ImageOutputDirectory != null) options.Base64ImageOutputDirectory = ResolveOutputDirectory(Base64ImageOutputDirectory);
        if (ListingCardMetadataMode.HasValue) options.ListingCardMetadataMode = ListingCardMetadataMode.Value;
        if (MaxTableExpandedColumns.HasValue) options.MaxTableExpandedColumns = MaxTableExpandedColumns.Value;
        MarkdownSaveOptions? writer = MarkdownOptionUtilities.BuildWriteOptions(this, options.MarkdownWriteOptions);
        if (writer != null) options.MarkdownWriteOptions = writer;
        WriteObject(options);
    }
}
