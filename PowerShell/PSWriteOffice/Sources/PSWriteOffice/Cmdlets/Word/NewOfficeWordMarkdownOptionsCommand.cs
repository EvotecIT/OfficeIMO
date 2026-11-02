using System.Management.Automation;
using OfficeIMO.Word.Markdown;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates native Word-to-Markdown options for Convert-OfficeDocument.</summary>
/// <example><summary>Retain underline and highlight markup in a runbook.</summary>
/// <code>$options = New-OfficeWordMarkdownOptions -EnableUnderline -EnableHighlight
/// Convert-OfficeDocument -Path ./Runbook.docx -To Markdown -OutputPath ./Runbook.md -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeWordMarkdownOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordMarkdownExportOptions))]
public sealed class NewOfficeWordMarkdownOptionsCommand : OfficeCmdlet {
    /// <summary>Font family whose runs become inline code; omission uses the native monospace resolver.</summary>
    [Parameter(ParameterSetName = "Default"), ValidateNotNullOrEmpty]
    public string? Font { get; set; }

    /// <summary>Wrap underlined text in HTML underline tags.</summary>
    [Parameter]
    public SwitchParameter EnableUnderline { get; set; }

    /// <summary>Wrap highlighted text in double equals signs.</summary>
    [Parameter]
    public SwitchParameter EnableHighlight { get; set; }

    /// <summary>Embed images as Base64 or save them as File references.</summary>
    [Parameter]
    public WordMarkdownImageExportMode ImageExportMode { get; set; }

    /// <summary>Filesystem folder for images when File mode is selected; this command does not create it.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? ImageDirectory { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordMarkdownExportOptions();
        if (Font != null) options.FontFamily = Font;
        if (BoundSwitch(nameof(EnableUnderline)) is bool underline) options.EnableUnderline = underline;
        if (BoundSwitch(nameof(EnableHighlight)) is bool highlight) options.EnableHighlight = highlight;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ImageExportMode))) options.ImageExportMode = ImageExportMode;
        if (ImageDirectory != null) options.ImageDirectory = ResolveOutputDirectory(ImageDirectory);
        WriteObject(options);
    }
}
