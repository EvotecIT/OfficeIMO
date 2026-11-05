using System.Management.Automation;
using OfficeIMO.Rtf.Markdown;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Creates native RtfMarkdownExportOptions settings for document operations.</summary>
/// <example><summary>Create default settings.</summary><code>$options = New-OfficeRtfMarkdownOptions</code></example>
[Cmdlet(VerbsCommon.New, "OfficeRtfMarkdownOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeIMO.Rtf.Markdown.RtfMarkdownExportOptions))]
public sealed class NewOfficeRtfMarkdownOptionsCommand : OfficeCmdlet {
    /// <summary>IncludeHiddenText override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter IncludeHiddenText { get; set; }
    /// <summary>NoUnsupportedHtmlComments override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter NoUnsupportedHtmlComments { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeIMO.Rtf.Markdown.RtfMarkdownExportOptions();
        if (BoundSwitch(nameof(IncludeHiddenText)) is bool valueIncludeHiddenText) options.IncludeHiddenText = valueIncludeHiddenText;
        if (BoundSwitch(nameof(NoUnsupportedHtmlComments)) is bool valueNoUnsupportedHtmlComments) options.EmitUnsupportedHtmlComments = !valueNoUnsupportedHtmlComments;
        WriteObject(options);
    }
}
