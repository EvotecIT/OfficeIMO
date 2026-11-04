using System.Management.Automation;
using OfficeIMO.Rtf.Markdown;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Creates native MarkdownRtfExportOptions settings for document operations.</summary>
/// <example><summary>Create default settings.</summary><code>$options = New-OfficeMarkdownRtfOptions</code></example>
[Cmdlet(VerbsCommon.New, "OfficeMarkdownRtfOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeIMO.Rtf.Markdown.MarkdownRtfExportOptions))]
public sealed class NewOfficeMarkdownRtfOptionsCommand : OfficeCmdlet {
    /// <summary>PreserveRawHtmlAsText override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter PreserveRawHtmlAsText { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeIMO.Rtf.Markdown.MarkdownRtfExportOptions();
        if (BoundSwitch(nameof(PreserveRawHtmlAsText)) is bool valuePreserveRawHtmlAsText) options.PreserveRawHtmlAsText = valuePreserveRawHtmlAsText;
        WriteObject(options);
    }
}
