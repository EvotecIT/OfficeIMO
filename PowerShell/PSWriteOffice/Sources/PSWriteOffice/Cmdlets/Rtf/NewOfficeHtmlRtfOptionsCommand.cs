using System.Management.Automation;
using OfficeIMO.Word.Html;
using PSWriteOffice.Cmdlets.Html;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Creates native HTML/CSS import settings for the HTML-to-Word-to-RTF route.</summary>
[Cmdlet(VerbsCommon.New, "OfficeHtmlRtfOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(HtmlWordExportOptions))]
public sealed class NewOfficeHtmlRtfOptionsCommand : PSCmdlet {
    /// <summary>Default font family for imported HTML.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Font { get; set; }
    /// <summary>Base directory or URI used to resolve resources.</summary>
    [Parameter]
    public string? BasePath { get; set; }
    /// <summary>Literal stylesheet files applied to the imported HTML.</summary>
    [Parameter]
    public string[]? StylesheetPath { get; set; }
    /// <summary>Additional stylesheet source text.</summary>
    [Parameter]
    public string[]? StylesheetContent { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new HtmlWordExportOptions();
        if (Font != null) options.FontFamily = Font;
        if (BasePath != null) {
            var uri = HtmlOptionsCommandUtilities.NormalizeBaseUri(SessionState, BasePath);
            options.BasePath = uri.IsFile ? uri.LocalPath : uri.AbsoluteUri;
        }
        foreach (var path in StylesheetPath ?? System.Array.Empty<string>())
            options.StylesheetPaths.Add(OfficePath.ResolveInput(SessionState, path, literal: true)[0]);
        if (StylesheetContent != null) options.StylesheetContents.AddRange(StylesheetContent);
        WriteObject(options);
    }
}
