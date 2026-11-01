using System;
using System.Management.Automation;
using OfficeIMO.Word.Html;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates native HTML-to-Word options for Convert-OfficeDocument.</summary>
/// <example><summary>Apply an explicit stylesheet while keeping document-linked stylesheets disabled.</summary>
/// <code>$options = New-OfficeHtmlWordOptions -StylesheetPath ./report.css -RenderPreAsTable
/// Convert-OfficeDocument -Path ./Report.html -To Word -OutputPath ./Report.docx -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeHtmlWordOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordHtmlImportOptions))]
public sealed class NewOfficeHtmlWordOptionsCommand : OfficeCmdlet {
    /// <summary>Font family for imported text.</summary>
    [Parameter(ParameterSetName = "Default"), ValidateNotNullOrEmpty]
    public string? Font { get; set; }
    /// <summary>Filesystem base folder for relative resources, without creating it.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? BasePath { get; set; }
    /// <summary>Explicit literal stylesheet files, resolved through the FileSystem provider.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string[]? StylesheetPath { get; set; }
    /// <summary>Additional CSS stylesheet text.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string[]? StylesheetContent { get; set; }
    /// <summary>Import list style information.</summary>
    [Parameter]
    public SwitchParameter IncludeListStyles { get; set; }
    /// <summary>Continue numbering across imported lists.</summary>
    [Parameter]
    public SwitchParameter ContinueNumbering { get; set; }
    /// <summary>Allow heading numbering during import.</summary>
    [Parameter]
    public SwitchParameter SupportsHeadingNumbering { get; set; }
    /// <summary>Render preformatted blocks inside a single-cell table.</summary>
    [Parameter]
    public SwitchParameter RenderPreAsTable { get; set; }
    /// <summary>Position of an imported table caption.</summary>
    [Parameter]
    public WordHtmlTableCaptionPosition TableCaptionPosition { get; set; }
    /// <summary>How HTML section tags map into Word.</summary>
    [Parameter]
    public WordHtmlSectionTagHandling SectionTagHandling { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordHtmlImportOptions();
        if (Font != null) options.FontFamily = Font;
        if (BasePath != null) options.BasePath = ResolveOutputDirectory(BasePath);
        foreach (string path in StylesheetPath ?? Array.Empty<string>())
            options.StylesheetPaths.AddRange(OfficePath.ResolveInput(SessionState, path, literal: true));
        if (StylesheetContent != null) options.StylesheetContents.AddRange(StylesheetContent);
        if (BoundSwitch(nameof(IncludeListStyles)) is bool lists) options.IncludeListStyles = lists;
        if (BoundSwitch(nameof(ContinueNumbering)) is bool numbering) options.ContinueNumbering = numbering;
        if (BoundSwitch(nameof(SupportsHeadingNumbering)) is bool headings) options.SupportsHeadingNumbering = headings;
        if (BoundSwitch(nameof(RenderPreAsTable)) is bool pre) options.RenderPreAsTable = pre;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TableCaptionPosition))) options.TableCaptionPosition = TableCaptionPosition;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(SectionTagHandling))) options.SectionTagHandling = SectionTagHandling;
        WriteObject(options);
    }
}
