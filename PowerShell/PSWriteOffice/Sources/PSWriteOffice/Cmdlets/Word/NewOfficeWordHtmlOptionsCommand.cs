using System.Management.Automation;
using OfficeIMO.Word.Html;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates native Word-to-HTML options for Convert-OfficeDocument.</summary>
/// <example><summary>Export reusable paragraph and run style classes.</summary>
/// <code>$options = New-OfficeWordHtmlOptions -IncludeParagraphClasses -IncludeRunClasses
/// Convert-OfficeDocument -Path ./Guide.docx -To Html -OutputPath ./Guide.html -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeWordHtmlOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordHtmlExportOptions))]
public sealed class NewOfficeWordHtmlOptionsCommand : OfficeCmdlet {
    /// <summary>Font family applied to exported text.</summary>
    [Parameter(ParameterSetName = "Default"), ValidateNotNullOrEmpty]
    public string? Font { get; set; }

    /// <summary>Include run font information as inline styles.</summary>
    [Parameter]
    public SwitchParameter IncludeFontStyles { get; set; }

    /// <summary>Include list style information.</summary>
    [Parameter]
    public SwitchParameter IncludeListStyles { get; set; }

    /// <summary>Emit paragraph styles as CSS classes.</summary>
    [Parameter]
    public SwitchParameter IncludeParagraphClasses { get; set; }

    /// <summary>Emit run character styles as CSS classes.</summary>
    [Parameter]
    public SwitchParameter IncludeRunClasses { get; set; }

    /// <summary>Include the native default stylesheet. Omission preserves the engine default.</summary>
    [Parameter]
    public SwitchParameter IncludeDefaultCss { get; set; }

    /// <summary>Reference image file paths instead of embedding image data.</summary>
    [Parameter]
    public SwitchParameter UseImagePaths { get; set; }

    /// <summary>Omit footnotes from the exported HTML.</summary>
    [Parameter]
    public SwitchParameter ExcludeFootnotes { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordHtmlExportOptions();
        if (Font != null) options.FontFamily = Font;
        if (BoundSwitch(nameof(IncludeFontStyles)) is bool fonts) options.IncludeFontStyles = fonts;
        if (BoundSwitch(nameof(IncludeListStyles)) is bool lists) options.IncludeListStyles = lists;
        if (BoundSwitch(nameof(IncludeParagraphClasses)) is bool paragraphs) options.IncludeParagraphClasses = paragraphs;
        if (BoundSwitch(nameof(IncludeRunClasses)) is bool runs) options.IncludeRunClasses = runs;
        if (BoundSwitch(nameof(IncludeDefaultCss)) is bool css) options.IncludeDefaultCss = css;
        if (BoundSwitch(nameof(UseImagePaths)) is bool paths) options.EmbedImagesAsBase64 = !paths;
        if (BoundSwitch(nameof(ExcludeFootnotes)) is bool footnotes) options.ExportFootnotes = !footnotes;
        WriteObject(options);
    }
}
