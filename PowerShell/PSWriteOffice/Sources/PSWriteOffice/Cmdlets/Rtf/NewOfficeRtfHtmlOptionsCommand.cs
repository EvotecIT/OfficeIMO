using System.Management.Automation;
using OfficeIMO.Word.Html;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Creates native WordHtmlExportOptions settings for document operations.</summary>
/// <example><summary>Create default settings.</summary><code>$options = New-OfficeRtfHtmlOptions</code></example>
[Cmdlet(VerbsCommon.New, "OfficeRtfHtmlOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeIMO.Word.Html.WordHtmlExportOptions))]
public sealed class NewOfficeRtfHtmlOptionsCommand : OfficeCmdlet {
    /// <summary>Font override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public string? Font { get; set; }
    /// <summary>IncludeFontStyles override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter IncludeFontStyles { get; set; }
    /// <summary>IncludeListStyles override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter IncludeListStyles { get; set; }
    /// <summary>IncludeParagraphClasses override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter IncludeParagraphClasses { get; set; }
    /// <summary>IncludeRunClasses override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter IncludeRunClasses { get; set; }
    /// <summary>IncludeDefaultCss override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter IncludeDefaultCss { get; set; }
    /// <summary>UseImagePaths override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter UseImagePaths { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeIMO.Word.Html.WordHtmlExportOptions();
        if (Font != null) options.FontFamily = Font;
        if (BoundSwitch(nameof(IncludeFontStyles)) is bool valueIncludeFontStyles) options.IncludeFontStyles = valueIncludeFontStyles;
        if (BoundSwitch(nameof(IncludeListStyles)) is bool valueIncludeListStyles) options.IncludeListStyles = valueIncludeListStyles;
        if (BoundSwitch(nameof(IncludeParagraphClasses)) is bool valueIncludeParagraphClasses) options.IncludeParagraphClasses = valueIncludeParagraphClasses;
        if (BoundSwitch(nameof(IncludeRunClasses)) is bool valueIncludeRunClasses) options.IncludeRunClasses = valueIncludeRunClasses;
        if (BoundSwitch(nameof(IncludeDefaultCss)) is bool valueIncludeDefaultCss) options.IncludeDefaultCss = valueIncludeDefaultCss;
        if (BoundSwitch(nameof(UseImagePaths)) is bool valueUseImagePaths) options.EmbedImagesAsBase64 = !valueUseImagePaths;
        WriteObject(options);
    }
}
