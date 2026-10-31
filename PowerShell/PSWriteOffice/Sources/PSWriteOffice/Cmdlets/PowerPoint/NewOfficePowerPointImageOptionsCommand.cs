using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.Imaging;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates discoverable slide selection and rendering settings for Export-OfficePowerPointImage.</summary>
/// <example>
///   <summary>Render selected slides with their backgrounds and content.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficePowerPointImageOptions -SlideNumber 1,3 -NoPictures
/// Export-OfficePowerPointImage -Path .\Deck.pptx -OutputPath .\Slides -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePowerPointImageOptions")]
[OutputType(typeof(PowerPointImageExportOptions))]
public sealed class NewOfficePowerPointImageOptionsCommand : OfficeImageOptionsCommandBase<PowerPointImageExportOptions> {
    /// <summary>One-based slide numbers to export.</summary>
    [Parameter] public int[]? SlideNumber { get; set; }
    /// <summary>Include hidden slides.</summary>
    [Parameter] public SwitchParameter IncludeHiddenSlides { get; set; }
    /// <summary>Omit slide backgrounds.</summary>
    [Parameter] public SwitchParameter NoSlideBackground { get; set; }
    /// <summary>Omit slide content.</summary>
    [Parameter] public SwitchParameter NoSlideContent { get; set; }
    /// <summary>Omit pictures.</summary>
    [Parameter] public SwitchParameter NoPictures { get; set; }
    /// <summary>Omit auto shapes.</summary>
    [Parameter] public SwitchParameter NoAutoShapes { get; set; }
    /// <summary>Omit text boxes.</summary>
    [Parameter] public SwitchParameter NoTextBoxes { get; set; }
    /// <summary>Omit tables.</summary>
    [Parameter] public SwitchParameter NoTables { get; set; }
    /// <summary>Omit charts.</summary>
    [Parameter] public SwitchParameter NoCharts { get; set; }
    /// <summary>Render hidden shapes.</summary>
    [Parameter] public SwitchParameter IncludeHiddenShapes { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PowerPointImageExportOptions();
        ApplyCommon(options);
        if (SlideNumber != null) options.SlideNumbers = SlideNumber;
        if (BoundSwitch(nameof(IncludeHiddenSlides)) is bool hiddenSlides) options.IncludeHiddenSlides = hiddenSlides;
        options.IncludeSlideBackground = FeatureEnabled("SlideBackground");
        options.IncludeSlideContent = FeatureEnabled("SlideContent");
        options.IncludePictures = FeatureEnabled("Pictures");
        options.IncludeAutoShapes = FeatureEnabled("AutoShapes");
        options.IncludeTextBoxes = FeatureEnabled("TextBoxes");
        options.IncludeTables = FeatureEnabled("Tables");
        options.IncludeCharts = FeatureEnabled("Charts");
        if (BoundSwitch(nameof(IncludeHiddenShapes)) is bool hiddenShapes) options.IncludeHiddenShapes = hiddenShapes;
        WriteObject(options);
    }
}
