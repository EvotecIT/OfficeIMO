using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.Imaging;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates discoverable slide selection and rendering settings for Export-OfficePowerPointImage.</summary>
/// <example>
///   <summary>Render selected slides with their backgrounds and content.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficePowerPointImageOptions -SlideNumber 1,3 -IncludeSlideBackground -IncludeSlideContent
/// Export-OfficePowerPointImage -Path .\Deck.pptx -OutputPath .\Slides -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePowerPointImageOptions")]
[OutputType(typeof(PowerPointPresentationImageExportOptions))]
public sealed class NewOfficePowerPointImageOptionsCommand : OfficeImageOptionsCommandBase<PowerPointPresentationImageExportOptions> {
    /// <summary>One-based slide numbers to export.</summary>
    [Parameter] public int[]? SlideNumber { get; set; }
    /// <summary>Include hidden slides.</summary>
    [Parameter] public SwitchParameter IncludeHiddenSlides { get; set; }
    /// <summary>Render slide backgrounds.</summary>
    [Parameter] public SwitchParameter IncludeSlideBackground { get; set; }
    /// <summary>Render slide content.</summary>
    [Parameter] public SwitchParameter IncludeSlideContent { get; set; }
    /// <summary>Render pictures.</summary>
    [Parameter] public SwitchParameter IncludePictures { get; set; }
    /// <summary>Render auto shapes.</summary>
    [Parameter] public SwitchParameter IncludeAutoShapes { get; set; }
    /// <summary>Render text boxes.</summary>
    [Parameter] public SwitchParameter IncludeTextBoxes { get; set; }
    /// <summary>Render tables.</summary>
    [Parameter] public SwitchParameter IncludeTables { get; set; }
    /// <summary>Render charts.</summary>
    [Parameter] public SwitchParameter IncludeCharts { get; set; }
    /// <summary>Render hidden shapes.</summary>
    [Parameter] public SwitchParameter IncludeHiddenShapes { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PowerPointPresentationImageExportOptions();
        ApplyCommon(options);
        if (SlideNumber != null) options.SlideNumbers = SlideNumber;
        Apply(nameof(IncludeHiddenSlides), value => options.IncludeHiddenSlides = value);
        Apply(nameof(IncludeSlideBackground), value => options.IncludeSlideBackground = value);
        Apply(nameof(IncludeSlideContent), value => options.IncludeSlideContent = value);
        Apply(nameof(IncludePictures), value => options.IncludePictures = value);
        Apply(nameof(IncludeAutoShapes), value => options.IncludeAutoShapes = value);
        Apply(nameof(IncludeTextBoxes), value => options.IncludeTextBoxes = value);
        Apply(nameof(IncludeTables), value => options.IncludeTables = value);
        Apply(nameof(IncludeCharts), value => options.IncludeCharts = value);
        Apply(nameof(IncludeHiddenShapes), value => options.IncludeHiddenShapes = value);
        WriteObject(options);
    }
    private void Apply(string name, System.Action<bool> setter) { if (IsBound(name)) setter(((SwitchParameter)MyInvocation.BoundParameters[name]).IsPresent); }
}
