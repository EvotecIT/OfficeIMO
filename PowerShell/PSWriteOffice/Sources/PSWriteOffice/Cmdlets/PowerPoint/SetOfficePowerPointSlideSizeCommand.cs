using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets a presentation's slide size using a preset or physical lengths.</summary>
/// <para>Numbers mean points; length strings accept pt, cm, mm, in, px or emu. PassThru returns the native slide size.</para>
/// <example><summary>Choose a portrait presentation preset.</summary>
/// <code>Set-OfficePowerPointSlideSize -Document $deck -Preset Screen4x3 -Portrait</code></example>
/// <example><summary>Set a custom size before authoring slides.</summary>
/// <code>New-OfficePowerPoint -Path ./Review.pptx {
///     Set-OfficePowerPointSlideSize -Width '25.4cm' -Height '14cm'
///     PowerPointSlide -Title 'Service review'
/// }</code></example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointSlideSize", DefaultParameterSetName = "Preset")]
[OutputType(typeof(PowerPointSlideSize))]
public sealed class SetOfficePowerPointSlideSizeCommand : OfficeDocumentCmdlet<PowerPointDocument> {
    /// <summary>Gets or sets a standard slide size.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Preset")]
    public PowerPointSlideSizePreset Preset { get; set; }

    /// <summary>Gets or sets whether to rotate the chosen preset to portrait orientation.</summary>
    [Parameter(ParameterSetName = "Preset")]
    public SwitchParameter Portrait { get; set; }

    /// <summary>Gets or sets the custom slide width; numbers mean points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Size"), OfficeLengthTransformation]
    public OfficeLength Width { get; set; }

    /// <summary>Gets or sets the custom slide height; numbers mean points.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Size"), OfficeLengthTransformation]
    public OfficeLength Height { get; set; }

    /// <summary>Gets or sets whether to return the native slide size after editing.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointDocument document = GetDocument();
        try {
            var size = document.SlideSize;
            if (ParameterSetName == "Preset") size.SetPreset(Preset, Portrait.IsPresent);
            else size.SetSize(Width, Height);
            if (PassThru) WriteObject(size);
        } catch (ArgumentException error) { InvalidArgument(error, document); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-slide-size-edit-failed", document);
        }
    }
}
