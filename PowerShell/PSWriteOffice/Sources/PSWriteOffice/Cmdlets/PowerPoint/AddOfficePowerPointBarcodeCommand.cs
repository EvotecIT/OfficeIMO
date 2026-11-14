using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a QR code or barcode to a slide.</summary>
/// <para>Inside PowerPointSlide without coordinates or a size, the code flows through the slide layout and keeps its proportions. X, Y, Width, Height or Slide place it at a fixed position instead. The code is an SVG picture with a PNG fallback; PDF conversion keeps it as vector content. The alternative text defaults to the symbology and the encoded text.</para>
/// <example><summary>Close a talk with a link to the slides.</summary>
/// <code>PowerPointSlide -Title 'Thank you' {
///     PowerPointBarcode 'https://evotec.xyz/talks/officeimo'
/// }</code></example>
/// <example><summary>Pin a small QR code to a corner of an existing slide.</summary>
/// <code>PowerPointBarcode 'https://evotec.xyz' -Slide 3 -X 22cm -Y 15cm -Width 2.5cm -PassThru</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointBarcode")]
[OutputType(typeof(PowerPointImage))]
[Alias("PowerPointBarcode")]
public sealed class AddOfficePowerPointBarcodeCommand : OfficeBarcodeCommandBase {
    /// <summary>Native slide or a one-based position in the current document; omitted inside PowerPointSlide.</summary>
    [Parameter(ValueFromPipeline = true), PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }

    /// <summary>Left position; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? X { get; set; }

    /// <summary>Top position; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Y { get; set; }

    /// <summary>Return the native editable picture.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var scope = PowerPointContentScope.Current;
        try {
            var target = PowerPointManualPlacement.Resolve(Slide);
            PowerPointImage picture = target.Slide.AddBarcode(Type, Text!, X ?? OfficeLength.Zero, Y ?? OfficeLength.Zero, Configure);
            bool fixedPlacement = Slide != null || PowerPointManualPlacement.HasExplicitBounds(MyInvocation.BoundParameters.Keys);
            if (target.Scope != null) {
                if (fixedPlacement) target.Scope.TrackFixedFigure(picture);
                else target.Scope.AddFigure(new PowerPointLayoutImage(picture) { Contain = true }, picture);
            }
            if (PassThru) WriteObject(picture);
        } catch (Exception error) {
            if (error is PipelineStoppedException or ActionPreferenceStopException) throw;
            var record = OfficeError.Create(error, "invalid-argument", Text, ErrorCategory.InvalidArgument);
            if (scope != null) ThrowTerminatingError(record);
            else WriteError(record);
        }
    }
}
