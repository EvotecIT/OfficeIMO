using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a native text box with deliberate placement and plain or rich content.</summary>
/// <example><summary>Place rich text at physical coordinates.</summary>
/// <code>PowerPointSlide -Title 'Service review' {
///     PowerPointTextBox -X 2cm -Y 4cm -Width 12cm -Height 3cm {
///         PowerPointParagraph { PowerPointText 'Status: ' -Bold; PowerPointText 'Healthy' }
///     }
/// }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointTextBox", DefaultParameterSetName = "Text")]
[Alias("PowerPointTextBox")]
[OutputType(typeof(PowerPointTextBox))]
public sealed class AddOfficePowerPointTextBoxCommand : OfficeCmdlet {
    /// <summary>Gets or sets a native slide or one-based position; omission uses the current slide block.</summary>
    [Parameter(ValueFromPipeline = true), PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }
    /// <summary>Gets or sets plain text; an empty string creates an empty box.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Gets or sets native paragraph and run content.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content"), ValidateNotNull]
    public ScriptBlock? Content { get; set; }
    /// <summary>Gets or sets the native shape name.</summary>
    [Parameter]
    public string? Name { get; set; }
    /// <summary>Gets or sets the left coordinate; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? X { get; set; }
    /// <summary>Gets or sets the top coordinate; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Y { get; set; }
    /// <summary>Gets or sets the width; omission uses the native default.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }
    /// <summary>Gets or sets the height; omission uses the native default.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }
    /// <summary>Gets or sets the one-based theme grid column; explicit X overrides its position.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Column { get; set; }
    /// <summary>Gets or sets the number of grid columns; the default is twelve, and explicit Width overrides it.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Span { get; set; }
    /// <summary>Gets or sets whether the live native text box is returned.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointTextBox? box = null;
        try {
            if (Width is OfficeLength w && w <= OfficeLength.Zero || Height is OfficeLength h && h <= OfficeLength.Zero)
                throw new ArgumentException("Text-box width and height must be positive.");
            var target = PowerPointManualPlacement.Resolve(Slide);
            var horizontal = PowerPointManualPlacement.Horizontal(target.Slide, X, Width, Column, Span);
            box = target.Slide.AddTextBox(Content == null ? Text : string.Empty, horizontal.X, Y, horizontal.Width, Height);
            if (Name != null) box.Name = Name;
            if (Content != null) {
                var document = target.Slide.OwnerDocument ?? throw new ArgumentException("The slide has no owning document.");
                box.Paragraphs[0].ClearInlineContent();
                using (OfficeDocumentContext.Push(document))
                using (OfficeDocumentContext.Push(new PowerPointTextScope(document, box))) Content.Invoke();
            }
            target.Scope?.TrackFixedFigure(box);
            if (PassThru) WriteObject(box);
        } catch (Exception error) {
            box?.Remove();
            if (error is PipelineStoppedException or ActionPreferenceStopException) throw;
            if (PowerPointContentScope.Current != null)
                ThrowTerminatingError(OfficeError.Create(error, "powerpoint-textbox-create-failed", Slide, ErrorCategory.InvalidArgument));
            if (error is ArgumentException argument) InvalidArgument(argument, Slide);
            else ItemError(error, "powerpoint-textbox-create-failed", Slide);
        }
    }
}
