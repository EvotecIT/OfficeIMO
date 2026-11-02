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
public sealed class AddOfficePowerPointTextBoxCommand : PowerPointTextPlacementCmdlet {
    /// <summary>Gets or sets plain text; an empty string creates an empty box.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Gets or sets native paragraph and run content.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content"), ValidateNotNull]
    public ScriptBlock? Content { get; set; }
    /// <summary>Gets or sets the native shape name.</summary>
    [Parameter]
    public string? Name { get; set; }
    /// <summary>Gets or sets whether the live native text box is returned.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointTextBox? box = null;
        try {
            var placed = CreatePlacedTextBox(Content == null ? Text : string.Empty);
            box = placed.Box;
            if (Name != null) box.Name = Name;
            if (Content != null) {
                var document = box.OwnerSlide?.OwnerDocument ?? throw new ArgumentException("The slide has no owning document.");
                box.Paragraphs[0].ClearInlineContent();
                using (OfficeDocumentContext.Push(document))
                using (OfficeDocumentContext.Push(new PowerPointTextScope(document, box))) Content.Invoke();
            }
            placed.Scope?.TrackFixedFigure(box);
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
