using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Replaces a native slide placeholder with plain text or rich paragraph content.</summary>
/// <para>Edits the existing text box, preserving its type, index and geometry. Use a New-OfficePowerPoint or Open-OfficePowerPoint content block to save completed edits atomically.</para>
/// <example><summary>Fill the current slide's title.</summary>
/// <code>PowerPointPlaceholder -Type Title 'Quarterly review'</code></example>
/// <example><summary>Fill a body placeholder with native rich paragraphs.</summary>
/// <code>Open-OfficePowerPoint -Path ./Review.pptx {
///     PowerPointPlaceholder -Slide 2 -Type Body {
///         PowerPointParagraph { PowerPointText 'Status: ' -Bold; PowerPointText 'Healthy' }
///         PowerPointParagraph 'All regions are operational.'
///     }
/// }</code></example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointPlaceholder", DefaultParameterSetName = "Text")]
[Alias("PowerPointPlaceholder")]
[OutputType(typeof(PowerPointTextBox))]
public sealed class SetOfficePowerPointPlaceholderCommand : OfficeCmdlet {
    /// <summary>Native slide or one-based slide number; omission uses the current slide block.</summary>
    [Parameter, PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }

    /// <summary>Placeholder type to update.</summary>
    [Parameter(Mandatory = true)]
    public PowerPointPlaceholderType Type { get; set; }

    /// <summary>Optional native placeholder index within the selected type.</summary>
    [Parameter]
    public uint? Index { get; set; }

    /// <summary>Replacement plain text; an empty string clears the placeholder.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Native rich paragraphs and runs to place in the existing text box.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content"), ValidateNotNull]
    public ScriptBlock? Content { get; set; }

    /// <summary>Leave the slide unchanged when the placeholder is absent; the content callback is not run.</summary>
    [Parameter]
    public SwitchParameter IgnoreMissing { get; set; }

    /// <summary>Return the same live native text box after editing.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var slide = PowerPointManualPlacement.Resolve(Slide).Slide;
            var placeholder = slide.GetPlaceholder(Type, Index);
            if (placeholder == null) {
                if (IgnoreMissing) return;
                throw new ArgumentException("Placeholder was not found on the selected slide.");
            }
            placeholder.Text = Content == null ? Text : string.Empty;
            if (Content != null) {
                var document = slide.OwnerDocument ?? throw new ArgumentException("The slide has no owning document.");
                placeholder.Paragraphs[0].ClearInlineContent();
                using (OfficeDocumentContext.Push(document))
                using (OfficeDocumentContext.Push(new PowerPointTextScope(document, placeholder))) Content.Invoke();
            }
            if (PassThru) WriteObject(placeholder);
        } catch (Exception error) {
            if (error is PipelineStoppedException or ActionPreferenceStopException) throw;
            if (PowerPointContentScope.Current != null)
                ThrowTerminatingError(OfficeError.Create(error, "powerpoint-placeholder-edit-failed", Slide, ErrorCategory.InvalidArgument));
            if (error is ArgumentException argument) InvalidArgument(argument, Slide);
            else ItemError(error, "powerpoint-placeholder-edit-failed", Slide);
        }
    }
}
