using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Changes a selected slide's title, template layout, background or transition.</summary>
/// <example><summary>Rename the first slide.</summary>
/// <code>Set-OfficePowerPointSlide 1 -Document $deck -Title 'Quarterly review'</code></example>
/// <example><summary>Apply transitions through the slide pipeline.</summary>
/// <code>Get-OfficePowerPointSlide -Document $deck | Set-OfficePowerPointSlide -Transition Fade</code></example>
/// <example><summary>Select a layout from the document's template.</summary>
/// <code>Set-OfficePowerPointSlide 2 -Document $deck -TemplateLayout 'Title and Content'</code></example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointSlide", DefaultParameterSetName = "Default")]
[OutputType(typeof(PowerPointSlide))]
public sealed class SetOfficePowerPointSlideCommand : OfficeCmdlet {
    /// <summary>Gets or sets a native slide or one-based slide position.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true), PowerPointSlideTransformation]
    public PowerPointSlideSelector Slide { get; set; } = null!;
    /// <summary>Gets or sets the document for numeric slide and template layout selection; omission uses the enclosing document block.</summary>
    [Parameter, ValidateNotNull]
    public PowerPointPresentation? Document { get; set; }
    /// <summary>Gets or sets the title text; an empty string clears an existing title.</summary>
    [Parameter, AllowEmptyString]
    public string? Title { get; set; }
    /// <summary>Gets or sets a template layout by literal name, one-based position, native type or returned metadata. Supply Document when selecting a template layout.</summary>
    [Parameter, PowerPointLayoutTransformation]
    public PowerPointLayoutSelector? TemplateLayout { get; set; }
    /// <summary>Gets or sets a one-based template master position; defaults to the first.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int Master { get; set; } = 1;
    /// <summary>Gets or sets the background color, including theme roles and transparent literals.</summary>
    [Parameter(ParameterSetName = "Default"), OfficeColorTransformation]
    public OfficeColor Background { get; set; }
    /// <summary>Gets or sets a literal local background image path.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "BackgroundImage"), ValidateNotNullOrEmpty]
    public string? BackgroundImage { get; set; }
    /// <summary>Gets or sets whether to remove the slide's background override.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "ClearBackground")]
    public SwitchParameter ClearBackground { get; set; }
    /// <summary>Gets or sets the slide transition.</summary>
    [Parameter]
    public PowerPointSlideTransition Transition { get; set; }
    /// <summary>Gets or sets whether the original native slide is returned.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = Document ?? OfficeDocumentContext.Find<PowerPointPresentation>();
        try {
            PowerPointSlide slide = Slide.Resolve(document);
            if (TemplateLayout == null && MyInvocation.BoundParameters.ContainsKey(nameof(Master)))
                throw new ArgumentException("Master requires TemplateLayout.", nameof(Master));
            PowerPointTemplateLayoutInfo? layout = TemplateLayout?.Resolve(ResolveDocument(document), Master,
                MyInvocation.BoundParameters.ContainsKey(nameof(Master)));
            string? image = null;
            if (BackgroundImage != null) {
                image = OfficePath.ResolveInput(SessionState, BackgroundImage, literal: true)[0];
            }
            if (image != null) slide.SetBackgroundImage(image);
            else if (ClearBackground) slide.ClearBackground();
            else if (MyInvocation.BoundParameters.ContainsKey(nameof(Background))) slide.SetBackgroundColor(Background);
            if (layout != null) slide.SetLayout(layout.MasterIndex, layout.LayoutIndex);
            if (MyInvocation.BoundParameters.ContainsKey(nameof(Title))) {
                var title = slide.GetPlaceholder(PowerPointPlaceholderType.Title) ?? slide.GetPlaceholder(PowerPointPlaceholderType.CenteredTitle);
                if (title != null) title.Text = Title ?? string.Empty;
                else if (!string.IsNullOrEmpty(Title)) slide.AddTitle(Title!);
            }
            if (MyInvocation.BoundParameters.ContainsKey(nameof(Transition))) slide.Transition = Transition;
            if (PassThru) WriteObject(slide);
        } catch (ArgumentException error) { InvalidArgument(error, Slide); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-slide-edit-failed", Slide);
        }
    }
}
