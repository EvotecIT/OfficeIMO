using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a designed slide, choosing its layout from typed content unless Layout is supplied.</summary>
/// <example>
/// <summary>Create a service review.</summary>
/// <code>New-OfficePowerPoint -Path .\Review.pptx {
///     PowerPointSlide -Title 'Service review' -Subtitle 'October 2026'
///     PowerPointSlide -Title 'Signals' { PowerPointCard 'Availability' -Text 'All regions healthy' }
/// }</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointSlide", DefaultParameterSetName = "Content")]
[Alias("PowerPointSlide")]
[OutputType(typeof(PowerPointSlide))]
public sealed class AddOfficePowerPointSlideCommand : OfficeDocumentCmdlet<PowerPointPresentation> {
    /// <summary>Gets or sets the slide's typed content block.</summary>
    [Parameter(Position = 0)]
    public ScriptBlock? Content { get; set; }
    /// <summary>Gets or sets the slide title.</summary>
    [Parameter]
    public string? Title { get; set; }
    /// <summary>Gets or sets the subtitle.</summary>
    [Parameter]
    public string? Subtitle { get; set; }
    /// <summary>Gets or sets the small label above the title.</summary>
    [Parameter]
    public string? Eyebrow { get; set; }
    /// <summary>Gets or sets the caption of the coverage panel.</summary>
    [Parameter]
    public string? MapLabel { get; set; }
    /// <summary>Gets or sets a designed layout; omission selects it from content.</summary>
    [Parameter]
    public PowerPointLayout Layout { get; set; } = PowerPointLayout.Auto;
    /// <summary>Gets or sets a specific template layout by name, one-based position, native type or returned metadata.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Template"), PowerPointLayoutTransformation]
    public PowerPointLayoutSelector? TemplateLayout { get; set; }
    /// <summary>Gets or sets a one-based master position for a template layout; defaults to the first.</summary>
    [Parameter(ParameterSetName = "Template"), ValidateRange(1, int.MaxValue)]
    public int Master { get; set; } = 1;
    /// <summary>Gets or sets a deterministic design seed for this slide.</summary>
    [Parameter]
    public string? Seed { get; set; }
    /// <summary>Gets or sets whether this slide omits its footer.</summary>
    [Parameter]
    public SwitchParameter NoFooter { get; set; }
    /// <summary>Gets or sets whether this slide omits its slide number.</summary>
    [Parameter]
    public SwitchParameter NoSlideNumber { get; set; }
    /// <summary>Gets or sets the background color for this slide and its continuations, including theme roles.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor Background { get; set; }
    /// <summary>Gets or sets a literal local background image path for this slide and its continuations.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? BackgroundImage { get; set; }
    /// <summary>Gets or sets the transition for this slide and its continuations.</summary>
    [Parameter]
    public PowerPointSlideTransition Transition { get; set; }
    /// <summary>Gets or sets whether the created physical slides are returned.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = GetDocument();
        if (ReferenceEquals(PowerPointContentScope.Current?.Document, document)) {
            InvalidArgument(new ArgumentException("A slide cannot be nested inside another slide of the same document."), document);
            return;
        }
        PowerPointTemplateLayoutInfo? template;
        string? image;
        bool backgroundBound = MyInvocation.BoundParameters.ContainsKey(nameof(Background));
        bool transitionBound = MyInvocation.BoundParameters.ContainsKey(nameof(Transition));
        try {
            if (backgroundBound && BackgroundImage != null)
                throw new ArgumentException("Choose Background or BackgroundImage for one slide.");
            if (transitionBound && !Enum.IsDefined(typeof(PowerPointSlideTransition), Transition))
                throw new ArgumentException("Choose a supported slide transition.", nameof(Transition));
            template = TemplateLayout?.Resolve(document, Master, MyInvocation.BoundParameters.ContainsKey(nameof(Master)));
            image = BackgroundImage == null ? null : OfficePath.ResolveInput(SessionState, BackgroundImage, literal: true)[0];
        }
        catch (ArgumentException error) { InvalidArgument(error, TemplateLayout); return; }
        using var content = new PowerPointContentScope(document);
        // The native transaction is rolled back if collection or layout fails.
        using (OfficeDocumentContext.Push(document))
        using (OfficeDocumentContext.Push(content)) Content?.Invoke();
        try {
            var slide = new PowerPointContentPlanSlide(Title, content.Blocks, Layout, Subtitle, Seed) {
                Eyebrow = Eyebrow, MapLabel = MapLabel, NoFooter = NoFooter, NoSlideNumber = NoSlideNumber, TemplateLayout = template
            };
            var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand(
                document.GetThemeColor(PowerPointThemeColor.Accent1) ?? "#2563EB", Seed ?? "document"));
            options.ApplyTheme = false;
            var result = content.Complete(slide, options, physicalSlides => {
                foreach (var physical in physicalSlides) {
                    if (image != null) physical.SetBackgroundImage(image);
                    else if (backgroundBound) physical.SetBackgroundColor(Background);
                    if (transitionBound) physical.Transition = Transition;
                }
            });
            if (PassThru) foreach (var physical in result) WriteObject(physical);
        } catch (ArgumentException error) { InvalidArgument(error, document); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-slide-create-failed", document);
        }
    }
}
