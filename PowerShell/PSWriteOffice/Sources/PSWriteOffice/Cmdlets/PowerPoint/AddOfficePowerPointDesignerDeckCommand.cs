using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Renders a semantic deck plan through OfficeIMO PowerPoint designer helpers.</summary>
/// <example>
///   <summary>Render a designer deck from a semantic plan.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$plan = New-OfficePowerPointDeckPlan {
///     Add-OfficePowerPointPlanSection -Title 'Service Review'
///     Add-OfficePowerPointPlanCardGrid -Title 'Current signals' -Cards @(
///       @{ Title = 'Availability'; Items = @('Healthy', 'No critical incidents') }
///       @{ Title = 'Risk'; Items = @('One dependency on watch') }
///     )
/// }
/// New-OfficePowerPoint -Path .\Examples\Documents\DesignerDeck.pptx {
///     Add-OfficePowerPointDesignerDeck -Plan $plan -AccentColor '#0F766E' -Purpose 'monthly service brief'
/// }</code>
///   <para>Uses OfficeIMO design selection to turn semantic content into slides.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointDesignerDeck")]
[Alias("PptDesignerDeck")]
[OutputType(typeof(PowerPointDeckPlanSlideRenderSummary), typeof(PowerPointSlide))]
public sealed class AddOfficePowerPointDesignerDeckCommand : PSCmdlet
{
    /// <summary>Presentation to update. Optional inside New-OfficePowerPoint.</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Deck plan to render.</summary>
    [Parameter(Mandatory = true)]
    public PowerPointDeckPlan Plan { get; set; } = null!;

    /// <summary>Brand accent color used to derive the deck palette.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string AccentColor { get; set; } = "#008C95";

    /// <summary>Stable seed used for deterministic design choices.</summary>
    [Parameter]
    public string Seed { get; set; } = "pswriteoffice";

    /// <summary>Plain-language purpose used to select a built-in design recipe.</summary>
    [Parameter]
    public string Purpose { get; set; } = "technical service brief";

    /// <summary>Deck theme name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Default slide eyebrow.</summary>
    [Parameter]
    public string? Eyebrow { get; set; }

    /// <summary>Left footer text.</summary>
    [Parameter]
    public string? FooterLeft { get; set; }

    /// <summary>Right footer text.</summary>
    [Parameter]
    public string? FooterRight { get; set; }

    /// <summary>Creative direction pack name, such as Boardroom, FieldProof, TechnicalMap, or QuietAppendix.</summary>
    [Parameter]
    public PowerPointCreativeDirectionPack? CreativeDirectionPack { get; set; }

    /// <summary>Auto layout strategy, such as ContentFirst, DesignFirst, Compact, or VisualFirst.</summary>
    [Parameter]
    public PowerPointAutoLayoutStrategy? LayoutStrategy { get; set; }

    /// <summary>Design alternative count to consider. 0 uses OfficeIMO defaults.</summary>
    [Parameter]
    public int AlternativeCount { get; set; } = 3;

    /// <summary>Do not automatically apply the design theme to the presentation.</summary>
    [Parameter]
    public SwitchParameter NoApplyTheme { get; set; }

    /// <summary>Preview resolved slides without rendering them.</summary>
    [Parameter]
    public SwitchParameter Preview { get; set; }

    /// <summary>Emit rendered slides instead of the render summary.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Plan == null)
        {
            throw new PSArgumentException("Provide a deck plan.", nameof(Plan));
        }

        if (AlternativeCount < 0)
        {
            throw new PSArgumentException("AlternativeCount must be zero or greater.", nameof(AlternativeCount));
        }

        var brief = PowerPointDesignBrief
            .FromBrand(AccentColor, Seed, Purpose)
            .WithIdentity(Name, Eyebrow, FooterLeft, FooterRight);

        if (CreativeDirectionPack.HasValue)
        {
            brief.WithCreativeDirectionPack(CreativeDirectionPack.Value);
        }

        if (LayoutStrategy.HasValue)
        {
            brief.WithLayoutStrategy(LayoutStrategy.Value);
        }

        if (Preview.IsPresent)
        {
            WriteObject(brief.DescribeDeckPlan(Plan, AlternativeCount), enumerateCollection: true);
            return;
        }

        var presentation = Presentation ?? PowerPointDslContext.Require(this).Presentation;
        var slideIndexOffset = presentation.Slides.Count;
        var options = PowerPointCompositionOptions.FromBrief(brief);
        options.AlternativeCount = AlternativeCount;
        options.ApplyTheme = !NoApplyTheme.IsPresent;
        var result = presentation.Compose(Plan, options);
        var summary = result.Plan.DescribeSlides(result.Design, slideIndexOffset);

        if (PassThru.IsPresent)
        {
            WriteObject(result.Slides, enumerateCollection: true);
        }
        else
        {
            WriteObject(summary, enumerateCollection: true);
        }
    }
}
