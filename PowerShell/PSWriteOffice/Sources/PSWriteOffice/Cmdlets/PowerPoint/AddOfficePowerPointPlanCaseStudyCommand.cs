using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a semantic case-study slide to a PowerPoint deck plan.</summary>
/// <example>
///   <summary>Add a case-study slide with metrics.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$sections = @(
///     @{ Heading = 'Challenge'; Body = 'Manual reports took too long to produce.' }
///     @{ Heading = 'Outcome'; Body = 'Automated generation made the review repeatable.' }
/// )
/// $metrics = @(
///     @{ Value = '4h'; Label = 'saved each cycle' }
///     @{ Value = '0'; Label = 'manual copy steps' }
/// )
/// New-OfficePowerPointDeckPlan {
///     Add-OfficePowerPointPlanCaseStudy -Title 'Automation impact' -Sections $sections -Metrics $metrics
/// }</code>
///   <para>Adds a proof-oriented case-study slide to the plan.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointPlanCaseStudy")]
[Alias("PptPlanCaseStudy")]
[OutputType(typeof(PowerPointDeckPlan))]
public sealed class AddOfficePowerPointPlanCaseStudyCommand : PSCmdlet
{
    /// <summary>Plan to update. Optional inside New-OfficePowerPointDeckPlan.</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointDeckPlan? Plan { get; set; }

    /// <summary>Slide title.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Objects with Heading/Title and Body/Description/Text properties.</summary>
    [Parameter(Mandatory = true)]
    public object[] Sections { get; set; } = System.Array.Empty<object>();

    /// <summary>Objects with Value and Label/Name/Title properties.</summary>
    [Parameter]
    public object[]? Metrics { get; set; }

    /// <summary>Stable seed for deterministic visual selection.</summary>
    [Parameter]
    public string? Seed { get; set; }

    /// <summary>Emit the updated plan.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var plan = Plan ?? PowerPointDeckPlanDslContext.Require(this).Plan;
        plan.AddCaseStudy(
            Title,
            PowerPointDesignerDataMapper.ToCaseStudySections(Sections),
            PowerPointDesignerDataMapper.ToMetrics(Metrics),
            Seed);
        if (PassThru.IsPresent)
        {
            WriteObject(plan);
        }
    }
}
