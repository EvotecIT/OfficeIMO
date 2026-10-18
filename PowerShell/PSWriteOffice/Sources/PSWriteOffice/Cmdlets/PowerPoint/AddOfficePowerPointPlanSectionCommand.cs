using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a semantic section/title slide to a PowerPoint deck plan.</summary>
/// <example>
///   <summary>Add a title section to a deck plan.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$plan = New-OfficePowerPointDeckPlan {
///     Add-OfficePowerPointPlanSection -Title 'Service Review' -Subtitle 'Monthly operating brief' -Seed 'service-review'
///     Add-OfficePowerPointPlanCardGrid -Title 'Signals' -Cards @(
///       @{ Title = 'Health'; Items = @('Green', 'No blockers') }
///     )
/// }</code>
///   <para>Adds a semantic opening section slide to the plan.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointPlanSection")]
[Alias("PptPlanSection")]
[OutputType(typeof(PowerPointDeckPlan))]
public sealed class AddOfficePowerPointPlanSectionCommand : PSCmdlet
{
    /// <summary>Plan to update. Optional inside New-OfficePowerPointDeckPlan.</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointDeckPlan? Plan { get; set; }

    /// <summary>Slide title.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Optional subtitle.</summary>
    [Parameter]
    public string? Subtitle { get; set; }

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
        plan.AddSection(Title, Subtitle, Seed);
        if (PassThru.IsPresent)
        {
            WriteObject(plan);
        }
    }
}
