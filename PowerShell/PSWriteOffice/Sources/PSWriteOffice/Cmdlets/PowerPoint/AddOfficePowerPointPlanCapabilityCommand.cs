using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a semantic capability/content slide to a PowerPoint deck plan.</summary>
/// <example>
///   <summary>Add capability sections to a deck plan.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$sections = @(
///     @{ Heading = 'Monitoring'; Body = 'Signals and ownership'; Items = @('Alerts', 'Dashboards') }
///     @{ Heading = 'Reporting'; Body = 'Executive-ready output'; Items = @('Summary', 'Appendix') }
/// )
/// New-OfficePowerPointDeckPlan {
///     Add-OfficePowerPointPlanCapability -Title 'Capabilities' -Subtitle 'What the team provides' -Sections $sections
/// }</code>
///   <para>Adds a semantic capability/content slide to the plan.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointPlanCapability")]
[Alias("PptPlanCapability")]
[OutputType(typeof(PowerPointDeckPlan))]
public sealed class AddOfficePowerPointPlanCapabilityCommand : PSCmdlet
{
    /// <summary>Plan to update. Optional inside New-OfficePowerPointDeckPlan.</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointDeckPlan? Plan { get; set; }

    /// <summary>Slide title.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Optional slide subtitle.</summary>
    [Parameter]
    public string? Subtitle { get; set; }

    /// <summary>Objects with Heading/Title, optional Body, Items/Bullets/Details, and AccentColor properties.</summary>
    [Parameter(Mandatory = true)]
    public object[] Sections { get; set; } = System.Array.Empty<object>();

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
        plan.AddCapability(Title, Subtitle, PowerPointDesignerDataMapper.ToCapabilitySections(Sections), Seed);
        if (PassThru.IsPresent)
        {
            WriteObject(plan);
        }
    }
}
