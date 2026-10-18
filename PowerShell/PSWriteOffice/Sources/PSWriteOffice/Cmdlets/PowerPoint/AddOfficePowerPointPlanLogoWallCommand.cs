using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a semantic logo/proof-wall slide to a PowerPoint deck plan.</summary>
/// <example>
///   <summary>Add a proof wall to a deck plan.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$logos = @(
///     @{ Name = 'Directory'; Subtitle = 'Identity platform'; AccentColor = '#2563EB' }
///     @{ Name = 'Mail'; Subtitle = 'Messaging platform'; AccentColor = '#0F766E' }
/// )
/// New-OfficePowerPointDeckPlan {
///     Add-OfficePowerPointPlanLogoWall -Title 'Systems covered' -Subtitle 'Representative services' -Logos $logos
/// }</code>
///   <para>Adds a semantic logo/proof-wall slide to the plan.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointPlanLogoWall")]
[Alias("PptPlanLogoWall")]
[OutputType(typeof(PowerPointDeckPlan))]
public sealed class AddOfficePowerPointPlanLogoWallCommand : PSCmdlet
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

    /// <summary>Objects with Name, optional Subtitle, ImagePath, and AccentColor properties.</summary>
    [Parameter(Mandatory = true)]
    public object[] Logos { get; set; } = System.Array.Empty<object>();

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
        plan.AddLogoWall(Title, Subtitle, PowerPointDesignerDataMapper.ToLogoItems(Logos), Seed);
        if (PassThru.IsPresent)
        {
            WriteObject(plan);
        }
    }
}
