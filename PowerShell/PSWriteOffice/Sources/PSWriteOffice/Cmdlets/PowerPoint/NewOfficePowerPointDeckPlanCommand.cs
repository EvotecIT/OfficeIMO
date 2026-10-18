using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates a semantic PowerPoint deck plan for designer rendering.</summary>
/// <example>
///   <summary>Create a semantic service brief plan.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$plan = New-OfficePowerPointDeckPlan {
///     Add-OfficePowerPointPlanSection -Title 'Service Review' -Subtitle 'Monthly operating brief'
///     Add-OfficePowerPointPlanProcess -Title 'Operating rhythm' -Steps @(
///       @{ Title = 'Collect'; Body = 'Gather health signals' }
///       @{ Title = 'Review'; Body = 'Confirm owner decisions' }
///       @{ Title = 'Publish'; Body = 'Share the final brief' }
///     )
/// }
/// New-OfficePowerPoint -Path .\Examples\Documents\DesignerDeck.pptx {
///     Add-OfficePowerPointDesignerDeck -Plan $plan
/// }</code>
///   <para>Builds a deck plan and renders it through the OfficeIMO designer helpers.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePowerPointDeckPlan")]
[Alias("PptDeckPlan")]
[OutputType(typeof(PowerPointDeckPlan))]
public sealed class NewOfficePowerPointDeckPlanCommand : PSCmdlet
{
    /// <summary>Nested deck-plan DSL content.</summary>
    [Parameter(Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var plan = new PowerPointDeckPlan();
        if (Content != null)
        {
            using (PowerPointDeckPlanDslContext.Enter(plan))
            {
                Content.InvokeReturnAsIs();
            }
        }

        WriteObject(plan);
    }
}
