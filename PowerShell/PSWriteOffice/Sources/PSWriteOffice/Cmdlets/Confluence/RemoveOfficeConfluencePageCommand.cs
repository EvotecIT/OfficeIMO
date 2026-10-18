using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Confluence;

namespace PSWriteOffice.Cmdlets.Confluence;

/// <summary>Plans or deletes a Confluence Cloud page.</summary>
/// <example>
/// <summary>Inspect a permanent-delete request without contacting Confluence.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>Remove-OfficeConfluencePage -PageId 12345 -Purge -PlanOnly</code>
/// <para>Returns the exact DELETE request plan without using a session.</para>
/// </example>
/// <example>
/// <summary>Move a current page to the trash.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>Remove-OfficeConfluencePage -Session $session -PageId 12345</code>
/// <para>Uses PowerShell ShouldProcess before sending the non-retried delete request.</para>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficeConfluencePage", SupportsShouldProcess = true)]
[OutputType(typeof(ConfluencePageWritePlan))]
public sealed class RemoveOfficeConfluencePageCommand : AsyncPSCmdlet {
    /// <summary>Configured session required for a live delete operation.</summary>
    [Parameter]
    public ConfluenceSession? Session { get; set; }

    /// <summary>Page identifier.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [ValidateNotNullOrEmpty]
    public string PageId { get; set; } = string.Empty;

    /// <summary>Permanently delete a page that is already in the trash.</summary>
    [Parameter]
    public SwitchParameter Purge { get; set; }

    /// <summary>Delete a draft page.</summary>
    [Parameter]
    public SwitchParameter Draft { get; set; }

    /// <summary>Return the exact delete plan without contacting Confluence.</summary>
    [Parameter]
    public SwitchParameter PlanOnly { get; set; }

    /// <summary>Return the completed delete plan after a live operation.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        ConfluencePageWritePlan plan = ConfluenceClient.PlanDeletePage(PageId, Purge.IsPresent, Draft.IsPresent);
        if (PlanOnly.IsPresent) {
            WriteObject(plan);
            return;
        }

        if (Session == null) {
            throw new PSInvalidOperationException("Provide -Session for a live operation, or use -PlanOnly.");
        }

        if (!ShouldProcess(PageId, Purge.IsPresent ? "Permanently delete Confluence page" : "Delete Confluence page")) {
            return;
        }

        using var client = Session.CreateClient();
        await client.DeletePageAsync(PageId, Purge.IsPresent, Draft.IsPresent, CancelToken).ConfigureAwait(false);
        if (PassThru.IsPresent) {
            WriteObject(plan);
        }
    }
}