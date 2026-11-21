using System.Management.Automation;

namespace PSWriteOffice.Cmdlets.Security;

/// <summary>Chooses which hidden data Remove-OfficeHiddenData and Optimize-OfficePdf -Sanitize remove.</summary>
/// <para>With no parameters the options select the Default categories: document and custom properties, comments, personal
/// identities, speaker notes, the thumbnail, custom XML, image metadata, printer settings, external references, document
/// variables, revision identifiers and active content such as scripts. Tracked changes, hidden content, provenance,
/// attachments and bookmarks change what readers see or what the file discloses, so they are removed only when you name them.</para>
/// <example>
///   <summary>Also remove Content Credentials and AI labels, and accept tracked changes.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeHiddenDataOptions -Kind Default, Provenance -TrackedChanges AcceptAll
/// Remove-OfficeHiddenData -Path .\Proposal.docx -Options $options</code>
/// </example>
/// <example>
///   <summary>Remove AI-generated labels from a poster but keep its Content Credentials.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeHiddenDataOptions -Kind Default, Provenance -KeepProvenance C2pa, External
/// Remove-OfficeHiddenData -Path .\Poster.pdf -Options $options</code>
/// </example>
/// <example>
///   <summary>Clean everything, including hidden sheets and attachments.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Remove-OfficeHiddenData -Path .\Budget.xlsx -Options (New-OfficeHiddenDataOptions -Kind All)</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeHiddenDataOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeHiddenDataOptions))]
public sealed class NewOfficeHiddenDataOptionsCommand : PSCmdlet {
    /// <summary>Categories to remove, combined. Defaults to Default; All selects every category.</summary>
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty]
    public OfficeHiddenDataKind[] Kind { get; set; } = new[] { OfficeHiddenDataKind.Default };

    /// <summary>Keep, accept or reject tracked changes. Defaults to Keep, which reports them and leaves them in place.</summary>
    [Parameter]
    public OfficeTrackedChangesAction TrackedChanges { get; set; } = OfficeTrackedChangesAction.Keep;

    /// <summary>Provenance carriers to keep when Provenance is selected: C2pa (Content Credentials), External (links to Content Credentials) or AiSource (AI-generated labels).</summary>
    [Parameter]
    [ValidateSet("C2pa", "External", "AiSource")]
    public string[] KeepProvenance { get; set; } = Array.Empty<string>();

    /// <summary>Name that replaces authors on comments and revisions that stay in the file. Defaults to Author.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? AnonymousAuthor { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        OfficeHiddenDataKind kinds = OfficeHiddenDataKind.None;
        foreach (OfficeHiddenDataKind kind in Kind) kinds |= kind;
        if (kinds == OfficeHiddenDataKind.None) {
            ThrowTerminatingError(new ErrorRecord(new PSArgumentException("Select at least one hidden-data category."),
                "hidden-data-kind-required", ErrorCategory.InvalidArgument, Kind));
        }
        var options = new OfficeHiddenDataOptions { Kinds = kinds, TrackedChanges = TrackedChanges };
        if (AnonymousAuthor != null) options.AnonymousAuthor = AnonymousAuthor;
        foreach (string carrier in KeepProvenance) {
            if (carrier.Equals("C2pa", StringComparison.OrdinalIgnoreCase)) options.Provenance.RemoveC2paManifests = false;
            else if (carrier.Equals("External", StringComparison.OrdinalIgnoreCase)) options.Provenance.RemoveExternalC2paReferences = false;
            else options.Provenance.RemoveAiSourceMetadata = false;
        }
        WriteObject(options);
    }
}
