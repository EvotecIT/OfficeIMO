using System.Management.Automation;
using OfficeIMO.Confluence;

namespace PSWriteOffice.Cmdlets.Confluence;

/// <summary>Safely replaces one marker-delimited section in a Confluence storage body.</summary>
/// <example>
/// <summary>Replace a generated report while preserving owner-authored content.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>$result = Set-OfficeConfluenceManagedSection -ExistingBody $storage -SectionId daily-report -Replacement $html -AppendIfMissing</code>
/// <para>Returns before/after hashes and the updated body without contacting Confluence.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeConfluenceManagedSection")]
[OutputType(typeof(ConfluenceManagedSectionResult), typeof(string))]
public sealed class SetOfficeConfluenceManagedSectionCommand : PSCmdlet
{
    /// <summary>Existing Confluence storage body.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    [AllowEmptyString]
    public string ExistingBody { get; set; } = string.Empty;

    /// <summary>Stable marker identifier.</summary>
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string SectionId { get; set; } = string.Empty;

    /// <summary>Replacement storage-format content.</summary>
    [Parameter(Mandatory = true)]
    [AllowEmptyString]
    public string Replacement { get; set; } = string.Empty;

    /// <summary>Append a new marker pair when the section does not exist.</summary>
    [Parameter]
    public SwitchParameter AppendIfMissing { get; set; }

    /// <summary>Return only the updated body string.</summary>
    [Parameter]
    public SwitchParameter PassThruBody { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var result = ConfluenceManagedSection.Apply(
            ExistingBody,
            SectionId,
            Replacement,
            AppendIfMissing.IsPresent ? ConfluenceMissingSectionBehavior.Append : ConfluenceMissingSectionBehavior.Fail);
        WriteObject(PassThruBody.IsPresent ? result.UpdatedBody : result);
    }
}
