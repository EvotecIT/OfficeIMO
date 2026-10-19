using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Security;

namespace PSWriteOffice.Cmdlets.Security;

/// <summary>Returns OfficeIMO's machine-readable protected-content support contract.</summary>
/// <example>
///   <summary>List incomplete protection capabilities.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeProtectionCapability -IncompleteOnly | Format-Table FormatId, Kind, Open, Create, Validate</code>
///   <para>Shows formats whose encrypted, signed, restricted, or obfuscated workflows still have unsupported operations.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeProtectionCapability")]
[OutputType(typeof(OfficeProtectionCapability))]
[OutputType(typeof(string))]
public sealed class GetOfficeProtectionCapabilityCommand : PSCmdlet
{
    /// <summary>Exact stable capability identifier.</summary>
    [Parameter(Position = 0)]
    public string? Id { get; set; }

    /// <summary>Filter by format identifier or family, such as PDF, DOCX, or EML.</summary>
    [Parameter]
    public string? Format { get; set; }

    /// <summary>Filter by protected-content mechanism.</summary>
    [Parameter]
    public OfficeProtectionKind? Kind { get; set; }

    /// <summary>Return only rows with at least one unsupported operation.</summary>
    [Parameter]
    public SwitchParameter IncompleteOnly { get; set; }

    /// <summary>Return the complete catalog as deterministic JSON. Filtering parameters cannot be combined with this switch.</summary>
    [Parameter]
    public SwitchParameter AsJson { get; set; }

    /// <summary>Return the complete catalog as a Markdown table. Filtering parameters cannot be combined with this switch.</summary>
    [Parameter]
    public SwitchParameter AsMarkdown { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var catalog = OfficeProtectionCapabilityCatalog.Current;
        if (AsJson.IsPresent || AsMarkdown.IsPresent)
        {
            if (AsJson.IsPresent && AsMarkdown.IsPresent)
            {
                ThrowTerminatingError(new ErrorRecord(
                    new PSArgumentException("Specify only one of -AsJson or -AsMarkdown."),
                    "ProtectionCapabilityOutputConflict",
                    ErrorCategory.InvalidArgument,
                    null));
                return;
            }

            if (!string.IsNullOrWhiteSpace(Id) || !string.IsNullOrWhiteSpace(Format) || Kind.HasValue || IncompleteOnly.IsPresent)
            {
                ThrowTerminatingError(new ErrorRecord(
                    new PSArgumentException("Catalog text output cannot be combined with row filters."),
                    "ProtectionCapabilityFilterConflict",
                    ErrorCategory.InvalidArgument,
                    null));
                return;
            }

            WriteObject(AsJson.IsPresent ? catalog.ToJson() : catalog.ToMarkdown());
            return;
        }

        var capabilities = catalog.Capabilities.AsEnumerable();
        if (!string.IsNullOrWhiteSpace(Id))
        {
            var id = Id!;
            capabilities = capabilities.Where(capability =>
                string.Equals(capability.Id, id.Trim(), StringComparison.Ordinal));
        }

        if (!string.IsNullOrWhiteSpace(Format))
        {
            var format = Format!;
            capabilities = capabilities.Where(capability =>
                capability.FormatId.IndexOf(format.Trim(), StringComparison.OrdinalIgnoreCase) >= 0);
        }

        if (Kind.HasValue)
        {
            capabilities = capabilities.Where(capability => capability.Kind == Kind.Value);
        }

        if (IncompleteOnly.IsPresent)
        {
            var incomplete = catalog.IncompleteCapabilities.Select(capability => capability.Id).ToHashSet(StringComparer.Ordinal);
            capabilities = capabilities.Where(capability => incomplete.Contains(capability.Id));
        }

        WriteObject(capabilities.ToArray(), enumerateCollection: true);
    }
}
