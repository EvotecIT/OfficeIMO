using System.IO;
using System.Management.Automation;
using OfficeIMO;

namespace PSWriteOffice.Cmdlets.Security;

/// <summary>Inspects an Open XML or compound Office package without opening active content.</summary>
/// <example>
///   <summary>Inventory active content and reject it under an untrusted-input policy.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$report = Get-OfficePackageSecurity -Path .\Incoming\Report.xlsm -Untrusted
/// $report | Select-Object IsValid, MacroPartCount, EmbeddedPayloadPartCount, ExternalRelationshipCount
/// $report.Findings | Format-Table Severity, Rule, PartName, Message</code>
///   <para>Returns observations and policy violations; it does not execute package content.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePackageSecurity")]
[OutputType(typeof(OfficePackageSecurityReport))]
public sealed class GetOfficePackageSecurityCommand : PSCmdlet
{
    /// <summary>Path to an Open XML or compound Office package.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Custom package size, expansion, and active-content policy.</summary>
    [Parameter]
    public OfficePackageSecurityOptions? Options { get; set; }

    /// <summary>Use the bounded policy that rejects macros, embedded payloads, ActiveX, and external relationships.</summary>
    [Parameter]
    public SwitchParameter Untrusted { get; set; }

    /// <summary>Throw on the first policy violation instead of returning only the report.</summary>
    [Parameter]
    public SwitchParameter ThrowOnViolation { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Options != null && Untrusted.IsPresent)
        {
            throw new PSArgumentException("-Options and -Untrusted cannot be combined.");
        }

        var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var options = Untrusted.IsPresent ? OfficePackageSecurityOptions.UntrustedDefaults : Options;
        using var stream = File.OpenRead(path);
        WriteObject(ThrowOnViolation.IsPresent
            ? OfficePackageSecurityInspector.Validate(stream, options)
            : OfficePackageSecurityInspector.Inspect(stream, options));
    }
}
