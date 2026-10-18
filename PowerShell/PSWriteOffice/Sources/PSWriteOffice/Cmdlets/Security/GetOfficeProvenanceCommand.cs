using System.Management.Automation;
using OfficeIMO.Provenance;
using OfficeIMO.Provenance.C2pa;

namespace PSWriteOffice.Cmdlets.Security;

/// <summary>Collects structural, text-integrity, and optional provider verification evidence for a file.</summary>
/// <para>The result keeps carrier discovery, cryptographic verification, and provider signals separate; it does not infer authorship from their presence or absence.</para>
/// <example>
///   <summary>Inspect embedded provenance without invoking an external verifier.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$evidence = Get-OfficeProvenance -Path .\Published\cover.png
/// $evidence.Structural | Select-Object Format, HasC2paManifest
/// $evidence.TextIntegrity</code>
/// </example>
/// <example>
///   <summary>Ask an explicitly supplied c2patool installation to verify Content Credentials.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$evidence = Get-OfficeProvenance -Path .\Published\cover.png -C2paToolPath C:\Tools\c2patool.exe
/// $evidence.Verification | Select-Object Status, ProviderName, Findings</code>
///   <para>Network access remains disabled unless enabled in the supplied assessment options.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeProvenance")]
[OutputType(typeof(OfficeProvenanceAssessmentReport))]
public sealed class GetOfficeProvenanceCommand : PSCmdlet
{
    /// <summary>Asset or document path to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Structural, text-integrity, and provider verification limits and policy.</summary>
    [Parameter]
    public OfficeProvenanceAssessmentOptions? Options { get; set; }

    /// <summary>Explicit c2patool executable path or controlled command name used for cryptographic verification.</summary>
    [Parameter]
    public string? C2paToolPath { get; set; }

    /// <summary>Optional provider-specific watermark or disclosure detectors.</summary>
    [Parameter]
    public IOfficeProvenanceSignalDetector[]? SignalDetector { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        IOfficeProvenanceVerifier? verifier = string.IsNullOrWhiteSpace(C2paToolPath)
            ? null
            : new C2paToolProvenanceVerifier(C2paToolPath!);
        WriteObject(OfficeProvenanceAssessment.InspectFile(path, Options, verifier, SignalDetector));
    }
}
