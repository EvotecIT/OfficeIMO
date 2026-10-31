using System;
using System.Collections;
using System.Management.Automation;
using PSWriteOffice.Foundation;
using OfficeIMO.Provenance;
using OfficeIMO.Provenance.C2pa;

namespace PSWriteOffice.Cmdlets.Security;

/// <summary>Collects structural, text-integrity, and optional provider verification evidence for a file.</summary>
/// <para>The result keeps carrier discovery, cryptographic verification, and provider signals separate; it does not infer authorship from their presence or absence.</para>
/// <example>
///   <summary>Inspect embedded provenance without invoking an external verifier.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$evidence = Get-OfficeDocumentProvenance -Path .\Published\cover.png
/// $evidence.Structural | Select-Object Format, HasC2paManifest
/// $evidence.TextIntegrity</code>
/// </example>
/// <example>
///   <summary>Ask an explicitly supplied c2patool installation to verify Content Credentials.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$evidence = Get-OfficeDocumentProvenance -Path .\Published\cover.png -C2paToolPath C:\Tools\c2patool.exe
/// $evidence.Verification | Select-Object Status, ProviderName, Findings</code>
///   <para>Network access remains disabled unless enabled in the supplied assessment options.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentProvenance", DefaultParameterSetName = "Path")]
[OutputType(typeof(OfficeProvenanceAssessmentReport))]
public sealed class GetOfficeDocumentProvenanceCommand : PSCmdlet
{
    private OfficeProvenanceAssessmentOptions? _options;

    /// <summary>Asset or document paths, including provider wildcards.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal paths, including FileInfo.FullName and provider PSPath pipeline properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Typed assessment policies or hashtables with Structural, TextIntegrity and Verification groups. The last policy takes precedence.</summary>
    [Parameter, OfficeProvenanceOptionsTransformation]
    public object[]? Options { get; set; }

    /// <summary>Explicit c2patool executable path or controlled command name used for cryptographic verification.</summary>
    [Parameter]
    public string? C2paToolPath { get; set; }

    /// <summary>Optional provider-specific watermark or disclosure detectors.</summary>
    [Parameter]
    public IOfficeProvenanceSignalDetector[]? SignalDetector { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing() {
        foreach (OfficeProvenanceAssessmentOptions policy in Options ?? Array.Empty<object>()) _options = policy;
    }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        IOfficeProvenanceVerifier? verifier = string.IsNullOrWhiteSpace(C2paToolPath)
            ? null : new C2paToolProvenanceVerifier(C2paToolPath!);
        bool literal = ParameterSetName == "LiteralPath";
        // Detectors need not be thread-safe; keep this read-only batch sequential.
        foreach (string input in literal ? LiteralPath : Path) {
            try {
                foreach (string path in OfficePath.ResolveInput(SessionState, input, literal, reportItemError:
                    (error, item) => WriteError(OfficeError.Create(error, "input-unreadable", item))))
                    Inspect(path, verifier);
            } catch (Exception error) when (error is not (ActionPreferenceStopException or PipelineStoppedException)) {
                WriteError(OfficeError.Create(error, "input-unreadable", input));
            }
        }
    }

    private void Inspect(string path, IOfficeProvenanceVerifier? verifier) {
        if (Stopping) return;
        try { WriteObject(OfficeProvenanceAssessment.InspectFile(path, _options, verifier, SignalDetector)); }
        catch (Exception error) when (error is not (ActionPreferenceStopException or PipelineStoppedException)) {
            WriteError(OfficeError.Create(error, "provenance-inspection-failed", path));
        }
    }
}

/// <summary>Binds the provenance owner's read-only option groups without changing the shared options binder.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
internal sealed class OfficeProvenanceOptionsTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        object? inputValue = OfficeArgumentValue.Unwrap(inputData);
        object[] values = inputValue is object[] array ? array : new[] { inputValue! };
        var result = new System.Collections.Generic.List<OfficeProvenanceAssessmentOptions>();
        foreach (object input in values) {
            object? value = OfficeArgumentValue.Unwrap(input);
            if (value is OfficeProvenanceAssessmentOptions typed) { result.Add(typed); continue; }
            if (value is not IDictionary dictionary)
                throw new PSArgumentException("Expected OfficeProvenanceAssessmentOptions or a hashtable.");
            var policy = new OfficeProvenanceAssessmentOptions();
            foreach (DictionaryEntry entry in dictionary) {
                object? group = entry.Key is string key ? key.ToLowerInvariant() switch {
                    "structural" => policy.Structural,
                    "textintegrity" => policy.TextIntegrity,
                    "verification" => policy.Verification,
                    _ => null
                } : null;
                var item = new Hashtable { [entry.Key] = entry.Value };
                if (group == null)
                    new OfficeOptionsTransformationAttribute(typeof(OfficeProvenanceAssessmentOptions)).ApplyTo(
                        engineIntrinsics, policy, item);
                else if (OfficeArgumentValue.Unwrap(entry.Value) is IDictionary settings)
                    new OfficeOptionsTransformationAttribute(group.GetType()).ApplyTo(engineIntrinsics, group, settings);
                else throw new PSArgumentException($"Expected a hashtable for provenance group '{entry.Key}'.");
            }
            result.Add(policy);
        }
        return result.ToArray();
    }
}
