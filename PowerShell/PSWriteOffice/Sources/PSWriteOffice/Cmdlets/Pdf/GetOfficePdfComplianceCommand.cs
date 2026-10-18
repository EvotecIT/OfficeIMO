using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets a generated PDF document compliance readiness report.</summary>
/// <example>
///   <summary>Check generated PDF compliance readiness.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$pdf = New-OfficePdf {
///     Set-OfficePdfCompliance -Profile PdfA3B -Groundwork
///     Add-OfficePdfHeading -Text 'Compliance readiness'
/// } -NoSave
/// $pdf | Get-OfficePdfCompliance -Profile PdfA3B</code>
///   <para>Returns the OfficeIMO.Pdf readiness report before saving.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfCompliance", DefaultParameterSetName = ParameterSetDocument)]
[OutputType(typeof(PdfComplianceReadinessReport), typeof(PdfComplianceProofReport))]
public sealed class GetOfficePdfComplianceCommand : PSCmdlet
{
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Generated PDF document to assess outside the DSL context.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Existing PDF file path to assess after generation.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Compliance profile to assess. When omitted, the document's configured profile is used.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPath)]
    public PdfComplianceProfile? Profile { get; set; }

    /// <summary>Password used to inspect a Standard password-encrypted PDF.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Return a proof report that combines readiness with required external validator evidence placeholders.</summary>
    [Parameter]
    public SwitchParameter Proof { get; set; }

    /// <summary>External validator families whose result should be attached to the proof report.</summary>
    [Parameter]
    public PdfExternalValidatorKind[]? ExternalValidator { get; set; }

    /// <summary>
    /// Artifact-bound results produced by the external validation lane.
    /// Use PdfExternalValidationResult.PassedForArtifact or FromExitCodeForArtifact with the exact validated bytes.
    /// </summary>
    [Parameter]
    public PdfExternalValidationResult[]? ExternalValidation { get; set; }

    /// <summary>Unbound external validator status to attach when -ExternalValidator is provided.</summary>
    [Parameter]
    public PdfExternalValidationStatus ExternalStatus { get; set; } = PdfExternalValidationStatus.NotRun;

    /// <summary>Profile string reported by the external validator, for example PDF/A-3b.</summary>
    [Parameter]
    public string? ExternalProfile { get; set; }

    /// <summary>Human-readable external validation diagnostic.</summary>
    [Parameter]
    public string? ExternalDiagnostic { get; set; }

    /// <summary>Human-readable external validator name.</summary>
    [Parameter]
    public string? ExternalValidatorName { get; set; }

    /// <summary>External validator version recorded in the artifact-bound proof evidence.</summary>
    [Parameter]
    public string? ExternalValidatorVersion { get; set; }

    /// <summary>External validator process exit code. When provided, status is inferred from -ExternalSuccessExitCode.</summary>
    [Parameter]
    public int? ExternalExitCode { get; set; }

    /// <summary>External validator process exit code that means success.</summary>
    [Parameter]
    public int ExternalSuccessExitCode { get; set; } = 0;

    /// <summary>External validator executable path recorded in the proof evidence.</summary>
    [Parameter]
    public string? ExternalExecutablePath { get; set; }

    /// <summary>External validator command-line arguments recorded in the proof evidence.</summary>
    [Parameter]
    public string? ExternalArguments { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == ParameterSetPath)
        {
            var pathDocument = PdfDocument.Load(
                PdfCommandUtilities.ResolvePath(this, Path),
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));
            WriteObject(Proof.IsPresent
                ? AssessProof(pathDocument, Profile!.Value)
                : pathDocument.AssessCompliance(Profile!.Value));
            return;
        }

        var document = Document ?? throw new PSArgumentNullException(
            nameof(Document),
            "New-OfficePdf builds content as one composition. Run Get-OfficePdfCompliance on the document returned by New-OfficePdf.");
        var profile = Profile ?? document.AssessCompliance().Profile;
        WriteObject(Proof.IsPresent
            ? AssessProof(document, profile)
            : Profile.HasValue
                ? document.AssessCompliance(Profile.Value)
                : document.AssessCompliance());
    }

    private PdfComplianceProofReport AssessProof(PdfDocument document, PdfComplianceProfile profile)
    {
        var artifact = document.CreateComplianceArtifact(profile);
        return artifact.AssessProof(BuildExternalValidationResults());
    }

    private PdfExternalValidationResult[] BuildExternalValidationResults()
    {
        var suppliedResults = ExternalValidation ?? System.Array.Empty<PdfExternalValidationResult>();
        if (ExternalValidator == null || ExternalValidator.Length == 0)
        {
            return suppliedResults;
        }

        var results = new PdfExternalValidationResult[suppliedResults.Length + ExternalValidator.Length];
        System.Array.Copy(suppliedResults, results, suppliedResults.Length);
        for (int i = 0; i < ExternalValidator.Length; i++)
        {
            var validator = ExternalValidator[i];
            var name = string.IsNullOrWhiteSpace(ExternalValidatorName) ? validator.ToString() : ExternalValidatorName!;
            var diagnostic = string.IsNullOrWhiteSpace(ExternalDiagnostic) ? ExternalStatus.ToString() : ExternalDiagnostic!;
            var status = ExternalExitCode.HasValue
                ? ExternalExitCode.Value == ExternalSuccessExitCode
                    ? PdfExternalValidationStatus.Passed
                    : PdfExternalValidationStatus.Failed
                : ExternalStatus;
            results[suppliedResults.Length + i] = new PdfExternalValidationResult(
                validator,
                status,
                name,
                diagnostic,
                ExternalProfile,
                ExternalExecutablePath,
                ExternalArguments,
                ExternalExitCode,
                ExternalValidatorVersion);
        }

        return results;
    }
}
