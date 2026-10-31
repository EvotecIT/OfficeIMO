using System;
using System.Management.Automation;
using System.Net;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Inspects a live or saved presentation without changing or reopening a supplied document.</summary>
/// <example><summary>Inspect presentation files with literal names.</summary>
/// <code>Get-ChildItem ./Publish -Filter *.pptx | Test-OfficePowerPoint</code></example>
/// <example><summary>Include feature and visual evidence for an unsaved deck.</summary>
/// <code>$policy = New-OfficePowerPointInspectionOptions -Features -Visuals
/// Test-OfficePowerPoint -Document $deck -Options $policy</code></example>
[Cmdlet(VerbsDiagnostic.Test, "OfficePowerPoint", DefaultParameterSetName = "Document")]
[OutputType(typeof(PowerPointInspectionReport))]
public sealed class TestOfficePowerPointCommand : OfficeInputCmdlet {
    private OfficeInspectOptions _options = null!;

    /// <summary>Gets or sets a live document; omission uses the enclosing presentation block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document")]
    public PowerPointDocument? Document { get; set; }
    /// <summary>Gets or sets inspection, native format or Reader policy objects, including hashtables.</summary>
    [Parameter, OfficeInspectionOptionsTransformation]
    public object[]? Options { get; set; }
    /// <summary>Gets or sets the password for opening saved presentations.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath"), OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <inheritdoc />
    protected override Task BeginProcessingAsync() {
        _options = OfficeInspectionConfiguration.Bind(Options);
        _options.IncludeSummary = true;
        _options.IncludeContent = false;
        _options.Detailed = true;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Password)))
            _options.Reader.OpenPassword = Password == null ? null : new NetworkCredential(string.Empty, Password).Password;
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        if (ParameterSetName == "Document") {
            var document = Document ?? OfficeDocumentContext.Find<PowerPointDocument>();
            try {
                if (document == null) throw new PSArgumentException("Supply -Document or invoke Test-OfficePowerPoint inside a presentation block.");
                WriteObject(Report(OfficeDocument.Inspect(document, _options, CancelToken)));
            } catch (Exception error) when (error is not (OperationCanceledException or PipelineStoppedException or ActionPreferenceStopException)) {
                WriteError(OfficeError.Create(error, "powerpoint-inspection-failed", document));
            }
            return;
        }
        var inputs = ResolveInputs();
        await RunBatchAsync(inputs, (path, token) => Task.FromResult<object?>(Report(OfficeDocument.Inspect(path, _options, token))),
            "Inspecting presentations", "powerpoint-inspection-failed");
    }

    private static PowerPointInspectionReport Report(OfficeDocumentInfo result) =>
        result.Details as PowerPointInspectionReport ?? throw new NotSupportedException(
            "Expected a PowerPoint presentation; detected " + result.Format + ".");
}
