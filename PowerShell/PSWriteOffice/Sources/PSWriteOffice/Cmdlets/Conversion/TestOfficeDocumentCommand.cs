using System;
using System.Management.Automation;
using System.Security;
using System.Threading;
using OfficeIMO;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Validates native document structure, schema, security, accessibility or compliance without changing the input.</summary>
/// <para>Unsupported checks are reported as not applicable. NativeReports retains native security, capability and PDF proof evidence. Options accepts typed check settings or hashtables; later values replace the matching settings group.</para>
/// <example>
/// <summary>Validate workbook structure</summary>
/// <code>Test-OfficeDocument -Path .\Reports\Sales.xlsx</code>
/// <para>Returns the common report and the native workbook doctor findings.</para>
/// </example>
/// <example>
/// <summary>Check untrusted packages</summary>
/// <code>Test-OfficeDocument -LiteralPath .\Uploads\Proposal.docx -Check Security -Untrusted -Quiet</code>
/// <para>Returns a Boolean after applying the untrusted package policy.</para>
/// </example>
[Cmdlet(VerbsDiagnostic.Test, "OfficeDocument", DefaultParameterSetName = "Path")]
[OutputType(typeof(OfficeValidationReport), typeof(bool))]
public sealed class TestOfficeDocumentCommand : PSCmdlet {
    private readonly CancellationTokenSource _cancellation = new();

    /// <summary>Filesystem paths with wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal filesystem paths; binds FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    [ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Caller-owned live native model, including unsaved edits.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document")]
    [ValidateNotNull]
    public object? Document { get; set; }
    /// <summary>Checks to perform. Defaults to Structure.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public OfficeDocumentCheck[] Check { get; set; } = new[] { OfficeDocumentCheck.Structure };
    /// <summary>Also applies the existing untrusted package-security policy.</summary>
    [Parameter]
    public SwitchParameter Untrusted { get; set; }
    /// <summary>Native check-specific options or flat hashtables, routed to the selected format through the shared binder. For PDF compliance checks, a flat Profile means the compliance profile; use typed PdfLoadOptions for a separate read profile.</summary>
    [Parameter]
    public object[]? Options { get; set; }
    /// <summary>Password for a path input. Accepts SecureString or PSCredential; plain strings are transformed with a warning.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>Returns IsValid as a Boolean instead of the report.</summary>
    [Parameter]
    public SwitchParameter Quiet { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Document") {
            object input = OfficeArgumentValue.Unwrap(Document)!;
            var literal = OfficePath.GetPipelineLiteralPaths(Document!);
            if (literal != null) ValidatePaths(literal, true);
            else if (input is string path) ValidatePaths(new[] { path }, false);
            else WriteValidation(input);
        } else ValidatePaths(ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath");
    }

    private void ValidatePaths(System.Collections.Generic.IEnumerable<string> inputs, bool literal) {
        foreach (string path in OfficePath.ResolveInputs(this, inputs, literal)) WriteValidation(path);
    }

    private void WriteValidation(object input) {
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        var options = OfficeDocumentValidationOptionBinding.Capture(engine, input, Options, Check);
        options.Untrusted = Untrusted.IsPresent;
        options.Password = Password.ToPlainText();
        if (options.Pdf?.DifferencePath != null) {
            var paths = OfficePath.ResolveInput(SessionState, options.Pdf.DifferencePath, literal: true);
            if (paths.Count != 1) throw new PSArgumentException("DifferencePath must identify one PDF.");
            options.Pdf.DifferencePath = paths[0];
        }
        var report = input is string path
            ? OfficeDocument.Validate(path, Check, options, _cancellation.Token)
            : OfficeDocument.Validate(input, Check, options, _cancellation.Token);
        WriteObject(Quiet.IsPresent ? (object)report.IsValid : report);
    }

    /// <inheritdoc />
    protected override void StopProcessing() => _cancellation.Cancel();
    /// <inheritdoc />
    protected override void EndProcessing() => _cancellation.Dispose();
}
