using System.Management.Automation;
using System.Security;
using OfficeIMO;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Accessibility;

/// <summary>Checks Word, Excel and PowerPoint files for accessibility issues. The input is never changed.</summary>
/// <para>Runs the rules of the Office Accessibility Checker: alternative text, table header rows, document title and language,
/// headings, hyperlink text, text contrast, slide titles, merged cells, default and hidden sheets, and blank cells or paragraphs
/// used for spacing. All three formats return the same finding codes and severities. Strict turns structural warnings into errors
/// and requires a document title; heuristic findings stay warnings.</para>
/// <example>
///   <summary>List the accessibility findings of a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>(Test-OfficeAccessibility -Path .\Budget.xlsx).Findings | Format-Table Severity, Code, Location, Message</code>
/// </example>
/// <example>
///   <summary>Fail a build when any document in a folder has accessibility errors.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-ChildItem .\Docs -Include *.docx, *.xlsx, *.pptx -Recurse | Test-OfficeAccessibility -Strict |
///     Where-Object { -not $_.IsSuccessful }</code>
/// </example>
/// <example>
///   <summary>Write a SARIF log for code-scanning tools.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Test-OfficeAccessibility -Path .\Deck.pptx -Sarif | Set-Content .\Deck.sarif</code>
/// </example>
[Cmdlet(VerbsDiagnostic.Test, "OfficeAccessibility", DefaultParameterSetName = "Path")]
[OutputType(typeof(OfficeAccessibilityFileReport), typeof(string), typeof(bool))]
public sealed class TestOfficeAccessibilityCommand : PSCmdlet {
    private readonly CancellationTokenSource _cancellation = new();

    /// <summary>Files to inspect; supports wildcards.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Exact file paths, including FileInfo.FullName and provider PSPath pipeline properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>An open Word, Excel or PowerPoint document, including unsaved edits.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Document")]
    [ValidateNotNull]
    public object? Document { get; set; }

    /// <summary>Uses the strict profile: structural warnings become errors and a document title is required.</summary>
    [Parameter]
    public SwitchParameter Strict { get; set; }

    /// <summary>Returns each report as a SARIF 2.1.0 log string instead of a report object.</summary>
    [Parameter]
    public SwitchParameter Sarif { get; set; }

    /// <summary>Returns a Boolean per input: true when there are no errors (and no warnings with IncludeWarnings).</summary>
    [Parameter]
    public SwitchParameter Quiet { get; set; }

    /// <summary>With Quiet, warnings also make the result false.</summary>
    [Parameter]
    public SwitchParameter IncludeWarnings { get; set; }

    /// <summary>Password for an encrypted path input. Accepts SecureString or PSCredential; plain strings are transformed with a warning.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    private OfficeAccessibilityPolicyProfile Profile => Strict.IsPresent ? OfficeAccessibilityPolicyProfile.Strict : OfficeAccessibilityPolicyProfile.Default;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Document") {
            object input = OfficeArgumentValue.Unwrap(Document)!;
            Emit(OfficeDocument.InspectAccessibility(input, Profile), null);
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string path in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) {
            if (Stopping) return;
            if (!OfficeDocument.SupportsAccessibility(path)) {
                WriteError(new ErrorRecord(new PSNotSupportedException("Accessibility checks support Word, Excel and PowerPoint files: " + path),
                    "accessibility-unsupported-format", ErrorCategory.InvalidArgument, path));
                continue;
            }
            try {
                Emit(OfficeDocument.InspectAccessibility(path, Profile, Password.ToPlainText(), _cancellation.Token), path);
            } catch (Exception error) when (error is not (ActionPreferenceStopException or PipelineStoppedException or OperationCanceledException)) {
                WriteError(new ErrorRecord(error, "accessibility-inspection-failed", ErrorCategory.ReadError, path));
            }
        }
    }

    private void Emit(OfficeAccessibilityReport report, string? path) {
        if (Quiet.IsPresent) WriteObject(report.ErrorCount == 0 && (!IncludeWarnings.IsPresent || report.WarningCount == 0));
        else if (Sarif.IsPresent) WriteObject(report.ToSarif(path));
        else WriteObject(new OfficeAccessibilityFileReport(report, path));
    }

    /// <inheritdoc />
    protected override void StopProcessing() => _cancellation.Cancel();
    /// <inheritdoc />
    protected override void EndProcessing() => _cancellation.Dispose();
}
