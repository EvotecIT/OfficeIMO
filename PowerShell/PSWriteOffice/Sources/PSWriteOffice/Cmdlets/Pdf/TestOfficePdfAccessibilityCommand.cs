using System.Management.Automation;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Checks a PDF against PDF/UA-1 using the Matterhorn Protocol failure conditions.</summary>
/// <para>Works on any PDF, including files OfficeIMO did not produce. Machine-checkable conditions are evaluated;
/// conditions that need a person's judgement are returned with status NeedsReview and are never reported as passed.</para>
/// <example>
///   <summary>List the failed accessibility conditions of a PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$report = Test-OfficePdfAccessibility -Path .\Report.pdf
/// $report.FailedConditions | Select-Object @{ n = 'Condition'; e = { $_.Condition.ConditionId } }, OccurrenceCount</code>
///   <para>Returns the accessibility report; Findings carry page, object and structure-element locations.</para>
/// </example>
/// <example>
///   <summary>Show where each failure is.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>(Test-OfficePdfAccessibility -Path .\Report.pdf).Findings |
///     Where-Object Severity -eq 'Error' |
///     Format-Table ConditionId, PageNumber, ObjectNumber, StructureType, Message</code>
///   <para>Lists the failing objects with the Matterhorn condition they break.</para>
/// </example>
[Cmdlet(VerbsDiagnostic.Test, "OfficePdfAccessibility", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(PdfAccessibilityReport))]
public sealed class TestOfficePdfAccessibilityCommand : PSCmdlet, IDisposable
{
    private readonly CancellationTokenSource _stop = new();
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";
    private const string ParameterSetLiteralPath = "LiteralPath";

    /// <summary>PDF document to check.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>PDF file paths to check; wildcards are expanded.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal PDF file paths to check.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetLiteralPath)]
    [Alias("FullName", "PSPath")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Native input loading policy.</summary>
    [Parameter(ParameterSetName = ParameterSetPath), Parameter(ParameterSetName = ParameterSetLiteralPath), OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }

    /// <summary>Password used to open a Standard password-encrypted PDF.</summary>
    [Parameter(ParameterSetName = ParameterSetPath), Parameter(ParameterSetName = ParameterSetLiteralPath)]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter(ParameterSetName = ParameterSetPath), Parameter(ParameterSetName = ParameterSetLiteralPath)]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        PdfInspectionSource.Visit(this, ParameterSetName, Path, LiteralPath, Document, LoadOptions, Password,
            IgnorePermissionRestrictions.IsPresent, document => WriteObject(document.Proof.CheckAccessibility(_stop.Token)));
    }

    /// <inheritdoc />
    protected override void StopProcessing() => _stop.Cancel();

    /// <summary>Releases this invocation's cancellation source.</summary>
    public void Dispose() => _stop.Dispose();
}
