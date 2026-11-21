using System.Management.Automation;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models;

namespace PSWriteOffice.Cmdlets.Security;

/// <summary>Reports the metadata and hidden data that would leave with a file. The file is never changed.</summary>
/// <para>Finds author and company properties, comments, tracked changes, hidden text, rows, sheets and slides, speaker notes,
/// image metadata such as GPS locations, scripts, attachments, Content Credentials and AI labels. Inputs are Word, Excel,
/// PowerPoint, Visio, OpenDocument, EPUB, PDF, JPEG, PNG, WebP, GIF and TIFF files. Each finding has the Outcome Found.</para>
/// <example>
///   <summary>List what would leave with a proposal.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>(Get-OfficeHiddenData -Path .\Proposal.docx).Findings</code>
/// </example>
/// <example>
///   <summary>Find the files in an outgoing folder that still carry comments or tracked changes.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-ChildItem .\Outgoing -File | Get-OfficeHiddenData |
///     Where-Object { $_.Findings.Kind -contains 'Comments' -or $_.Findings.Kind -contains 'TrackedChanges' }</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeHiddenData", DefaultParameterSetName = "Path")]
[OutputType(typeof(OfficeHiddenDataFileReport))]
public sealed class GetOfficeHiddenDataCommand : OfficeCmdlet {
    /// <summary>Files to inspect; supports wildcards.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Exact file paths, including FileInfo.FullName and provider PSPath pipeline properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string path in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) {
            if (Stopping) return;
            if (!OfficeDocument.SupportsHiddenData(path)) {
                ItemError(OfficeHiddenDataFiles.Unsupported(path), "hidden-data-unsupported-format", path);
                continue;
            }
            try {
                WriteObject(new OfficeHiddenDataFileReport(OfficeDocument.InspectHiddenData(path), path, null));
            } catch (Exception error) when (error is not (ActionPreferenceStopException or PipelineStoppedException)) {
                ItemError(error, "hidden-data-inspection-failed", path);
            }
        }
    }
}

internal static class OfficeHiddenDataFiles {
    internal static Exception Unsupported(string path) => new PSNotSupportedException(
        "Hidden data is supported for Word, Excel, PowerPoint, Visio, OpenDocument, EPUB, PDF and JPEG, PNG, WebP, GIF and TIFF files: " + path);
}
