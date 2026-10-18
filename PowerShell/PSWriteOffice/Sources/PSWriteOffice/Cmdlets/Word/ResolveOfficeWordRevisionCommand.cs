using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Accepts or rejects filtered Word revisions and returns an operation report.</summary>
/// <example>
///   <summary>Accept revisions by one author into a new document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$filter = New-OfficeWordRevisionFilter -Author 'Reviewer' -InContentControl
/// Resolve-OfficeWordRevision -Path .\Draft.docx -OutputPath .\Accepted.docx -Action Accept -Filter $filter</code>
///   <para>Applies only matching revisions, saves the result, and returns the matched revision report.</para>
/// </example>
[Cmdlet(VerbsDiagnostic.Resolve, "OfficeWordRevision", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(WordRevisionOperationReport))]
public sealed class ResolveOfficeWordRevisionCommand : PSCmdlet
{
    /// <summary>Path to the Word document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open Word document instance.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Accept or reject matching revisions.</summary>
    [Parameter(Mandatory = true)]
    public WordRevisionOperationKind Action { get; set; }

    /// <summary>Optional author, id, type, date, location, part, or container filter.</summary>
    [Parameter]
    public WordRevisionFilter? Filter { get; set; }

    /// <summary>Output document path. Required for path input.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Path")]
    public string? OutputPath { get; set; }

    /// <summary>Return the mutated document after the operation.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WordDocument? owned = null;
        try
        {
            var document = Document;
            string? output = null;
            if (ParameterSetName == "Path")
            {
                var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath!);
                if (!ShouldProcess(output, $"{Action} matching Word revisions")) return;
                owned = WordDocumentService.LoadDocument(input, readOnly: false, autoSave: false);
                document = owned;
            }
            else if (!ShouldProcess("WordDocument", $"{Action} matching Word revisions"))
            {
                return;
            }

            var filter = Filter ?? WordRevisionFilter.All();
            var result = Action == WordRevisionOperationKind.Accept
                ? document.AcceptRevisions(filter)
                : document.RejectRevisions(filter);
            if (output != null)
            {
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
                document.Save(output);
            }
            WriteObject(result);
            if (PassThru.IsPresent) WriteObject(document);
        }
        finally
        {
            owned?.Dispose();
        }
    }
}
