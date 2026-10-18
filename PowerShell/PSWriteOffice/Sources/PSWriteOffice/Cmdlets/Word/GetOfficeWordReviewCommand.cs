using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Inspects Word comments, threads, tracked revisions, and unsupported review metadata.</summary>
/// <example>
///   <summary>Inspect a document review state.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeWordReview -Path .\Draft.docx</code>
///   <para>Returns a structured WordReviewReport without changing the document.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeWordReview", DefaultParameterSetName = "Path")]
[OutputType(typeof(WordReviewReport))]
public sealed class GetOfficeWordReviewCommand : PSCmdlet
{
    /// <summary>Path to the Word document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open Word document instance.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public WordDocument Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WordDocument? owned = null;
        try
        {
            var document = Document;
            if (ParameterSetName == "Path")
            {
                owned = WordDocumentService.LoadDocument(
                    SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path), readOnly: true, autoSave: false);
                document = owned;
            }
            WriteObject(document.InspectReviewReport());
        }
        finally
        {
            owned?.Dispose();
        }
    }
}
