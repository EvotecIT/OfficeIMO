using System.Management.Automation;
using OfficeIMO.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Releases a live OpenDocument document without saving.</summary>
/// <example><summary>Release a document.</summary><code>Close-OfficeOpenDocument -Document $document</code></example>
[Cmdlet(VerbsCommon.Close, "OfficeOpenDocument", DefaultParameterSetName = "Document")]
[OutputType(typeof(void))]
public sealed class CloseOfficeOpenDocumentCommand : PSCmdlet {
    /// <summary>Live document to release.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public OdfDocument Document { get; set; } = null!;
    /// <inheritdoc />
    protected override void ProcessRecord() => Document.Dispose();
}
