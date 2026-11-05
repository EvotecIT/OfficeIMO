using System.Management.Automation;
using OfficeIMO.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Releases a live Visio document without saving.</summary>
/// <example><summary>Release a document.</summary><code>Close-OfficeVisio -Document $document</code></example>
[Cmdlet(VerbsCommon.Close, "OfficeVisio", DefaultParameterSetName = "Document")]
[OutputType(typeof(void))]
public sealed class CloseOfficeVisioCommand : PSCmdlet {
    /// <summary>Live document to release.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public VisioDocument Document { get; set; } = null!;
    /// <inheritdoc />
    protected override void ProcessRecord() => Document.Dispose();
}
