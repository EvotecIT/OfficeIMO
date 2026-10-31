using System.Management.Automation;
using OfficeIMO.Epub;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Completes an EPUB publication's lifecycle without saving.</summary>
/// <example><summary>Close a live publication after an explicit save.</summary>
/// <code>Close-OfficeEpub -Document $book</code></example>
[Cmdlet(VerbsCommon.Close, "OfficeEpub")]
public sealed class CloseOfficeEpubCommand : PSCmdlet {
    /// <summary>Live publication to close without saving.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public EpubPublication Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() => Document.Dispose();
}
