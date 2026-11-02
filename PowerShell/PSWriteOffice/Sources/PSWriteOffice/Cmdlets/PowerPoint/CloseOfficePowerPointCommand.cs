using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Closes a live presentation without saving it.</summary>
/// <para>Releases the native document and writes no output. Save changes explicitly with Save-OfficePowerPoint before closing. Documents configured for SaveOnDispose must be reopened with explicit persistence.</para>
/// <example>
///   <summary>Save and close a deck.</summary>
///   <code>$deck = Open-OfficePowerPoint -Path ./Review.pptx
/// Save-OfficePowerPoint -Document $deck
/// Close-OfficePowerPoint -Document $deck</code>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficePowerPoint", DefaultParameterSetName = "Document")]
public sealed class CloseOfficePowerPointCommand : OfficeDocumentCmdlet<PowerPointDocument> {
    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointDocument document = GetDocument();
        try { PowerPointDocumentService.CloseDocument(document, save: false, show: false); }
        catch (PipelineStoppedException) { throw; }
        catch (System.Exception error) { ItemError(error, "document-close-failed", document); }
    }
}
