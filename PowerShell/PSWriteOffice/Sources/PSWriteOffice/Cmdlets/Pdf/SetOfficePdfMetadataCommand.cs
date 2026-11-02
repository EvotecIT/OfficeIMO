using System.IO;
using System.Management.Automation;
using System.Threading;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Updates metadata on a generated PDF, an open editing block, or file copies.</summary>
/// <para>Omitted values retain existing metadata; an empty string clears a value. Incremental appends a native metadata revision. File outputs are atomic; document edits return the updated native document with PassThru.</para>
/// <example><summary>Set metadata while generating a report.</summary>
/// <code>New-OfficePdf -Path ./Report.pdf { PdfMetadata -Title 'Service review'; PdfHeading 'Service review' }</code></example>
/// <example><summary>Update a loaded document and save a copy.</summary>
/// <code>Open-OfficePdf ./Report.pdf -OutputPath ./Reviewed.pdf { PdfMetadata -Title 'Reviewed report' }</code></example>
[Cmdlet(VerbsCommon.Set, "OfficePdfMetadata", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[Alias("PdfMetadata")]
[OutputType(typeof(PdfDocument), typeof(FileInfo))]
public sealed class SetOfficePdfMetadataCommand : PdfMutationCmdlet {
    /// <summary>Document title; an empty string clears it.</summary>
    [Parameter]
    public string? Title { get; set; }
    /// <summary>Document author; an empty string clears it.</summary>
    [Parameter]
    public string? Author { get; set; }
    /// <summary>Document subject; an empty string clears it.</summary>
    [Parameter]
    public string? Subject { get; set; }
    /// <summary>Document keywords; an empty string clears them.</summary>
    [Parameter]
    public string? Keywords { get; set; }
    /// <summary>Append a metadata revision instead of rewriting existing PDF content.</summary>
    [Parameter]
    public SwitchParameter Incremental { get; set; }
    /// <inheritdoc />
    protected override string Operation => "Set PDF metadata";
    /// <inheritdoc />
    protected override PdfDocument Apply(PdfDocument document, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document.Pipeline.Input == null && !Incremental)
            return document.Meta(Title, Author, Subject, Keywords);
        return Incremental
            ? document.AppendMetadataRevision(cancellationToken, Title, Author, Subject, Keywords)
            : document.UpdateMetadata(cancellationToken, Title, Author, Subject, Keywords);
    }
}
