using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Updates metadata on a generated PDF, an open editing block, or file copies.</summary>
/// <para>Omitted values retain existing metadata; an empty string clears a value. Incremental appends a native metadata revision. File outputs are atomic; document edits return the updated native document with PassThru.</para>
/// <example><summary>Set metadata while generating a report.</summary>
/// <code>New-OfficePdf -Path ./Report.pdf { PdfMetadata -Title 'Service review'; PdfHeading 'Service review' }</code></example>
/// <example><summary>Update a loaded document and save a copy.</summary>
/// <code>Open-OfficePdf ./Report.pdf -OutputPath ./Reviewed.pdf { PdfMetadata -Title 'Reviewed report' }</code></example>
[Cmdlet(VerbsCommon.Set, "OfficePdfMetadata", DefaultParameterSetName = "Context", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument), typeof(FileInfo))]
[Alias("PdfMetadata") ]
public sealed class SetOfficePdfMetadataCommand : PdfMutationCmdlet {
    /// <summary>Explicit caller-owned document for inspection or editing.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public new PdfDocument? Document { get => base.Document; set => base.Document = value; }
    /// <summary>Document title; an empty string clears it.</summary>
    [Parameter(ParameterSetName = "Context"), Parameter(ParameterSetName = "Document"), Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "PathDirectory"), Parameter(ParameterSetName = "LiteralPath"), Parameter(ParameterSetName = "LiteralPathDirectory")]
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
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "PathDirectory"), Parameter(ParameterSetName = "LiteralPath"), Parameter(ParameterSetName = "LiteralPathDirectory")]
    public SwitchParameter Incremental { get; set; }
    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        if (ParameterSetName != "Context") return base.ProcessRecordAsync();
        if (ShouldProcess("PDF composition metadata", Operation)) {
            var current = OfficeDocumentContext.Find<PdfDocument>();
            if (current?.Pipeline.Input != null) {
                var updated = Apply(current, CancelToken);
                if (!OfficeDocumentContext.Replace(current, updated)) {
                    updated.Dispose();
                    throw new PSInvalidOperationException("Use -Document -PassThru to retain an edited PDF outside an editing block.");
                }
            } else PdfDslContext.Require(this).ConfigureDocument(document => document.Meta(Title, Author, Subject, Keywords));
        }
        return Task.CompletedTask;
    }
    /// <inheritdoc />
    protected override string Operation => "Set PDF metadata";
    /// <inheritdoc />
    protected override PdfDocument Apply(PdfDocument document, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document.Pipeline.Input == null && !Incremental)
            return document.Meta(Title, Author, Subject, Keywords);
        return Incremental
            ? document.AppendMetadataRevision(cancellationToken, Title, Author, Subject, Keywords)
            : document.UpdateProperties(cancellationToken, Title, Author, Subject, Keywords);
    }
}
