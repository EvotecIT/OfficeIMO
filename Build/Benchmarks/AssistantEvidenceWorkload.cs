using System;
using System.IO;
using System.Threading;
using OfficeIMO.AI;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

/// <summary>Validated document workloads; PowerForge owns warmup, timing, ordering and evidence output.</summary>
public sealed class AssistantEvidenceWorkload {
    private readonly byte[] _bytes;
    private readonly OfficeDocumentReader _reader;
    private readonly OfficeAiDocument _prepared;
    public OfficeAiDocument Result { get; private set; }
    public OfficeAiEvidenceReadiness Readiness { get; private set; }
    public long AllocatedBytes { get; private set; }
    public int PageCount { get; }

    public AssistantEvidenceWorkload(int pages) {
        PageCount = pages;
        _bytes = PdfDocument.Create(document => {
            for (int index = 1; index <= pages; index++) {
                int pageNumber = index;
                document.Page(page => page.Content(content => {
                    content.Text("Evidence workload page " + pageNumber);
                    for (int line = 0; line < 12; line++) content.Text("Invoice reference SAMPLE-2026. Quantity 12. Net total 42 EUR. Review the original document.");
                }));
            }
        }).ToBytes();
        _reader = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();
        _prepared = Read();
        Result = _prepared;
        Readiness = OfficeAiEvidenceReadiness.Inspect(Result);
        Validate();
    }

    public void PrepareAgain() {
        long allocated = GC.GetTotalAllocatedBytes(true);
        Result = Read();
        Readiness = OfficeAiEvidenceReadiness.Inspect(Result);
        AllocatedBytes = GC.GetTotalAllocatedBytes(true) - allocated;
    }

    public void ReusePrepared() {
        long allocated = GC.GetTotalAllocatedBytes(true);
        Result = _prepared;
        Readiness = OfficeAiEvidenceReadiness.Inspect(Result);
        AllocatedBytes = GC.GetTotalAllocatedBytes(true) - allocated;
    }

    public void Validate() {
        if (Result.SnapshotHash != _prepared.SnapshotHash || Result.SourceHash != _prepared.SourceHash
            || Readiness.Pages.Count != PageCount || !Readiness.HasText || Readiness.PagesWithoutText.Count != 0)
            throw new InvalidDataException("The measured lane did not preserve the document evidence.");
    }

    private OfficeAiDocument Read() {
        using (var stream = new MemoryStream(_bytes, false))
            return OfficeAiDocument.ReadAsync(_reader, stream, "benchmark.pdf", cancellationToken: CancellationToken.None).GetAwaiter().GetResult();
    }
}
