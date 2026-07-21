namespace OfficeIMO.Pdf;

/// <summary>Runtime evidence describing how one PDF save bounded completed payloads.</summary>
public sealed class PdfSerializationReport {
    internal PdfSerializationReport(
        int? pageCount,
        long bytesWritten,
        long pageContentMemoryLimitBytes,
        long objectBufferMemoryLimitBytes,
        long peakRetainedPageContentBytes,
        long peakRetainedObjectBytes,
        bool pageContentSpilled,
        bool objectBufferSpilled,
        bool finalArtifactBuffered,
        bool sourcePassthrough) {
        PageCount = pageCount;
        BytesWritten = bytesWritten;
        PageContentMemoryLimitBytes = pageContentMemoryLimitBytes;
        ObjectBufferMemoryLimitBytes = objectBufferMemoryLimitBytes;
        PeakRetainedPageContentBytes = peakRetainedPageContentBytes;
        PeakRetainedObjectBytes = peakRetainedObjectBytes;
        PageContentSpilled = pageContentSpilled;
        ObjectBufferSpilled = objectBufferSpilled;
        FinalArtifactBuffered = finalArtifactBuffered;
        SourcePassthrough = sourcePassthrough;
        IsForwardOnlyLayout = false;
    }

    /// <summary>Generated or inspected page count when known.</summary>
    public int? PageCount { get; }
    /// <summary>Complete PDF bytes written to the destination.</summary>
    public long BytesWritten { get; }
    /// <summary>Configured completed-page-content memory limit.</summary>
    public long PageContentMemoryLimitBytes { get; }
    /// <summary>Configured per-store completed-object memory limit.</summary>
    public long ObjectBufferMemoryLimitBytes { get; }
    /// <summary>Highest completed page-content byte count retained by the bounded store.</summary>
    public long PeakRetainedPageContentBytes { get; }
    /// <summary>Highest combined completed indirect-object byte count retained by simultaneously live assembly stores.</summary>
    public long PeakRetainedObjectBytes { get; }
    /// <summary>True when completed page/effect content spilled to temporary storage.</summary>
    public bool PageContentSpilled { get; }
    /// <summary>True when completed indirect objects spilled to temporary storage.</summary>
    public bool ObjectBufferSpilled { get; }
    /// <summary>True when the final artifact was intentionally materialized as a byte array.</summary>
    public bool FinalArtifactBuffered { get; }
    /// <summary>True when an opened artifact was copied without generated layout or object assembly.</summary>
    public bool SourcePassthrough { get; }

    /// <summary>
    /// False until layout, replay, and deterministic object allocation support genuinely forward-only output.
    /// </summary>
    public bool IsForwardOnlyLayout { get; }

    /// <summary>True when generated completed payloads were governed by explicit memory limits.</summary>
    public bool UsesBoundedCompletedPayloadStores => !SourcePassthrough;
}
