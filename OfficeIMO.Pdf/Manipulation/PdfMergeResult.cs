namespace OfficeIMO.Pdf;

/// <summary>Observable inventory for one source supplied to a PDF merge.</summary>
public sealed class PdfMergeSourceInventory {
    internal PdfMergeSourceInventory(int sourceIndex, int pageCount, int outlineCount, int namedDestinationCount, int pageLabelCount, int formFieldCount, int attachmentCount) {
        SourceIndex = sourceIndex; PageCount = pageCount; OutlineCount = outlineCount; NamedDestinationCount = namedDestinationCount;
        PageLabelCount = pageLabelCount; FormFieldCount = formFieldCount; AttachmentCount = attachmentCount;
    }
    /// <summary>Zero-based source index.</summary>
    public int SourceIndex { get; }
    /// <summary>Pages imported from the source.</summary>
    public int PageCount { get; }
    /// <summary>Readable outline nodes, including descendants.</summary>
    public int OutlineCount { get; }
    /// <summary>Readable named destinations.</summary>
    public int NamedDestinationCount { get; }
    /// <summary>Readable page-label rules.</summary>
    public int PageLabelCount { get; }
    /// <summary>Readable terminal form fields.</summary>
    public int FormFieldCount { get; }
    /// <summary>Readable embedded or associated files.</summary>
    public int AttachmentCount { get; }
}

/// <summary>One applied merge-policy decision.</summary>
public sealed class PdfMergeDecision {
    internal PdfMergeDecision(string structure, PdfMergeStructureMode mode, string action, int importedCount = 0, int droppedCount = 0, IReadOnlyList<string>? renamedItems = null) {
        Structure = structure; Mode = mode; Action = action; ImportedCount = importedCount; DroppedCount = droppedCount;
        RenamedItems = renamedItems ?? Array.Empty<string>();
    }
    /// <summary>Stable structure name.</summary>
    public string Structure { get; }
    /// <summary>Requested policy mode.</summary>
    public PdfMergeStructureMode Mode { get; }
    /// <summary>Human-readable action actually applied.</summary>
    public string Action { get; }
    /// <summary>Number of non-primary items imported.</summary>
    public int ImportedCount { get; }
    /// <summary>Number of incoming items deliberately dropped.</summary>
    public int DroppedCount { get; }
    /// <summary>Deterministic old-to-new name mappings in <c>old -&gt; new</c> form.</summary>
    public IReadOnlyList<string> RenamedItems { get; }
}

/// <summary>Policy and readback evidence returned by a first-party PDF merge.</summary>
public sealed class PdfMergeReport {
    internal PdfMergeReport(IReadOnlyList<PdfMergeSourceInventory> sources, IReadOnlyList<PdfMergeDecision> decisions, int outputPageCount) {
        Sources = sources; Decisions = decisions; OutputPageCount = outputPageCount;
    }
    /// <summary>Per-source structure inventory captured before mutation.</summary>
    public IReadOnlyList<PdfMergeSourceInventory> Sources { get; }
    /// <summary>Policy decisions actually applied.</summary>
    public IReadOnlyList<PdfMergeDecision> Decisions { get; }
    /// <summary>Page count read back from the saved artifact.</summary>
    public int OutputPageCount { get; }
}

/// <summary>Merged PDF bytes plus the policy decisions and readback evidence.</summary>
public sealed class PdfMergeResult {
    private readonly byte[] _pdf;
    internal PdfMergeResult(byte[] pdf, PdfMergeReport report) { _pdf = (byte[])pdf.Clone(); Report = report; }
    /// <summary>Merge policy report.</summary>
    public PdfMergeReport Report { get; }
    /// <summary>Returns a defensive copy of the merged artifact.</summary>
    public byte[] ToBytes() => (byte[])_pdf.Clone();
    /// <summary>Opens the merged artifact through the OfficeIMO.Pdf document surface.</summary>
    public PdfDocument ToDocument() => PdfDocument.Open(_pdf);
}
