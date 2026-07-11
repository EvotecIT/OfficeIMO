namespace OfficeIMO.Pdf;

/// <summary>Read-only diagnostic projection of PDF objects, revisions, pages, resources, and content operators.</summary>
public sealed class PdfDebuggerReport {
    internal PdfDebuggerReport(
        IReadOnlyList<PdfDebugObject> objects,
        IReadOnlyList<PdfDocumentRevisionInfo> revisions,
        IReadOnlyList<PdfDebugPage> pages,
        PdfRepairReport repairReport) {
        Objects = objects;
        Revisions = revisions;
        Pages = pages;
        RepairReport = repairReport;
    }

    /// <summary>Indirect object summaries in object-number order.</summary>
    public IReadOnlyList<PdfDebugObject> Objects { get; }

    /// <summary>Incremental revision chain.</summary>
    public IReadOnlyList<PdfDocumentRevisionInfo> Revisions { get; }

    /// <summary>Page resource and content summaries.</summary>
    public IReadOnlyList<PdfDebugPage> Pages { get; }

    /// <summary>Explicit repairs applied while parsing in lenient mode.</summary>
    public PdfRepairReport RepairReport { get; }

    /// <summary>Renders the projection as stable human-readable text.</summary>
    public string ToText() => PdfDebuggerTextFormatter.Format(this);
}

/// <summary>Debugger summary for one indirect object.</summary>
public sealed class PdfDebugObject {
    internal PdfDebugObject(int objectNumber, int generation, string kind, IReadOnlyList<string> dictionaryKeys, IReadOnlyList<int> references, bool reachable, long? streamLength, long? decodedStreamLength, string? decodedStreamPreview) {
        ObjectNumber = objectNumber;
        Generation = generation;
        Kind = kind;
        DictionaryKeys = dictionaryKeys;
        References = references;
        Reachable = reachable;
        StreamLength = streamLength;
        DecodedStreamLength = decodedStreamLength;
        DecodedStreamPreview = decodedStreamPreview;
    }

    /// <summary>Indirect object number.</summary>
    public int ObjectNumber { get; }
    /// <summary>Indirect object generation.</summary>
    public int Generation { get; }
    /// <summary>Stable object kind derived from its value, type, and subtype.</summary>
    public string Kind { get; }
    /// <summary>Sorted direct dictionary keys.</summary>
    public IReadOnlyList<string> DictionaryKeys { get; }
    /// <summary>Sorted indirect object references found in the value.</summary>
    public IReadOnlyList<int> References { get; }
    /// <summary>Whether the object is reachable from the active catalog, Info dictionary, or encryption dictionary.</summary>
    public bool Reachable { get; }
    /// <summary>Stored stream length, when this is a stream object.</summary>
    public long? StreamLength { get; }
    /// <summary>Decoded stream length when decoding completed within the configured preview bound.</summary>
    public long? DecodedStreamLength { get; }
    /// <summary>Optional bounded decoded stream preview.</summary>
    public string? DecodedStreamPreview { get; }
}

/// <summary>Debugger summary for one page.</summary>
public sealed class PdfDebugPage {
    internal PdfDebugPage(int pageNumber, int objectNumber, IReadOnlyList<string> resourceCategories, IReadOnlyList<int> contentObjectNumbers, IReadOnlyList<string> contentOperators, bool contentOperatorsTruncated) {
        PageNumber = pageNumber;
        ObjectNumber = objectNumber;
        ResourceCategories = resourceCategories;
        ContentObjectNumbers = contentObjectNumbers;
        ContentOperators = contentOperators;
        ContentOperatorsTruncated = contentOperatorsTruncated;
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }
    /// <summary>Underlying page object number.</summary>
    public int ObjectNumber { get; }
    /// <summary>Sorted direct resource dictionary categories, such as Font or XObject.</summary>
    public IReadOnlyList<string> ResourceCategories { get; }
    /// <summary>Referenced content stream object numbers.</summary>
    public IReadOnlyList<int> ContentObjectNumbers { get; }
    /// <summary>Content operators in source order.</summary>
    public IReadOnlyList<string> ContentOperators { get; }
    /// <summary>Whether operator collection reached its configured limit.</summary>
    public bool ContentOperatorsTruncated { get; }
}
