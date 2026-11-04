namespace PSWriteOffice.Models.Reader;

/// <summary>Representation projected from one normalized document read.</summary>
public enum OfficeReadView {
    /// <summary>Normalized document envelope.</summary>
    Document,
    /// <summary>Plain text.</summary>
    Text,
    /// <summary>Portable Markdown.</summary>
    Markdown,
    /// <summary>Serialized document envelope.</summary>
    Json,
    /// <summary>Source chunks.</summary>
    Chunks,
    /// <summary>Hierarchical chunks.</summary>
    Hierarchy,
    /// <summary>Structured extraction.</summary>
    Structure,
    /// <summary>Structured tables or in-memory export bundles.</summary>
    Tables,
    /// <summary>Visual models or in-memory export bundles.</summary>
    Visuals,
    /// <summary>Inert asset descriptors.</summary>
    Assets,
    /// <summary>Batch ingestion status and optional chunks.</summary>
    Ingest,
    /// <summary>OCR review evidence for an original image or PDF.</summary>
    Ocr
}
