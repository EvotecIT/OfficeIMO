namespace OfficeIMO.Pdf;

/// <summary>Bounded PDF read resource whose configured limit was exceeded.</summary>
public enum PdfReadLimitKind {
    /// <summary>Total PDF input byte count.</summary>
    InputBytes,

    /// <summary>Indirect object declaration or resolved-object count.</summary>
    IndirectObjects,

    /// <summary>Raw bytes in one PDF stream before decoding.</summary>
    RawStreamBytes,

    /// <summary>Bytes produced while decoding one filtered PDF stream.</summary>
    DecodedStreamBytes,

    /// <summary>Wall-clock time spent in the core object parsing pass.</summary>
    ObjectParsingTime
}
