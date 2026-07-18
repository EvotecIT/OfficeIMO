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

    /// <summary>Characters scanned for one object or dictionary.</summary>
    ObjectCharacters,

    /// <summary>Tokens produced for one object or dictionary.</summary>
    ObjectTokens,

    /// <summary>Nested array/dictionary parsing depth.</summary>
    ObjectNestingDepth,

    /// <summary>Wall-clock time spent in the core object parsing pass.</summary>
    ObjectParsingTime,

    /// <summary>Cross-reference revisions discovered in the input.</summary>
    Revisions,

    /// <summary>Page-tree dictionaries traversed while discovering pages.</summary>
    PageTreeNodes,

    /// <summary>Nested page-tree depth.</summary>
    PageTreeDepth,

    /// <summary>Pages discovered in the document.</summary>
    Pages,

    /// <summary>AcroForm field-tree nodes or terminal fields.</summary>
    FormFields,

    /// <summary>Nested AcroForm field-tree depth.</summary>
    FormFieldDepth,

    /// <summary>Annotations declared on one page.</summary>
    AnnotationsPerPage,

    /// <summary>Operators parsed from one page or form content stream.</summary>
    ContentOperations,

    /// <summary>Operand values and dictionary keys parsed from one page or form content stream.</summary>
    ContentOperands,

    /// <summary>Nested lexical arrays/dictionaries or form XObjects traversed while parsing page content.</summary>
    ContentNestingDepth,

    /// <summary>Pages requested in one managed render batch.</summary>
    RenderPages,

    /// <summary>Output pixels requested for one managed rendered page.</summary>
    RenderPixels,

    /// <summary>Selectable text regions requested for one page interaction map.</summary>
    InteractionRegions
}
