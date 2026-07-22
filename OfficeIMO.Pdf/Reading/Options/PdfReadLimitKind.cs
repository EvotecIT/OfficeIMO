namespace OfficeIMO.Pdf;

/// <summary>Bounded PDF read resource whose configured limit was exceeded.</summary>
public enum PdfReadLimitKind {
    /// <summary>Total PDF input byte count.</summary>
    InputBytes = 0,

    /// <summary>Indirect object declaration or resolved-object count.</summary>
    IndirectObjects = 1,

    /// <summary>Raw bytes in one PDF stream before decoding.</summary>
    RawStreamBytes = 2,

    /// <summary>Bytes produced while decoding one filtered PDF stream.</summary>
    DecodedStreamBytes = 3,

    /// <summary>Characters scanned for one object or dictionary.</summary>
    ObjectCharacters = 4,

    /// <summary>Tokens produced for one object or dictionary.</summary>
    ObjectTokens = 5,

    /// <summary>Nested array/dictionary parsing depth.</summary>
    ObjectNestingDepth = 6,

    /// <summary>Wall-clock time spent in the core object parsing pass.</summary>
    ObjectParsingTime = 7,

    /// <summary>Cross-reference revisions discovered in the input.</summary>
    Revisions = 8,

    /// <summary>Page-tree dictionaries traversed while discovering pages.</summary>
    PageTreeNodes = 9,

    /// <summary>Nested page-tree depth.</summary>
    PageTreeDepth = 10,

    /// <summary>Pages discovered in the document.</summary>
    Pages = 11,

    /// <summary>AcroForm field-tree nodes or terminal fields.</summary>
    FormFields = 12,

    /// <summary>Nested AcroForm field-tree depth.</summary>
    FormFieldDepth = 13,

    /// <summary>Annotations declared on one page.</summary>
    AnnotationsPerPage = 14,

    /// <summary>Operators parsed from one page or form content stream.</summary>
    ContentOperations = 15,

    /// <summary>Operand values and dictionary keys parsed from one page or form content stream.</summary>
    ContentOperands = 16,

    /// <summary>Nested lexical arrays/dictionaries or form XObjects traversed while parsing page content.</summary>
    ContentNestingDepth = 17,

    /// <summary>Pages requested in one managed render batch.</summary>
    RenderPages = 18,

    /// <summary>Output pixels requested for one managed rendered page.</summary>
    RenderPixels = 19,

    /// <summary>Selectable text regions requested for one page interaction map.</summary>
    InteractionRegions = 20,

    /// <summary>Intermediate artifacts produced by the PDF understanding pipeline.</summary>
    UnderstandingArtifacts = 21,

    /// <summary>Words, text, or diagnostics returned by an OCR provider.</summary>
    OcrArtifacts = 22,

    /// <summary>Rendered bytes retained by one managed render operation.</summary>
    RenderBytes = 23
}
