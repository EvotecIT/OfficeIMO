using System;

namespace OfficeIMO.Reader;

/// <summary>
/// Describes the success or failure of one path in a resilient multi-document read.
/// </summary>
public sealed class ReaderDocumentReadOutcome {
    internal ReaderDocumentReadOutcome(
        int index,
        string path,
        OfficeDocumentReadResult? document,
        Exception? error) {
        Index = index;
        Path = path;
        Document = document;
        Error = error;
    }

    /// <summary>Zero-based input position assigned after path expansion.</summary>
    public int Index { get; }

    /// <summary>Resolved source path.</summary>
    public string Path { get; }

    /// <summary>Document result when reading succeeded.</summary>
    public OfficeDocumentReadResult? Document { get; }

    /// <summary>Original read exception when reading failed.</summary>
    public Exception? Error { get; }

    /// <summary>True when the document was read successfully.</summary>
    public bool Succeeded => Document != null && Error == null;
}
