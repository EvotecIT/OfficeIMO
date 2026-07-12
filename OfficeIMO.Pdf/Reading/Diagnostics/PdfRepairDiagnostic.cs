namespace OfficeIMO.Pdf;

/// <summary>Outcome of a parser or semantic repair diagnostic.</summary>
public enum PdfRepairDisposition {
    /// <summary>The lenient reader applied a deterministic recovery.</summary>
    Recovered,
    /// <summary>The issue was diagnosed but not changed because repair could alter semantics.</summary>
    DetectedOnly
}

/// <summary>One explicit structural recovery performed while reading a PDF.</summary>
public sealed class PdfRepairDiagnostic {
    internal PdfRepairDiagnostic(string code, string message, int? objectNumber, PdfRepairDisposition disposition = PdfRepairDisposition.Recovered) {
        Code = code;
        Message = message;
        ObjectNumber = objectNumber;
        Disposition = disposition;
    }

    /// <summary>Stable machine-readable recovery code.</summary>
    public string Code { get; }

    /// <summary>Human-readable recovery explanation.</summary>
    public string Message { get; }

    /// <summary>Affected indirect object number, when applicable.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Whether the reader recovered the issue or only reported it.</summary>
    public PdfRepairDisposition Disposition { get; }

    /// <summary>True when the lenient reader applied a deterministic recovery.</summary>
    public bool WasRecovered => Disposition == PdfRepairDisposition.Recovered;
}
