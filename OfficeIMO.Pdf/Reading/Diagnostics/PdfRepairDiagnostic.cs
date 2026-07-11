namespace OfficeIMO.Pdf;

/// <summary>One explicit structural recovery performed while reading a PDF.</summary>
public sealed class PdfRepairDiagnostic {
    internal PdfRepairDiagnostic(string code, string message, int? objectNumber) {
        Code = code;
        Message = message;
        ObjectNumber = objectNumber;
    }

    /// <summary>Stable machine-readable recovery code.</summary>
    public string Code { get; }

    /// <summary>Human-readable recovery explanation.</summary>
    public string Message { get; }

    /// <summary>Affected indirect object number, when applicable.</summary>
    public int? ObjectNumber { get; }
}
