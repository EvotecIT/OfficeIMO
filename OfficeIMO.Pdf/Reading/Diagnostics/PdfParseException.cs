namespace OfficeIMO.Pdf;

/// <summary>Strict-mode failure for a structurally defective PDF.</summary>
public sealed class PdfParseException : IOException {
    internal PdfParseException(string code, string message, int? objectNumber)
        : base(message) {
        Code = code;
        ObjectNumber = objectNumber;
    }

    /// <summary>Stable machine-readable defect code.</summary>
    public string Code { get; }

    /// <summary>Affected indirect object number, when applicable.</summary>
    public int? ObjectNumber { get; }
}
