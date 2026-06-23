namespace OfficeIMO.Pdf;

/// <summary>
/// PDF/A identification metadata emitted in the generated XMP packet.
/// </summary>
/// <remarks>
/// This model writes the ISO 19005 PDF/A identification fields only. It does not
/// certify that the generated PDF satisfies PDF/A validation requirements.
/// </remarks>
public sealed class PdfAIdentification {
    /// <summary>Creates PDF/A identification metadata for a supported PDF/A part and conformance level.</summary>
    public PdfAIdentification(int part, string conformance) {
        ValidatePart(part, nameof(part));
        Guard.NotNull(conformance, nameof(conformance));

        string normalizedConformance = conformance.Trim().ToUpperInvariant();
        ValidateConformance(part, normalizedConformance, nameof(conformance));

        Part = part;
        Conformance = normalizedConformance;
    }

    /// <summary>PDF/A standard part, currently 2, 3, or 4.</summary>
    public int Part { get; }

    /// <summary>PDF/A conformance level: A/B/U for PDF/A-2 and PDF/A-3, empty/E/F for PDF/A-4.</summary>
    public string Conformance { get; }

    /// <summary>Creates base PDF/A-4 identification metadata.</summary>
    public static PdfAIdentification PdfA4() => new PdfAIdentification(4, string.Empty);

    /// <summary>Creates PDF/A-4e identification metadata.</summary>
    public static PdfAIdentification PdfA4E() => new PdfAIdentification(4, "E");

    /// <summary>Creates PDF/A-4f identification metadata.</summary>
    public static PdfAIdentification PdfA4F() => new PdfAIdentification(4, "F");

    internal PdfAIdentification Clone() => new PdfAIdentification(Part, Conformance);

    private static void ValidatePart(int part, string paramName) {
        if (part != 2 && part != 3 && part != 4) {
            throw new ArgumentOutOfRangeException(paramName, "PDF/A identification supports part 2, 3, or 4.");
        }
    }

    private static void ValidateConformance(int part, string conformance, string paramName) {
        if (part == 4) {
            if (conformance.Length == 0 || conformance == "E" || conformance == "F") {
                return;
            }

            throw new ArgumentException("PDF/A-4 conformance must be empty, E, or F.", paramName);
        }

        if (conformance != "A" && conformance != "B" && conformance != "U") {
            throw new ArgumentException("PDF/A-2 and PDF/A-3 conformance must be A, B, or U.", paramName);
        }
    }
}
