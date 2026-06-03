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
        Guard.NotNullOrWhiteSpace(conformance, nameof(conformance));

        string normalizedConformance = conformance.Trim().ToUpperInvariant();
        ValidateConformance(normalizedConformance, nameof(conformance));

        Part = part;
        Conformance = normalizedConformance;
    }

    /// <summary>PDF/A standard part, currently 2 or 3.</summary>
    public int Part { get; }

    /// <summary>PDF/A conformance level: A, B, or U.</summary>
    public string Conformance { get; }

    internal PdfAIdentification Clone() => new PdfAIdentification(Part, Conformance);

    private static void ValidatePart(int part, string paramName) {
        if (part != 2 && part != 3) {
            throw new ArgumentOutOfRangeException(paramName, "PDF/A identification supports part 2 or 3.");
        }
    }

    private static void ValidateConformance(string conformance, string paramName) {
        if (conformance != "A" && conformance != "B" && conformance != "U") {
            throw new ArgumentException("PDF/A conformance must be A, B, or U.", paramName);
        }
    }
}
