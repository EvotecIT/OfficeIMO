using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a reusable PDF layout or visual fidelity diagnostic produced while generating PDF output.
/// </summary>
public sealed class PdfLayoutDiagnostic {
    /// <summary>
    /// Creates a layout diagnostic without source bounds.
    /// </summary>
    public PdfLayoutDiagnostic(PdfLayoutDiagnosticKind kind, string source, string message)
        : this(kind, source, message, null, null, null, null) {
    }

    /// <summary>
    /// Creates a layout diagnostic with optional source bounds in points.
    /// </summary>
    public PdfLayoutDiagnostic(PdfLayoutDiagnosticKind kind, string source, string message, global::OfficeIMO.OfficeLength? x, global::OfficeIMO.OfficeLength? y, global::OfficeIMO.OfficeLength? width, global::OfficeIMO.OfficeLength? height) {
        Kind = kind;
        Source = source ?? string.Empty;
        Message = message ?? string.Empty;
        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    /// <summary>Diagnostic kind.</summary>
    public PdfLayoutDiagnosticKind Kind { get; }

    /// <summary>Source converter, block, or feature that emitted the diagnostic.</summary>
    public string Source { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Optional left coordinate in points.</summary>
    public global::OfficeIMO.OfficeLength? X { get; }

    /// <summary>Optional top coordinate in points.</summary>
    public global::OfficeIMO.OfficeLength? Y { get; }

    /// <summary>Optional width in points.</summary>
    public global::OfficeIMO.OfficeLength? Width { get; }

    /// <summary>Optional height in points.</summary>
    public global::OfficeIMO.OfficeLength? Height { get; }

    /// <summary>Whether source bounds were provided.</summary>
    public bool HasBounds => X.HasValue && Y.HasValue && Width.HasValue && Height.HasValue;

    /// <inheritdoc />
    public override string ToString() {
        if (!HasBounds) {
            return string.IsNullOrWhiteSpace(Source)
                ? Kind + ": " + Message
                : Kind + " [" + Source + "]: " + Message;
        }

        return string.Format(
            CultureInfo.InvariantCulture,
            "{0} [{1}] at {2:0.###},{3:0.###},{4:0.###},{5:0.###}: {6}",
            Kind,
            Source,
            X!.Value.ToPoints(),
            Y!.Value.ToPoints(),
            Width!.Value.ToPoints(),
            Height!.Value.ToPoints(),
            Message);
    }
}
