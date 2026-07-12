namespace OfficeIMO.Pdf;

/// <summary>Explicit policy for removing active content and embedded payloads from a PDF.</summary>
public sealed class PdfSanitizationOptions {
    /// <summary>Action types that may remain. Values are PDF action names without a leading slash.</summary>
    public ISet<string> AllowedActionTypes { get; } = new HashSet<string>(StringComparer.Ordinal);

    /// <summary>Absolute URI schemes that may remain. Relative URI targets are preserved.</summary>
    public ISet<string> AllowedUriSchemes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "http", "https", "mailto", "tel"
    };

    /// <summary>How embedded and associated files are removed. Defaults to removal without retaining payload bytes.</summary>
    public PdfEmbeddedFileSanitizationMode EmbeddedFiles { get; set; } = PdfEmbeddedFileSanitizationMode.Remove;

    /// <summary>When true, rich-media and payload-bearing annotations are removed from page annotation arrays.</summary>
    public bool RemoveRichMedia { get; set; } = true;

    internal bool IsActionAllowed(string actionType) => AllowedActionTypes.Contains(actionType);

    internal bool IsUriAllowed(string value) {
        if (!Uri.TryCreate(value, UriKind.RelativeOrAbsolute, out Uri? uri) || !uri.IsAbsoluteUri) {
            return true;
        }

        return AllowedUriSchemes.Contains(uri.Scheme);
    }
}
