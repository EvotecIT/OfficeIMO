namespace OfficeIMO.Pdf;

/// <summary>Explicit policy for removing active content and embedded payloads from a PDF.</summary>
public sealed class PdfSanitizationOptions {
    private PdfSanitizationActionKind? _actionKindsToRemove;

    /// <summary>Cancellation observed between inventory, object-graph rewrite, and verification stages.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }

    /// <summary>Optional maximum byte count for the sanitized full-rewrite artifact.</summary>
    public long? MaximumOutputBytes { get; set; }

    /// <summary>Action types that may remain. Values are PDF action names without a leading slash.</summary>
    public ISet<string> AllowedActionTypes { get; } = new HashSet<string>(StringComparer.Ordinal);

    /// <summary>
    /// Exact action kinds to remove. When null, the existing default policy removes known active-content
    /// actions and only URI targets whose schemes are not allowed. When set, unselected action kinds are
    /// preserved and selecting <see cref="PdfSanitizationActionKind.Uri"/> removes every URI target.
    /// </summary>
    public PdfSanitizationActionKind? ActionKindsToRemove {
        get => _actionKindsToRemove;
        set {
            if (value.HasValue && (value.Value & ~PdfSanitizationActionKind.All) != 0) {
                throw new ArgumentOutOfRangeException(nameof(ActionKindsToRemove), value, "Unsupported PDF sanitization action kind.");
            }
            _actionKindsToRemove = value;
        }
    }

    /// <summary>Absolute URI schemes that may remain. Relative URI targets are preserved.</summary>
    public ISet<string> AllowedUriSchemes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "http", "https", "mailto", "tel"
    };

    /// <summary>How embedded and associated files are removed. Defaults to removal without retaining payload bytes.</summary>
    public PdfEmbeddedFileSanitizationMode EmbeddedFiles { get; set; } = PdfEmbeddedFileSanitizationMode.Remove;

    /// <summary>When true, rich-media and payload-bearing annotations are removed from page annotation arrays.</summary>
    public bool RemoveRichMedia { get; set; } = true;

    internal bool IsActionAllowed(string actionType) => AllowedActionTypes.Contains(actionType);

    internal bool ShouldRemoveAction(string actionType, string? uri = null) {
        PdfSanitizationActionKind kind = GetActionKind(actionType);
        if (ActionKindsToRemove.HasValue) {
            if (IsActionAllowed(actionType)) return false;
            return kind != PdfSanitizationActionKind.None && (ActionKindsToRemove.Value & kind) == kind;
        }
        if (kind == PdfSanitizationActionKind.Uri) return uri != null && !IsUriAllowed(uri);
        if (IsActionAllowed(actionType)) return false;
        return PdfActiveContentPolicy.IsUnsafeActionType(actionType);
    }

    internal bool ShouldRemoveCatalogUriBase(string value) => ActionKindsToRemove.HasValue
        ? (ActionKindsToRemove.Value & PdfSanitizationActionKind.Uri) != 0
        : !IsUriAllowed(value);

    internal static PdfSanitizationActionKind GetActionKind(string actionType) => actionType switch {
        "JavaScript" => PdfSanitizationActionKind.JavaScript,
        "URI" => PdfSanitizationActionKind.Uri,
        "Launch" => PdfSanitizationActionKind.Launch,
        "SubmitForm" => PdfSanitizationActionKind.SubmitForm,
        "GoToR" => PdfSanitizationActionKind.GoToR,
        "GoToE" => PdfSanitizationActionKind.GoToE,
        "ImportData" => PdfSanitizationActionKind.ImportData,
        "Movie" => PdfSanitizationActionKind.Movie,
        "Rendition" => PdfSanitizationActionKind.Rendition,
        "RichMedia" => PdfSanitizationActionKind.RichMedia,
        _ => PdfSanitizationActionKind.None
    };

    internal bool IsUriAllowed(string value) {
        if (!Uri.TryCreate(value, UriKind.RelativeOrAbsolute, out Uri? uri) || !uri.IsAbsoluteUri) {
            return true;
        }

        return AllowedUriSchemes.Contains(uri.Scheme);
    }
}
