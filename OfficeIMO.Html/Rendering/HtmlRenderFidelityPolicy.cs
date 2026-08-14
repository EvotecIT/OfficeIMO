namespace OfficeIMO.Html;

/// <summary>Controls whether the static renderer may return a document with diagnosed fidelity loss.</summary>
public enum HtmlRenderFidelityPolicy {
    /// <summary>Returns the rendered document together with all structured diagnostics.</summary>
    AllowDiagnosedLoss,

    /// <summary>Throws when rendering emits a warning or error instead of silently accepting a fallback.</summary>
    RequireNoLoss
}
