namespace OfficeIMO.Html;

/// <summary>
/// Evidence captured after saving, reopening, and exporting a native artifact back to HTML.
/// The scorer compares the reloaded export with the original source instead of trusting a save call alone.
/// </summary>
public sealed class HtmlArtifactReloadEvidence {
    private HtmlArtifactReloadEvidence(string artifactKind, bool reloadSucceeded, string? reloadedHtml, string? detail) {
        ArtifactKind = artifactKind;
        ReloadSucceeded = reloadSucceeded;
        ReloadedHtml = reloadedHtml;
        Detail = detail;
    }

    /// <summary>Artifact kind, such as DOCX, XLSX, or PPTX.</summary>
    public string ArtifactKind { get; }

    /// <summary>Whether the artifact was successfully reopened.</summary>
    public bool ReloadSucceeded { get; }

    /// <summary>HTML exported from the reopened artifact when reload succeeded.</summary>
    public string? ReloadedHtml { get; }

    /// <summary>Optional caller diagnostic for failed reload evidence.</summary>
    public string? Detail { get; }

    /// <summary>Creates successful reload evidence from HTML exported after reopening the artifact.</summary>
    public static HtmlArtifactReloadEvidence Succeeded(string artifactKind, string reloadedHtml) {
        if (string.IsNullOrWhiteSpace(artifactKind)) throw new ArgumentNullException(nameof(artifactKind));
        if (string.IsNullOrWhiteSpace(reloadedHtml)) throw new ArgumentNullException(nameof(reloadedHtml));
        return new HtmlArtifactReloadEvidence(artifactKind.Trim(), true, reloadedHtml, null);
    }

    /// <summary>Creates explicit failed reload evidence.</summary>
    public static HtmlArtifactReloadEvidence Failed(string artifactKind, string? detail = null) {
        if (string.IsNullOrWhiteSpace(artifactKind)) throw new ArgumentNullException(nameof(artifactKind));
        return new HtmlArtifactReloadEvidence(artifactKind.Trim(), false, null, detail);
    }
}
