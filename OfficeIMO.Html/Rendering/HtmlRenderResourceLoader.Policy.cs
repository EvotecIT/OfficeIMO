namespace OfficeIMO.Html;

internal static partial class HtmlRenderResourceLoader {
    // A manifest may have been planned under a more permissive policy. Only this
    // operation's snapshot can authorize dispatch or admission of returned content.
    private static bool ApproveResourceUri(HtmlResourceReference reference, Uri uri,
        HtmlUrlPolicy policy, HtmlDiagnosticReport diagnostics, bool finalUri) {
        string approvedSource = HtmlUrlPolicyEvaluator.ResolveUrl(uri.AbsoluteUri, null, policy);
        if (Uri.TryCreate(approvedSource, UriKind.Absolute, out Uri? approvedUri)
            && approvedUri.Equals(uri)) return true;

        diagnostics.Add(ComponentName, GetPolicyRejectionCode(reference.Kind),
            finalUri
                ? "A resolver-reported final resource URI was rejected by the configured URL policy."
                : "A resource request was rejected by the operation's URL policy before resolution.",
            HtmlDiagnosticSeverity.Warning, reference.Source, uri.AbsoluteUri, OfficeConversionLossKind.Omission);
        return false;
    }

    private static string GetPolicyRejectionCode(HtmlResourceKind kind) => kind switch {
        HtmlResourceKind.Image => "ImageResourceRejectedByPolicy",
        HtmlResourceKind.Stylesheet => "StylesheetResourceRejectedByPolicy",
        HtmlResourceKind.Hyperlink => "HyperlinkRejectedByPolicy",
        HtmlResourceKind.Script => "ScriptResourceRejectedByPolicy",
        HtmlResourceKind.Media => "MediaResourceRejectedByPolicy",
        HtmlResourceKind.Font => "FontResourceRejectedByPolicy",
        _ => "HtmlResourceRejectedByPolicy"
    };
}
