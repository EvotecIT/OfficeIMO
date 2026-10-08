namespace OfficeIMO.Html;

internal static partial class HtmlRenderResourceLoader {
    // A manifest may have been planned under a more permissive policy. Only this
    // operation's snapshot can authorize dispatch or admission of returned content.
    private static bool TryAuthorizeResourceRequest(HtmlResourceReference reference,
        HtmlUrlPolicy policy, HtmlDiagnosticReport diagnostics, out Uri uri) {
        bool includesPlanningTransform = reference.PlanningUrlTransform != null
            && policy.BeginsWithResolvedUrlTransform(reference.PlanningUrlTransform);
        string approvedSource = HtmlUrlPolicyEvaluator.ResolveUrl(
            includesPlanningTransform ? reference.Source : reference.ResolvedSource,
            includesPlanningTransform ? reference.ResolutionBaseUri : null,
            policy);
        // A rewrite can cross an intersected policy's admission boundary. Check
        // the resulting target too, but keep the first rewrite as the request
        // identity: applying a signer again would change the resource URL.
        if (Uri.TryCreate(approvedSource, UriKind.Absolute, out Uri? approvedUri)
            && HtmlUrlPolicyEvaluator.ResolveUrl(approvedUri.AbsoluteUri, null, policy).Length > 0) {
            uri = approvedUri;
            return true;
        }

        uri = null!;
        AddPolicyRejection(reference, reference.ResolvedSource, diagnostics, finalUri: false);
        return false;
    }

    // A different resolver-reported identity must itself be admitted without
    // rewriting: these returned bytes were fetched from that actual URI.
    private static bool ApproveFinalResourceUri(HtmlResourceReference reference, Uri uri,
        HtmlUrlPolicy policy, HtmlDiagnosticReport diagnostics) {
        string approvedSource = HtmlUrlPolicyEvaluator.ResolveUrl(uri.AbsoluteUri, null, policy);
        if (Uri.TryCreate(approvedSource, UriKind.Absolute, out Uri? approvedUri)
            && approvedUri.Equals(uri)) return true;

        AddPolicyRejection(reference, uri.AbsoluteUri, diagnostics, finalUri: true);
        return false;
    }

    private static void AddPolicyRejection(HtmlResourceReference reference, string uri,
        HtmlDiagnosticReport diagnostics, bool finalUri) {
        diagnostics.Add(ComponentName, GetPolicyRejectionCode(reference.Kind),
            finalUri
                ? "A resolver-reported final resource URI was rejected by the configured URL policy."
                : "A resource request was rejected by the operation's URL policy before resolution.",
            HtmlDiagnosticSeverity.Warning, reference.Source, uri, OfficeConversionLossKind.Omission);
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
