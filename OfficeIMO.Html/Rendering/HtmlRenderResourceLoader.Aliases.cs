namespace OfficeIMO.Html;

public sealed partial class HtmlResourceSession {
    // Alias admission reuses accepted bytes and MIME checks without reserving
    // another request or counting the canonical resource again.
    internal void AliasAcceptedRequest(HtmlResourceReference reference, Uri requestUri) {
        if (_resources.TryGetValue(requestUri.AbsoluteUri, out HtmlResolvedResource? resource)) {
            TryAccept(reference, resource, out _, out _, requestUri);
        }
    }
}

internal static partial class HtmlRenderResourceLoader {
    private static void AliasAcceptedRequests(HtmlResourceSession session,
        Dictionary<string, List<HtmlResourceReference>> requestAliases, HtmlResourceKind kind, Uri requestUri) {
        if (!requestAliases.TryGetValue(GetSeenKey(kind, requestUri.AbsoluteUri), out List<HtmlResourceReference>? aliases)) return;
        foreach (HtmlResourceReference alias in aliases) session.AliasAcceptedRequest(alias, requestUri);
    }
}
