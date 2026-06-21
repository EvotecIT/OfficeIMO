namespace OfficeIMO.Html;

/// <summary>
/// Creates URL policies for non-hyperlink resource references.
/// </summary>
internal static class HtmlResourceUrlPolicy {
    /// <summary>
    /// Clones the supplied policy and applies resource-only restrictions shared by normalization and resource planning.
    /// </summary>
    internal static HtmlUrlPolicy Create(HtmlUrlPolicy? policy) {
        HtmlUrlPolicy resourcePolicy = (policy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone();
        resourcePolicy.AllowMailtoUrls = false;
        return resourcePolicy;
    }
}
