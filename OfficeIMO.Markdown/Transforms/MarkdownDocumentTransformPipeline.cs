namespace OfficeIMO.Markdown;

/// <summary>
/// Applies ordered post-parse transforms to typed markdown documents.
/// </summary>
public static class MarkdownDocumentTransformPipeline {
    /// <summary>
    /// Applies the supplied transforms in order.
    /// </summary>
    /// <param name="document">Document to transform.</param>
    /// <param name="transforms">Ordered transforms.</param>
    /// <param name="context">Execution context.</param>
    /// <returns>The final transformed document.</returns>
    public static MarkdownDoc Apply(
        MarkdownDoc document,
        IEnumerable<IMarkdownDocumentTransform>? transforms,
        MarkdownDocumentTransformContext context) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        if (transforms == null) {
            return document;
        }

        var current = document;
        foreach (var transform in transforms) {
            if (transform == null) {
                continue;
            }

            current = transform.Transform(current, context) ?? current;
        }

        return current;
    }
}
