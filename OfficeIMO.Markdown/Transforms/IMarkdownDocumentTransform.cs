namespace OfficeIMO.Markdown;

/// <summary>
/// Post-parse document transform applied to a typed <see cref="MarkdownDoc"/>.
/// </summary>
public interface IMarkdownDocumentTransform {
    /// <summary>
    /// Applies the transform to the supplied markdown document.
    /// Implementations may mutate <paramref name="document"/> in place or return a replacement instance.
    /// </summary>
    /// <param name="document">Parsed markdown document.</param>
    /// <param name="context">Transform execution context.</param>
    /// <returns>The transformed document.</returns>
    MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context);
}
