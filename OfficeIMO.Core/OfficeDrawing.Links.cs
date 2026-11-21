namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    /// <summary>Rewrites link targets on an owned drawing, removing links whose callback returns null without changing paint.</summary>
    internal void RewriteLinks(Func<string, string?> rewrite, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(rewrite);
        for (int index = _elements.Count - 1; index >= 0; index--) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDrawingElement element = _elements[index];
            if (element is OfficeDrawingLink link) {
                string? target = rewrite(link.Uri);
                if (target == null) {
                    _elements.RemoveAt(index);
                    _behindContentElements.Remove(link);
                } else if (!string.Equals(target, link.Uri, StringComparison.Ordinal)) {
                    ReplaceElement(index, link, new OfficeDrawingLink(target, link.X, link.Y,
                        link.Width, link.Height, link.AlternativeText) { SourceElementIds = link.SourceElementIds });
                }
            } else if (element is OfficeDrawingGroup group) {
                group.InnerDrawing.RewriteLinks(rewrite, cancellationToken);
            } else if (element is OfficeDrawingEffectGroup effect) {
                effect.InnerDrawing.RewriteLinks(rewrite, cancellationToken);
            } else if (element is OfficeDrawingTilingPattern pattern) {
                pattern.InnerTile.RewriteLinks(rewrite, cancellationToken);
            }
        }
    }
}
