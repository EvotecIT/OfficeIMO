namespace OfficeIMO.Html;

/// <summary>
/// Index-based queries for retained visual lists. Scene walks visit every visual of a page several times, and
/// enumerating an <see cref="IReadOnlyList{T}"/> through its interface allocates an enumerator per list.
/// </summary>
internal static class HtmlRenderVisualLists {
    internal static bool Any(IReadOnlyList<HtmlRenderVisual>? visuals, Func<HtmlRenderVisual, bool> predicate) {
        if (visuals == null) return false;
        for (int index = 0; index < visuals.Count; index++) {
            if (predicate(visuals[index])) return true;
        }
        return false;
    }
}
