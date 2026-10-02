using System.Runtime.CompilerServices;
using AngleSharp.Dom;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html;

internal static class NativeSourceMarkup {
    private static readonly ConditionalWeakTable<IElement, HtmlSourceMarkupSnapshot> Sources = new();
    internal static HtmlSourceMarkupSnapshot? Get(IElement element) => Sources.TryGetValue(element, out var source) ? source : null;
    internal static void Attach(IElement element, HtmlSourceMarkupSnapshot? source) {
        if (source != null) Sources.Add(element, source);
    }
}
