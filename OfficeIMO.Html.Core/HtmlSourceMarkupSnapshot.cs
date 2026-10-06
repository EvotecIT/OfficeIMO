namespace OfficeIMO.Html.Dom;

// Immutable lexical provenance. Consumers must compare it with the current subtree before
// identifying it as the source of that subtree; mutations deliberately retain no such guarantee.
internal sealed class HtmlSourceMarkupSnapshot {
    internal HtmlSourceMarkupSnapshot(string markup) => Markup = markup;
    internal string Markup { get; }
}
