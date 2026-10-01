using AngleSharp.Dom;

namespace OfficeIMO.Email;

/// <summary>Single owner of the email body projection's active-element and event-handler policy.</summary>
internal static class EmailHtmlActiveContent {
    private const string Selector = "script,iframe,object,embed,form,meta[http-equiv]";

    internal static (int Elements, int Attributes) Inspect(IDocument document) =>
        (document.QuerySelectorAll(Selector).Length,
            document.All.Sum(element => element.Attributes.Count(IsEventHandler)));

    internal static void Remove(IDocument document) {
        foreach (IElement element in document.QuerySelectorAll(Selector).ToArray()) element.Remove();
        foreach (IElement element in document.All) {
            foreach (IAttr attribute in element.Attributes.Where(IsEventHandler).ToArray())
                element.RemoveAttribute(attribute.Name);
        }
    }
    private static bool IsEventHandler(IAttr attribute) => attribute.Name.StartsWith("on", StringComparison.OrdinalIgnoreCase);
}
