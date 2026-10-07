using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookOnixMessage {
    private sealed record OnixMarketSelection(string[] Replace, string[] Remove) {
        internal bool HasChanges => Replace.Length + Remove.Length != 0;
    }

    private static OnixMarketSelection ReadMarkets(BookOnixBlockUpdate update) {
        ArgumentNullException.ThrowIfNull(update.ReplaceMarketReferences);
        ArgumentNullException.ThrowIfNull(update.RemoveMarketReferences);
        if (update.ReplaceMarketReferences.Count > 32 || update.RemoveMarketReferences.Count > 32 ||
            update.ReplaceMarketReferences.Count + update.RemoveMarketReferences.Count > 32)
            throw new ArgumentException("At most 32 market operations are supported.", nameof(update));
        var replace = update.ReplaceMarketReferences.ToArray(); var remove = update.RemoveMarketReferences.ToArray();
        var references = new HashSet<string>(StringComparer.Ordinal);
        foreach (string reference in replace.Concat(remove)) {
            BookProject.RequireOnixMarketReference(reference);
            if (!references.Add(reference)) throw new ArgumentException("Market operations must have distinct references, without replace/remove overlap.", nameof(update));
        }
        return new(replace, remove);
    }

    private static void AppendMarketChanges(XElement result, XElement source, OnixMarketSelection selection, CancellationToken token) {
        if (!selection.HasChanges) return;
        XNamespace ns = BookProject.OnixNamespace;
        var supplies = source.Elements(ns + "ProductSupply").ToArray();
        if (supplies.Any(supply => supply.Element(ns + "MarketReference") == null))
            throw new ArgumentException("Per-market changes require every source supply declaration to have a permanent market reference.");
        var byReference = supplies.ToDictionary(supply => supply.Element(ns + "MarketReference")!.Value, StringComparer.Ordinal);
        foreach (string reference in selection.Replace) {
            token.ThrowIfCancellationRequested();
            if (!byReference.TryGetValue(reference, out var supply))
                throw new ArgumentException("Selected replacement market is absent: " + reference + ".");
            result.Add(new XElement(supply));
        }
        foreach (string reference in selection.Remove) {
            token.ThrowIfCancellationRequested();
            result.Add(new XElement(ns + "ProductSupply", new XElement(ns + "MarketReference", reference)));
        }
    }
}
