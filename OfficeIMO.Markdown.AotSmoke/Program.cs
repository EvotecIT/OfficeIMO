using OfficeIMO;
using OfficeIMO.Markdown;

string markdown = MarkdownDocument.Create()
    .H1("OfficeIMO NativeAOT")
    .P("Markdown rendering is executing inside the published native binary.")
    .ToMarkdown();

if (!markdown.Contains("# OfficeIMO NativeAOT", StringComparison.Ordinal)) {
    throw new InvalidOperationException("The Markdown renderer lost its heading.");
}

Console.WriteLine("PASS | Markdown composition and rendering");

using var document = MarkdownDocument.Create();
var options = new MarkdownTableOptions();
options.Order.Add(nameof(Stock.Count));
options.HeaderRenames[nameof(Stock.Count)] = "In stock";
var table = document.AddTable(new[] { new Stock { Product = "Paper", Count = 12 } }, options);
if (table.Headers[0] != "In stock" || table.Rows[0][0] != "12" || table.Rows[0][1] != "Paper") {
    throw new InvalidOperationException("Typed Markdown table projection lost its declared properties or fields.");
}

var dictionaries = document.AddTable(new[] { new StockDictionary() });
if (dictionaries.Headers[0] != "Count" || dictionaries.Rows[0][0] != "24") {
    throw new InvalidOperationException("Generic dictionary projection lost its keys or values.");
}

bool rejectedEmptyDictionary = false;
try {
    document.AddTable(Array.Empty<IReadOnlyDictionary<string, object?>>());
} catch (ArgumentException) {
    rejectedEmptyDictionary = true;
}
if (!rejectedEmptyDictionary) {
    throw new InvalidOperationException("An empty dictionary sequence must not expose dictionary metadata as table columns.");
}

Console.WriteLine("PASS | Typed Markdown tables, shared options and generic dictionary rows");

internal sealed class Stock {
    public string Product { get; set; } = "";
    public int Count;
}

// This fixture deliberately has no non-generic IDictionary implementation.
internal sealed class StockDictionary : IReadOnlyDictionary<string, int> {
    private readonly Dictionary<string, int> _values = new() { ["Count"] = 24 };
    public int this[string key] => _values[key];
    public IEnumerable<string> Keys => _values.Keys;
    public IEnumerable<int> Values => _values.Values;
    public int Count => _values.Count;
    public bool ContainsKey(string key) => _values.ContainsKey(key);
    public bool TryGetValue(string key, out int value) => _values.TryGetValue(key, out value);
    public IEnumerator<KeyValuePair<string, int>> GetEnumerator() => _values.GetEnumerator();
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
