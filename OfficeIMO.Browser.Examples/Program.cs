using System.Text;
using HtmlForgeX;
using OfficeIMO.Browser;

if (args.Length != 1) throw new ArgumentException("Pass the output directory for the two offline example pages.");
string output = Path.GetFullPath(args[0]);
Directory.CreateDirectory(output);
string script = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "current-view.js"));
foreach (bool bundle in new[] { false, true }) {
    using var document = new Document { LibraryMode = LibraryMode.Offline };
    document.Head.AddTitle("Current-view exports");
    document.Head.AddDefaultStyles();
    document.Body.Add(new HtmlTag("h1").Value("Current-view exports"));
    document.Body.Add(new HtmlTag("p").Value("Filter, sort or change columns, then export the displayed table. Everything runs offline."));
    document.Body.Add(new HtmlTag("label").Attribute("for", "filter").Value("Search name or site "));
    document.Body.Add(new HtmlTag("input").Attribute("type", "search").Id("filter"));
    foreach (var button in new[] { ("sort", "Reverse name order"), ("hide-site", "Hide/show Site"), ("move-site", "Move Site first/last"),
                 ("export-xlsx", "Export Excel"), ("export-csv", "Export CSV") })
        document.Body.Add(new HtmlTag("button").Attribute("type", "button").Id(button.Item1).Value(button.Item2));
    var table = new Table().AddHeaders("Name", "Site", "Latency (ms)");
    table.AddRow(new List<string> { "DC01", "Warsaw", "12.5" });
    table.AddRow(new List<string> { "DC02", "Łódź", "8" });
    table.AddRow(new List<string> { "DC03", "Warsaw", "21" });
    // Table renders the typed rows; the plain wrapper needs no third-party table runtime.
    document.Body.Add(new HtmlTag("table").Value(table));
    document.Body.Add(new HtmlTag("p").Id("status").Attribute("role", "status").Attribute("aria-live", "polite").Value("Ready"));
    if (bundle) {
        File.WriteAllText(Path.Combine(output, BrowserAssets.Script.HashedFileName), BrowserAssets.Script.Content, new UTF8Encoding(false));
        document.Head.AddJsLink(BrowserAssets.Script.HashedFileName);
    } else document.Head.AddJsInline(TrustedJavaScript.FromTrustedSource(BrowserAssets.Script.Content));
    document.Body.Add(new InlineScript(TrustedJavaScript.FromTrustedSource(script)));
    document.Save(Path.Combine(output, bundle ? "current-view-bundle.html" : "current-view.html"), openInBrowser: false);
}
Console.WriteLine("Wrote inline and bundled HtmlForgeX examples to " + output);
