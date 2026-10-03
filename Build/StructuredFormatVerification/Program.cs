using System.Text.Json;
using OfficeIMO.Adf;
using OfficeIMO.DocBook;

if (args.Length != 1) throw new ArgumentException("Supply an explicit output directory.");
string output = Path.GetFullPath(args[0]);
Directory.CreateDirectory(output);
var cases = new List<object>();
void AdfCase(string name, AdfDocument document, bool valid) {
    string file = name + ".json";
    File.WriteAllText(Path.Combine(output, file), document.ToJson());
    cases.Add(new { file, format = "adf", expectedValid = valid, boundedValid = document.Validate().IsValid });
}
void NativeAdf(string name, string nodes, bool valid) => AdfCase(name, AdfDocument.Parse("{\"version\":1,\"type\":\"doc\",\"content\":" + nodes + "}"), valid);
void DocBookCase(string name, DocBookDocument document, bool valid) {
    string file = name + ".docbook";
    File.WriteAllText(Path.Combine(output, file), document.ToDocBook());
    cases.Add(new { file, format = "docbook", expectedValid = valid, boundedValid = document.Validate().IsValid });
}

AdfCase("markdown-basic", AdfConverter.FromMarkdown("# Heading\n\n**strong** and *emphasis*\n\n- One\n- Two").Value, true);
AdfCase("markdown-image", AdfConverter.FromMarkdown("![Diagram](https://example.com/a.png)").Value, true);
AdfCase("markdown-nested-tasks", AdfConverter.FromMarkdown("- [x] outer\n  - [ ] inner\n  - [x] done\n- [ ] next").Value, true);
NativeAdf("native-empty-row", "[{\"type\":\"table\",\"content\":[{\"type\":\"tableRow\",\"content\":[]}]}]", true);
NativeAdf("native-alignment", "[{\"type\":\"paragraph\",\"marks\":[{\"type\":\"alignment\",\"attrs\":{\"align\":\"center\"}}],\"content\":[{\"type\":\"text\",\"text\":\"Visible\"}]}]", true);
NativeAdf("native-media-link", "[{\"type\":\"mediaSingle\",\"marks\":[{\"type\":\"link\",\"attrs\":{\"href\":\"https://example.com\"}}],\"content\":[{\"type\":\"media\",\"attrs\":{\"type\":\"external\",\"url\":\"https://example.com/a.png\"}}]}]", true);
var media = new AdfNode("media").SetAttribute("type", "external").SetAttribute("url", "https://example.com/a.png");
var caption = new AdfNode("caption") { Content = { AdfNode.TextNode("Caption", new[] { new AdfMark("strong") }) } };
AdfCase("native-media-caption", new AdfDocument(new[] { new AdfNode("mediaSingle") { Content = { media, caption } } }), true);
AdfCase("invalid-media-duplicate", new AdfDocument(new[] { new AdfNode("mediaSingle") { Content = { media, media } } }), false);
AdfCase("invalid-media-caption-order", new AdfDocument(new[] { new AdfNode("mediaSingle") { Content = { caption, media } } }), false);
AdfCase("invalid-caption-block", new AdfDocument(new[] { new AdfNode("mediaSingle") { Content = { media, new AdfNode("caption") { Content = { new AdfNode("paragraph") } } } } }), false);
var inlineExtension = new AdfNode("inlineExtension").SetAttribute("extensionType", "com.example").SetAttribute("extensionKey", "widget");
AdfCase("invalid-caption-inline-extension", new AdfDocument(new[] { new AdfNode("mediaSingle") { Content = { media, new AdfNode("caption") { Content = { inlineExtension } } } } }), false);
AdfCase("native-paragraph-inline-extension", new AdfDocument(new[] { new AdfNode("paragraph") { Content = { inlineExtension } } }), true);
NativeAdf("invalid-empty-list", "[{\"type\":\"bulletList\",\"content\":[]}]", false);
NativeAdf("invalid-empty-table", "[{\"type\":\"table\",\"content\":[]}]", false);
NativeAdf("invalid-empty-cell", "[{\"type\":\"table\",\"content\":[{\"type\":\"tableRow\",\"content\":[{\"type\":\"tableCell\",\"content\":[]}]}]}]", false);
NativeAdf("invalid-panel-type", "[{\"type\":\"panel\",\"content\":[{\"type\":\"paragraph\"}]}]", false);

var article = DocBookDocument.CreateArticle();
article.Title = "Guide";
article.AddSection("Section").AddParagraph("Body");
article.AddParagraph("Introduction");
DocBookCase("typed-body-after-section", article, true);
DocBookCase("invalid-raw-body-order", DocBookDocument.Parse("<article xmlns='http://docbook.org/ns/docbook' version='5.2'><title>Guide</title><section><title>Section</title><para>Body</para></section><para>Late</para></article>"), false);
string producer = Path.Combine(AppContext.BaseDirectory, "Fixtures", "pandoc-3.12-common-structure.docbook");
DocBookCase("independent-producer", DocBookDocument.Load(producer), true);
File.WriteAllText(Path.Combine(output, "manifest.json"), JsonSerializer.Serialize(cases, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Generated {cases.Count} structured-format conformance cases.");
