#region QuickStart
using OfficeIMO.Markdown;

var document = MarkdownDocument.Create()
    .FrontMatter(new { title = "Release notes", tags = new[] { "docs", "reporting" } })
    .H1("Release notes")
    .P("Typed Markdown builder, reader and HTML renderer for .NET.")
    .Ul(list => list
        .Item("Build Markdown with a fluent API")
        .Item("Parse Markdown into typed blocks and inlines")
        .Item("Render HTML fragments or full documents"));

File.WriteAllText("Release-notes.md", document.ToMarkdown());
File.WriteAllText("Release-notes.html", document.ToHtml(new MarkdownHtmlExportOptions { Kind = MarkdownHtmlKind.Document }));

MarkdownDocument parsed = MarkdownDocument.Parse(File.ReadAllText("Release-notes.md"));
Console.WriteLine($"{parsed.Blocks.Count} blocks");
#endregion
