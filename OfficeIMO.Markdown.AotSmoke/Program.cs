using OfficeIMO.Markdown;

string markdown = MarkdownDocument.Create()
    .H1("OfficeIMO NativeAOT")
    .P("Markdown rendering is executing inside the published native binary.")
    .ToMarkdown();

if (!markdown.Contains("# OfficeIMO NativeAOT", StringComparison.Ordinal)) {
    throw new InvalidOperationException("The Markdown renderer lost its heading.");
}

Console.WriteLine("PASS | Markdown fluent composition and rendering");
