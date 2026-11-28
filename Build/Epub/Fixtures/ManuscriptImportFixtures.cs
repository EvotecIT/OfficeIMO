using OfficeIMO;
using OfficeIMO.Epub;
using OfficeIMO.Word;
using OfficeIMO.Workflows;
using System.Text;

internal static class ManuscriptImportFixtures {
    internal static async Task<(Dictionary<string, EpubPublication> Imports, Dictionary<string, byte[]> Catalog)> CreateAsync() {
        using var word = WordDocument.Create();
        word.Properties.Title = "Word manuscript";
        word.Properties.Creator = "OfficeIMO";
        word.AddParagraph("Opening").SetStyle(WordParagraphStyles.Heading1);
        word.AddParagraph("This chapter has a note.").AddFootNote("A note from the Word manuscript.");
        word.AddParagraph("Closing").SetStyle(WordParagraphStyles.Heading1);
        word.AddParagraph("The end of the manuscript.");

        const string markdown = """
            ---
            title: Markdown manuscript
            ---
            # Opening

            [Next chapter](#closing) and a note.[^source]

            # Closing

            The end of the manuscript.

            [^source]: A note from the Markdown manuscript.
            """;
        const string html = """
            <!DOCTYPE html>
            <html lang="en"><head><title>HTML manuscript</title></head><body>
            <h1>Opening</h1><p><a href="#closing">Next chapter</a></p>
            <h1 id="closing">Closing</h1><p>The end of the manuscript.</p>
            </body></html>
            """;

        var publications = new Dictionary<string, EpubPublication>(StringComparer.Ordinal);
        var catalog = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        OfficeWorkflow.RegisterConversions();
        await AddAsync("word", ".docx", OfficeDocumentFormat.Word, word.ToBytes());
        await AddAsync("markdown", ".md", OfficeDocumentFormat.Markdown, Encoding.UTF8.GetBytes(markdown));
        await AddAsync("html", ".html", OfficeDocumentFormat.Html, Encoding.UTF8.GetBytes(html));
        return (publications, catalog);

        async Task AddAsync(string format, string extension, OfficeDocumentFormat sourceFormat, byte[] bytes) {
            var options = new EpubManuscriptOptions { Identifier = "urn:officeimo:fixture:manuscript:" + format };
            EpubManuscriptResult result = await BookManuscriptImporter.ImportBytesAsync(bytes, extension, options);
            result.Report.ThrowIfLoss();
            EpubPublication publication = result.Publication;
            if (publication.Spine.Count != 2) throw new InvalidDataException("Expected two chapters in the " + format + " manuscript.");
            publication.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
                AccessModes = ["textual"],
                SufficientAccessModes = [new[] { "textual" }],
                Features = ["structuralNavigation", "tableOfContents"],
                Hazards = ["none"]
            });
            publications.Add("manuscript-" + format, publication);

            using var input = new MemoryStream(bytes, writable: false);
            using var output = new MemoryStream();
            OfficeConversionReport report = await OfficeConverter.ConvertAsync(input, sourceFormat, output, OfficeDocumentFormat.Epub,
                new OfficeConversionOptions {
                    RegisteredSettings = new OfficeWorkflowRouteSettings { EpubManuscript = options }
                });
            report.ThrowIfLoss();
            catalog.Add("catalog-manuscript-" + format, output.ToArray());
        }
    }
}
