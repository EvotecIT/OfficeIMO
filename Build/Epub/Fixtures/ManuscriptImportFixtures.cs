using OfficeIMO.Epub;
using OfficeIMO.Word;
using OfficeIMO.Workflows;
using System.Text;

internal static class ManuscriptImportFixtures {
    internal static async Task<Dictionary<string, EpubPublication>> CreateAsync() {
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
        await AddAsync("word", ".docx", word.ToBytes());
        await AddAsync("markdown", ".md", Encoding.UTF8.GetBytes(markdown));
        await AddAsync("html", ".html", Encoding.UTF8.GetBytes(html));
        return publications;

        async Task AddAsync(string format, string extension, byte[] bytes) {
            EpubManuscriptResult result = await BookManuscriptImporter.ImportBytesAsync(bytes, extension,
                new EpubManuscriptOptions { Identifier = "urn:officeimo:fixture:manuscript:" + format });
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
        }
    }
}
