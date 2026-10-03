namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyPathSaveTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PathSaveReplacesCompleteArtifactAndReopens(bool asynchronous) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-bibliography-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "references.bib");
        File.WriteAllText(path, "original");
        try {
            BibliographyDocument document = BibliographyDocument.Parse("@book{key,title={Title}}", BibliographyFormat.BibLatex).Document;
            document.Items[0].Title = "Edited";
            BibliographyWriteResult result = asynchronous ? await document.SaveAsync(path) : document.Save(path);
            Assert.Equal(result.Bytes, File.ReadAllBytes(path));
            Assert.Equal("Edited", BibliographyDocument.Load(path).Document.Items[0].Title);
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
        } finally {
            Directory.Delete(directory, true);
        }
    }
}
