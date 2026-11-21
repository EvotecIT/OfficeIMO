using System.IO.Compression;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class WordOdtStyleCatalogTests {
    [Theory]
    [InlineData("DocumentWithFootNotes.docx")]
    [InlineData("sample1.docx")]
    public void WordToOdtReadsStylesWithoutLoadingStylesPartAndMatchesLoadedDom(string fixture) {
        byte[] package = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", fixture));

        (string Output, bool StylesLoaded) streamed = Convert(package, loadStylesFirst: false);
        Assert.False(streamed.StylesLoaded);
        Assert.Equal(Convert(package, loadStylesFirst: true).Output, streamed.Output);
    }

    private static (string Output, bool StylesLoaded) Convert(byte[] package, bool loadStylesFirst) {
        using var input = new MemoryStream(package, writable: false);
        using WordDocument source = WordDocument.Load(input, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        var stylesPart = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!;
        if (loadStylesFirst) _ = stylesPart.Styles;
        OdfConversionResult<OdtDocument> conversion = source.ConvertToOpenDocumentWithEvidence();
        var output = new List<string>();
        output.AddRange(conversion.Report.Mappings.Select(mapping => mapping.Feature + "|" + mapping.Status + "|" + mapping.Count));
        using (var archive = new ZipArchive(new MemoryStream(conversion.Value.ToBytes()), ZipArchiveMode.Read)) {
            foreach (string part in new[] { "content.xml", "styles.xml" }) {
                using var reader = new StreamReader(archive.GetEntry(part)!.Open());
                output.Add(reader.ReadToEnd());
            }
        }
        return (string.Join("\n", output), stylesPart.IsRootElementLoaded);
    }
}
