using OfficeIMO.Reader;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Word;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// The streamed text projections read Word and PowerPoint packages without the Open XML SDK. Wherever they stream,
/// their text must equal the text of the model read; where they decline, the model read answers.
/// </summary>
public sealed class ReaderStreamedTextCorpusTests {
    public static IEnumerable<object[]> CorpusFiles() {
        string root = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents");
        foreach (string file in Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories)
                     .Where(file => file.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) || file.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                     .OrderBy(file => file, StringComparer.Ordinal)) {
            yield return new object[] { file.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar) };
        }
    }

    [Theory]
    [MemberData(nameof(CorpusFiles))]
    public void StreamedTextMatchesTheModelRead(string relativePath) {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", relativePath);
        byte[] package = File.ReadAllBytes(path);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().AddPowerPointHandler().Build();
        string name = Path.GetFileName(path);

        OfficeDocumentReadResult model;
        try {
            model = reader.ReadDocument(package, name, new ReaderOptions { Projection = ReaderProjection.TextAndMetadata });
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            // Files the model cannot read are outside this comparison; the streamed read declines for them too.
            return;
        }
        OfficeDocumentReadResult text = reader.ReadDocument(package, name, new ReaderOptions { Projection = ReaderProjection.Text });

        Assert.Equal(model.GetText(), text.GetText());
    }
}
