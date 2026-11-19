using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Reader.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPowerPointTextProjectionTests {
    private const string StreamedCapability = "officeimo.reader.powerpoint.streamed-text";

    public static IEnumerable<object[]> Decks() {
        string root = RepositoryTestPaths.Find();
        // PowerPoint-authored templates and Keynote exports: two independent producers.
        foreach (string path in Directory.GetFiles(Path.Combine(root, "Assets", "PowerPointTemplates"), "*.pptx")
                     .Concat(Directory.GetFiles(Path.Combine(root, "OfficeIMO.TestAssets", "Documents", "IWorkCorpus", "native-exports"), "*.pptx"))
                     .OrderBy(path => path, StringComparer.Ordinal)) {
            yield return new object[] { path.Substring(root.Length + 1).Replace('\\', '/') };
        }
    }

    // The text projection reads slides one at a time; its chunks must be the ones the model projection reduces to.
    [Theory]
    [MemberData(nameof(Decks))]
    public void TextProjectionStreamsSlidesWithTheModelText(string deck) {
        string path = Path.Combine(RepositoryTestPaths.Find(), deck);
        var options = new ReaderOptions { Projection = ReaderProjection.Text, MaxTableRows = 2 };
        foreach (bool includeHidden in new[] { true, false }) {
            var formatOptions = new ReaderPowerPointOptions { IncludeHiddenShapes = includeHidden };
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPowerPointHandler(formatOptions).Build();

            OfficeDocumentReadResult fromPath = reader.ReadDocument(path, options);
            OfficeDocumentReadResult fromStream;
            using (FileStream stream = File.OpenRead(path)) fromStream = reader.ReadDocument(stream, path, options);
            OfficeDocumentReadResult model;
            using (PowerPointDocument presentation = PowerPointDocument.Load(path, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                model = presentation.ReadDocument(path, options, formatOptions);
            }

            Assert.Contains(StreamedCapability, fromPath.CapabilitiesUsed);
            Assert.Contains(StreamedCapability, fromStream.CapabilitiesUsed);
            Assert.DoesNotContain(StreamedCapability, model.CapabilitiesUsed);
            Assert.Equal(Describe(model), Describe(fromPath));
            Assert.Equal(Describe(model), Describe(fromStream));
        }
    }

    [Fact]
    public void TextProjectionReportsInputLimitsLikeTheModelRead() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "Assets", "PowerPointTemplates", "Tables.pptx");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPowerPointHandler().Build();
        long limit = new FileInfo(path).Length - 1;

        Exception text = Record.Exception(() => reader.ReadDocument(path,
            new ReaderOptions { Projection = ReaderProjection.Text, MaxInputBytes = limit }));
        Exception lean = Record.Exception(() => reader.ReadDocument(path,
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, MaxInputBytes = limit }));

        Assert.NotNull(lean);
        Assert.Equal(lean.GetType(), text?.GetType());
        Assert.Equal(lean.Message, text!.Message);
    }

    private static string[] Describe(OfficeDocumentReadResult result) =>
        result.Chunks.Select(chunk => string.Join("|", chunk.Id, chunk.Kind, chunk.Location.Slide, chunk.Location.BlockIndex,
                chunk.Location.SourceBlockIndex, chunk.Location.SourceBlockKind, chunk.Location.BlockAnchor, chunk.Location.LogicalOrder,
                chunk.Text))
            .Concat(new[] {
                "source|" + result.Source.Title + "|" + result.Source.Author + "|" + result.Source.Subject + "|" + result.Source.Keywords,
                "text|" + result.GetText()
            })
            .ToArray();
}
