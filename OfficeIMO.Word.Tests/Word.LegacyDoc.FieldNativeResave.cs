using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using System.Text;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    private static bool HasCachedPageField(WordParagraph paragraph, string cache = "[Document Page]", bool bold = false) =>
        paragraph._paragraph.Descendants<W.SimpleField>()
            .Any(field => IsFieldInstruction(field, "PAGE") && field.InnerText == cache && field.Descendants<W.Run>()
                .All(run => run.RunProperties?.NoProof is { } noProof && (noProof.Val?.Value ?? true)
                    && (run.RunProperties?.Bold is { } value && (value.Val?.Value ?? true)) == bold));

    [Theory]
    [InlineData(0, 0)]
    [InlineData(0, 1)]
    [InlineData(0, 2)]
    [InlineData(1, 0)]
    [InlineData(1, 1)]
    [InlineData(1, 2)]
    [InlineData(2, 0)]
    [InlineData(2, 1)]
    [InlineData(2, 2)]
    [InlineData(3, 0)]
    [InlineData(3, 1)]
    [InlineData(3, 2)]
    public void LegacyDoc_FieldCaches_PreserveNativeResaveAcrossStories(int story, int revision) {
        using var input = new MemoryStream(CreateNativeDateFieldCache(story, revision));
        using WordDocument imported = WordDocument.Load(input);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        try {
            imported.Save(path);
            using WordDocument reopened = WordDocument.Load(path);
            Assert.Equal(WordFileFormat.Doc, reopened.SourceFormat);
            Assert.Empty(reopened.LegacyDocUnsupportedFeatures);
            W.SimpleField field = Assert.Single(EnumerateFieldFormattingRoots(reopened._wordprocessingDocument.MainDocumentPart!)
                .SelectMany(root => root.Descendants<W.SimpleField>()));
            Assert.Equal("DATE", field.Instruction!.Value!.Trim());
            W.Text first = Assert.Single(field.Descendants<W.Text>(), text => text.Text == "FIRST");
            W.Bold? firstBold = first.Parent!.GetFirstChild<W.RunProperties>()?.Bold;
            Assert.True(firstBold == null || firstBold.Val?.Value == false);
            if (revision == 0) {
                W.Run second = Assert.Single(field.Descendants<W.Run>(), run => run.InnerText == "SECOND");
                Assert.NotNull(second.RunProperties?.Bold);
                Assert.True(second.RunProperties.Bold.Val?.Value ?? true);
            } else if (revision == 1) {
                W.InsertedRun inserted = Assert.Single(field.Descendants<W.InsertedRun>());
                Assert.Equal("SECOND", inserted.InnerText);
                Assert.False(string.IsNullOrWhiteSpace(inserted.Author?.Value));
            } else {
                W.DeletedRun deleted = Assert.Single(field.Descendants<W.DeletedRun>());
                Assert.Equal("SECOND", Assert.Single(deleted.Descendants<W.DeletedText>()).Text);
                Assert.False(string.IsNullOrWhiteSpace(deleted.Author?.Value));
            }
            Assert.Empty(reopened.ValidateDocument());
        } finally { DeleteIfExists(path); }
    }

    [Theory]
    [InlineData("PAGE", false)]
    [InlineData("PAGE", true)]
    [InlineData("NUMPAGES", false)]
    [InlineData("NUMPAGES", true)]
    [InlineData("PAGE \\* ARABIC", false)]
    [InlineData("NUMPAGES \\* ARABIC", true)]
    public void LegacyDoc_FieldCountCaches_RetainNativeTextAndInteriorBookmarkOffsets(string instruction, bool complex) {
        using WordDocument source = WordDocument.Create();
        W.Paragraph[] paragraphs = CreateFieldFormattingStories(source);
        for (int index = 0; index < paragraphs.Length; index++) {
            string name = "CountCache" + index;
            AppendCountCache(paragraphs[index], (200 + index).ToString(), name, instruction, complex);
        }
        string first = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        string second = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        try {
            source.Save(first);
            using WordDocument imported = WordDocument.Load(first);
            AssertCountCaches(imported);
            imported.Save(second);
            using WordDocument reopened = WordDocument.Load(second);
            AssertCountCaches(reopened);
        } finally { DeleteIfExists(first); DeleteIfExists(second); }

        void AssertCountCaches(WordDocument document) {
            W.SimpleField[] fields = EnumerateFieldFormattingRoots(document._wordprocessingDocument.MainDocumentPart!)
                .SelectMany(root => root.Descendants<W.SimpleField>()).ToArray();
            Assert.Equal(6, fields.Length);
            Assert.All(fields, field => {
                Assert.Equal(instruction, field.Instruction!.Value!.Trim());
                Assert.Equal("1234", field.InnerText);
                W.BookmarkStart start = Assert.Single(field.Descendants<W.BookmarkStart>());
                W.BookmarkEnd end = Assert.Single(field.Descendants<W.BookmarkEnd>());
                Assert.Equal(start.Id!.Value, end.Id!.Value);
                var selection = field.Descendants().SkipWhile(element => !ReferenceEquals(element, start)).Skip(1)
                    .TakeWhile(element => !ReferenceEquals(element, end)).OfType<W.Text>();
                Assert.Equal("34", string.Concat(selection.Select(text => text.Text)));
            });
            Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "1234");
            Assert.Empty(document.ValidateDocument());
        }
    }

    private static void AppendCountCache(W.Paragraph paragraph, string id, string name, string instruction, bool complex) {
        paragraph.RemoveAllChildren<W.Run>();
        var result = new DocumentFormat.OpenXml.OpenXmlElement[] {
            new W.Run(new W.Text("12")), new W.BookmarkStart { Id = id, Name = name },
            new W.Run(new W.Text("34")), new W.BookmarkEnd { Id = id }
        };
        if (!complex) { paragraph.Append(new W.SimpleField(result) { Instruction = " " + instruction + " " }); return; }
        paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" " + instruction + " ")),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }));
        paragraph.Append(result);
        paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
    }

    private static byte[] CreateNativeDateFieldCache(int story, int revision) {
        byte[] input = LegacyDocTestBuilder.CreateFieldCacheCarrier(story);
        Assert.True(OfficeCompoundFileReader.TryRead(input, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"].ToArray();
        byte[] table = compound.Streams["1Table"];
        int instruction = Encoding.ASCII.GetString(main).IndexOf(" FILENAME ", StringComparison.Ordinal);
        Assert.True(instruction >= 0);
        Encoding.ASCII.GetBytes(" DATE     ").CopyTo(main, instruction);
        if (revision != 0) {
            int bin = BitConverter.ToInt32(main, 0xFA);
            int property = BitConverter.ToInt32(table, bin + 8) * 512 + 0xF0;
            Assert.Equal(new byte[] { 3, 0x35, 8, 1 }, main.Skip(property).Take(4));
            main[property + 1] = revision == 1 ? (byte)1 : (byte)0;
        }
        return OfficeCompoundFileWriter.Write(compound.Streams.Select(stream =>
            new OfficeCompoundStream(stream.Key, stream.Key == "WordDocument" ? main : stream.Value)).ToArray());
    }
}
