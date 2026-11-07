using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using System.Text;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, "_PREFIX")]
    [InlineData(1, " PREFIX")]
    [InlineData(3, "   EFIX")]
    public void LegacyDoc_NotePrefixes_StrictDocxExportPreservesLiteralSourceAndNativeResaveAvoidsDuplication(int spaces, string prefix) {
        using WordDocument authoring = CreateNotePrefixDocument("PREFIX");
        byte[] source = authoring.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc });
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out _));
        byte[] text = compound!.Streams["WordDocument"].ToArray();
        // This ASCII-only carrier uses the writer's compressed text profile; replacements retain every CP and byte offset.
        byte[] pattern = Encoding.ASCII.GetBytes("\u0002 PREFIX");
        byte[] replacement = Encoding.ASCII.GetBytes("\u0002" + prefix);
        Assert.Equal(pattern.Length, replacement.Length);
        int replacements = 0;
        for (int index = 0; index <= text.Length - pattern.Length; index++) {
            if (!text.Skip(index).Take(pattern.Length).SequenceEqual(pattern)) continue;
            Buffer.BlockCopy(replacement, 0, text, index, replacement.Length);
            replacements++;
            index += pattern.Length - 1;
        }
        Assert.Equal(2, replacements);
        source = OfficeCompoundFileWriter.Write(compound.Streams.Select(stream =>
            new OfficeCompoundStream(stream.Key, stream.Key == "WordDocument" ? text : stream.Value)).ToArray());

        using WordDocument imported = WordDocument.Load(new MemoryStream(source));
        AssertNotePrefixContent(imported, prefix, spaces == 3 ? "  EFIX" : "PREFIX");
        using var modern = new MemoryStream();
        imported.Save(modern);
        modern.Position = 0;
        using WordDocument reopened = WordDocument.Load(modern);
        AssertNotePrefixContent(reopened, prefix, spaces == 3 ? "  EFIX" : "PREFIX");
        byte[] native = reopened.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc });
        string nativePrefix = spaces == 0 ? " _PREFIX" : prefix;
        AssertNativeNotePrefixBytes(native, nativePrefix);
        using WordDocument nativeReopened = WordDocument.Load(new MemoryStream(native));
        AssertNotePrefixContent(nativeReopened, nativePrefix, spaces == 3 ? "  EFIX" : "PREFIX");
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(3)]
    public void LegacyDoc_NotePrefixes_DocxNativeEncodingRetainsExplicitRevisionPrefixAndAuthoredDefault(int spaces) {
        string prefix = new string(' ', spaces) + "PREFIX";
        using WordDocument authoring = CreateNotePrefixDocument(prefix);
        using var modern = new MemoryStream();
        authoring.Save(modern);
        modern.Position = 0;
        using WordDocument reopened = WordDocument.Load(modern);
        AssertNotePrefixContent(reopened, prefix, prefix);
        byte[] native = reopened.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc });
        string expectedPrefix = spaces == 0 ? " PREFIX" : prefix;
        AssertNativeNotePrefixBytes(native, expectedPrefix);
        using WordDocument nativeReopened = WordDocument.Load(new MemoryStream(native));
        AssertNotePrefixContent(nativeReopened, expectedPrefix, prefix);
    }

    private static WordDocument CreateNotePrefixDocument(string prefix) {
        WordDocument source = WordDocument.Create();
        WordFootNote footnote = source.AddParagraph("Footnote anchor").AddFootNote("PREFIX");
        WordEndNote endnote = source.AddParagraph("Endnote anchor").AddEndNote("PREFIX");
        var main = source._wordprocessingDocument.MainDocumentPart!;
        W.Footnote first = main.FootnotesPart!.Footnotes!.Elements<W.Footnote>().Single(note => note.Id?.Value == footnote.ReferenceId);
        W.Endnote second = main.EndnotesPart!.Endnotes!.Elements<W.Endnote>().Single(note => note.Id?.Value == endnote.ReferenceId);
        Set(first, new W.FootnoteReferenceMark(), "FootnotePrefix", "301");
        Set(second, new W.EndnoteReferenceMark(), "EndnotePrefix", "302");
        return source;

        void Set(OpenXmlCompositeElement note, OpenXmlElement mark, string name, string id) {
            note.RemoveAllChildren();
            var paragraph = new W.Paragraph(new W.Run(mark), new W.BookmarkStart { Name = name, Id = id },
                new W.InsertedRun(new W.Run(new W.RunProperties(new W.Bold()), new W.Text(prefix) { Space = SpaceProcessingModeValues.Preserve })) {
                    Id = id, Author = "Prefix author", Date = new DateTime(2020, 1, 2, 3, 4, 0, DateTimeKind.Utc) },
                new W.BookmarkEnd { Id = id },
                new W.SimpleField(new W.Run(new W.Text("FIRST")),
                    new W.DeletedRun(new W.Run(new W.RunProperties(new W.Bold()), new W.DeletedText("SECOND"))) {
                        Id = (int.Parse(id) + 10).ToString(), Author = "Cache author", Date = new DateTime(2020, 1, 2, 3, 4, 0, DateTimeKind.Utc) }) {
                    Instruction = " DATE ", FieldLock = true });
            note.Append(paragraph);
        }
    }

    private static void AssertNotePrefixContent(WordDocument source, string prefix, string revisionText) {
        var main = source._wordprocessingDocument.MainDocumentPart!;
        OpenXmlCompositeElement[] notes = {
            main.FootnotesPart!.Footnotes!.Elements<W.Footnote>().Single(note => note.Id?.Value > 0),
            main.EndnotesPart!.Endnotes!.Elements<W.Endnote>().Single(note => note.Id?.Value > 0)
        };
        foreach (OpenXmlCompositeElement note in notes) {
            string text = string.Concat(note.Descendants().Select(element => element switch {
                W.Text value => value.Text, W.DeletedText value => value.Text, _ => string.Empty }));
            Assert.Equal(prefix + "FIRSTSECOND", text);
            W.InsertedRun inserted = Assert.Single(note.Descendants<W.InsertedRun>());
            Assert.Equal(revisionText, inserted.InnerText);
            Assert.Equal("Prefix author", inserted.Author?.Value);
            Assert.All(inserted.Descendants<W.Run>(), run => Assert.True(run.RunProperties?.Bold?.Val?.Value ?? run.RunProperties?.Bold != null));
            W.SimpleField field = Assert.Single(note.Descendants<W.SimpleField>());
            Assert.Equal(" DATE ", field.Instruction?.Value);
            Assert.Equal("FIRSTSECOND", field.InnerText);
            Assert.True(field.FieldLock?.Value);
            Assert.Equal("SECOND", Assert.Single(field.Descendants<W.DeletedText>()).Text);
            Assert.Equal("Cache author", Assert.Single(field.Descendants<W.DeletedRun>()).Author?.Value);
            W.BookmarkStart bookmark = Assert.Single(note.Descendants<W.BookmarkStart>());
            W.BookmarkEnd end = Assert.Single(note.Descendants<W.BookmarkEnd>());
            var selected = new StringBuilder();
            bool inside = false;
            foreach (OpenXmlElement element in note.Descendants()) {
                if (ReferenceEquals(element, bookmark)) inside = true;
                else if (ReferenceEquals(element, end)) inside = false;
                else if (inside && element is W.Text value) selected.Append(value.Text);
            }
            Assert.Equal(revisionText, selected.ToString());
        }
    }

    private static void AssertNativeNotePrefixBytes(byte[] bytes, string prefix) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out _));
        string text = Encoding.ASCII.GetString(compound!.Streams["WordDocument"]);
        Assert.Equal(2, text.Split(new[] { "\u0002" + prefix + "\u0013 DATE " }, StringSplitOptions.None).Length - 1);
    }
}
