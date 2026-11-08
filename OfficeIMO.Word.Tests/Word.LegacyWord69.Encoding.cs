using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyWord69_DirectOrInheritedWindows1257Font_PreservesBalticTextBesideWesternText(bool inheritedFont) {
        byte[] original = CreateWord69Fixture(104, cyrillicFont: inheritedFont ? 2 : 1);
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"].ToArray();
        main[0x270 + 2 + 8 + 4] = 186;
        byte[] baltic = { 0xC2, 0xE2, 0xC7, 0xE7, 0xCC, 0xEC, 0xCE };
        baltic.CopyTo(main, 0x300);
        main[0x300 + 13] = 0xE2;
        byte[] input = OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream("WordDocument", main) });

        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        Assert.Empty(document.LegacyDocUnsupportedFeatures);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "ĀāĒēĢģĪ");
        WordTableRow row = Assert.Single(Assert.Single(reopened.Tables).Rows);
        Assert.Equal(new[] { "Left", "âight" }, row.Cells.Select(cell => Assert.Single(cell.Paragraphs).Text));
    }

    [Fact]
    public void LegacyWord69_Windows1257ControlSlotsRequireExplicitLossBeforeDocxSave() {
        byte[] original = CreateWord69Fixture(104, cyrillicFont: 1);
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"].ToArray();
        main[0x270 + 2 + 8 + 4] = 186;
        main[0x300] = 0x8A;
        byte[] input = OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream("WordDocument", main) });
        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        Assert.Equal("doc-word69-text-codepage-unsupported", Assert.Single(document.LegacyDocUnsupportedFeatures).Code);
        using var strict = new MemoryStream();
        byte[] sentinel = { 0x51, 0x52, 0x53 };
        strict.Write(sentinel, 0, sentinel.Length);
        strict.Position = 0;
        Assert.Throws<NotSupportedException>(() => document.Save(strict));
        Assert.Equal(sentinel, strict.ToArray());
        using var allowed = new MemoryStream();
        document.Save(allowed, new WordSaveOptions { AllowLoss = true });
        allowed.Position = 0;
        using WordDocument reopened = WordDocument.Load(allowed);
        Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "eading");
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void LegacyWord69_Windows1257NotesRetainBalticTextAndWesternStoryBoundary() {
        const string expected = "Ābele ģimenē ēd ķiršus; žąsis Įlankoje; Õun.";
        byte[] baltic = {
            0xC2, 0x62, 0x65, 0x6C, 0x65, 0x20, 0xEC, 0x69, 0x6D, 0x65, 0x6E,
            0xE7, 0x20, 0xE7, 0x64, 0x20, 0xED, 0x69, 0x72, 0xF0, 0x75, 0x73,
            0x3B, 0x20, 0xFE, 0xE0, 0x73, 0x69, 0x73, 0x20, 0xC1, 0x6C, 0x61,
            0x6E, 0x6B, 0x6F, 0x6A, 0x65, 0x3B, 0x20, 0xD5, 0x75, 0x6E, 0x2E
        };
        string rawBytes = new string(baltic.Select(value => (char)value).ToArray());
        byte[] input = CreateWord69NotesFixture(104,
            footnoteText: "\u0002\t" + rawBytes + "\rSecond Baltic paragraph\r",
            mutate: main => main[0x282 + 8 + 4] = 186);
        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        Assert.Empty(document.LegacyDocUnsupportedFeatures);
        WordFootNote footnote = Assert.Single(document.FootNotes);
        Assert.Equal(new[] { "\t" + expected, "Second Baltic paragraph" }, footnote.TextParagraphs);
        Assert.Equal(new[] { "Eøndnote" }, Assert.Single(document.EndNotes).TextParagraphs);
        WordParagraph editableRun = footnote.Paragraphs!.Single(run => run.Text.Contains("Ābele"));
        editableRun.SetText(editableRun.Text + " Reviewed");
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        Assert.Equal(new[] { "\t" + expected + " Reviewed", "Second Baltic paragraph" }, Assert.Single(reopened.FootNotes).TextParagraphs);
        Assert.Equal(new[] { "Eøndnote" }, Assert.Single(reopened.EndNotes).TextParagraphs);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void LegacyWord69_OpenOfficeWindows1250_LoadEditSaveRetainsCentralEuropeanText() {
        string source = GetFixtureDoc(Path.Combine("LegacyDocCorpus", "apache-poi-word69", "Bug60936.doc"));
        using WordDocument document = WordDocument.Load(source);
        Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Contains("4 skóre a před 7 lety"));
        Assert.DoesNotContain(document.LegacyDocUnsupportedFeatures,
            finding => finding.Code == "doc-word69-text-codepage-unsupported");
        Assert.Contains(document.LegacyDocUnsupportedFeatures, finding => finding.Code == "doc-word69-formatting-unsupported");
        Assert.Contains(document.LegacyDocUnsupportedFeatures, finding => finding.Code == "doc-word69-bookmarks");
        document.AddParagraph("Reviewed for migration");
        using var output = new MemoryStream();
        Assert.Throws<NotSupportedException>(() => document.Save(output));
        Assert.Equal(0, output.Length);
        document.Save(output, new WordSaveOptions { AllowLoss = true });

        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text.Contains("4 skóre a před 7 lety"));
        Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "Reviewed for migration");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyWord69_DirectOrInheritedWindows1250Font_PreservesAdjacentWesternText(bool inheritedFont) {
        byte[] original = CreateWord69Fixture(104, cyrillicFont: inheritedFont ? 2 : 1);
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"].ToArray();
        main[0x270 + 2 + 8 + 4] = 238;
        main[0x300] = 0xF8;
        main[0x300 + 13] = 0xF8;
        byte[] input = OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream("WordDocument", main) });

        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "řeading");
        WordTableRow row = Assert.Single(Assert.Single(reopened.Tables).Rows);
        Assert.Equal(new[] { "Left", "øight" },
            row.Cells.Select(cell => Assert.Single(cell.Paragraphs).Text));
    }

    [Fact]
    public void LegacyWord69_Windows1250HeaderFieldCache_PreservesStoryEncodingBoundary() {
        byte[] original = CreateWord69StructureFixture(104);
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"].ToArray();
        int textStart = BitConverter.ToInt32(main, 0x18);
        int bodyLength = BitConverter.ToInt32(main, 0x34);
        int textLength = BitConverter.ToInt32(main, 0x1C) - textStart;
        main[0x270 + 2 + 8 + 4] = 238;
        main[textStart] = 0xF8;
        main[textStart + bodyLength + "\u0013 FILENAME \u0014".Length] = 0xF8;
        WriteWord69EncodingCharacterPage(main, textStart, bodyLength, textLength);
        byte[] input = OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream("WordDocument", main) });

        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var package = new ZipArchive(output, ZipArchiveMode.Read);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using Stream headerStream = package.GetEntry("word/header1.xml")!.Open();
        XDocument header = XDocument.Load(headerStream);
        XElement filename = Assert.Single(header.Descendants(w + "fldSimple"),
            field => ((string?)field.Attribute(w + "instr"))?.Trim() == "FILENAME");
        Assert.Equal("řld.doc", string.Concat(filename.Descendants(w + "t").Select(text => text.Value)));
        using Stream bodyStream = package.GetEntry("word/document.xml")!.Open();
        XDocument body = XDocument.Load(bodyStream);
        Assert.Contains(body.Descendants(w + "p"),
            paragraph => string.Concat(paragraph.Descendants(w + "t").Select(text => text.Value)) == "øeading");
    }

    private static void WriteWord69EncodingCharacterPage(byte[] main, int textStart, int bodyLength, int textLength) {
        const int page = 0x400;
        Array.Clear(main, page, 512);
        WriteInt(main, page, textStart);
        WriteInt(main, page + 4, textStart + bodyLength);
        WriteInt(main, page + 8, textStart + textLength);
        main[page + 13] = 250;
        main[page + 500] = 3;
        main[page + 501] = 93;
        WriteWord(main, page + 502, 1);
        main[page + 511] = 2;
    }
}
