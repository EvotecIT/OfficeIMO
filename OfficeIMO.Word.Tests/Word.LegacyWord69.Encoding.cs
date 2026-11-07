using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
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
