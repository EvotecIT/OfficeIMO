using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(101)]
    [InlineData(104)]
    public void LegacyWord69_SectionsAndHeaderTables_SaveEditableDocx(int version) {
        using var input = new MemoryStream(CreateWord69StructureFixture((ushort)version));
        using WordDocument imported = WordDocument.Load(input);
        imported.AddParagraph("Edited after import");
        using var output = new MemoryStream();
        imported.Save(output);

        output.Position = 0;
        using (var package = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true)) {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XDocument body = ReadPart("word/document.xml");
            XElement[] sections = body.Descendants(w + "sectPr").ToArray();
            Assert.Equal(2, sections.Length);
            Assert.Equal("15840", (string?)sections[0].Element(w + "pgSz")?.Attribute(w + "w"));
            Assert.Equal("12240", (string?)sections[0].Element(w + "pgSz")?.Attribute(w + "h"));
            Assert.Equal("landscape", (string?)sections[0].Element(w + "pgSz")?.Attribute(w + "orient"));
            Assert.Equal("900", (string?)sections[0].Element(w + "pgMar")?.Attribute(w + "left"));
            Assert.Equal("600", (string?)sections[0].Element(w + "pgMar")?.Attribute(w + "header"));
            Assert.Equal("3", (string?)sections[0].Element(w + "pgNumType")?.Attribute(w + "start"));
            Assert.Equal("2", (string?)sections[1].Element(w + "cols")?.Attribute(w + "num"));
            XElement heading = Assert.Single(body.Descendants(w + "p"), p => string.Concat(p.Descendants(w + "t").Select(t => t.Value)) == "Heading");
            Assert.Equal("240", (string?)heading.Element(w + "pPr")?.Element(w + "spacing")?.Attribute(w + "before"));
            Assert.Equal("1440", (string?)heading.Descendants(w + "tab").FirstOrDefault()?.Attribute(w + "pos"));
            Assert.Equal("6", (string?)heading.Descendants(w + "top").FirstOrDefault()?.Attribute(w + "sz"));

            XDocument header = ReadPart("word/header1.xml");
            XElement row = Assert.Single(Assert.Single(header.Descendants(w + "tbl")).Elements(w + "tr"));
            Assert.Equal(new[] { "old.doc", "Heading" }, row.Elements(w + "tc")
                .Select(cell => string.Concat(cell.Descendants(w + "t").Select(t => t.Value))));
            Assert.Contains(header.Descendants(w + "fldSimple"), field => ((string?)field.Attribute(w + "instr"))?.Trim() == "FILENAME");
            Assert.Contains(header.Descendants(w + "fldSimple"), field => ((string?)field.Attribute(w + "instr"))?.Trim() == "STYLEREF \"heading 1\"");
            XDocument footer = ReadPart("word/footer1.xml");
            Assert.Single(footer.Descendants(w + "pgNum"));
            Assert.Contains("Page ", string.Concat(footer.Descendants(w + "t").Select(t => t.Value)));

            XDocument ReadPart(string name) {
                using Stream stream = package.GetEntry(name)!.Open();
                return XDocument.Load(stream);
            }
        }

        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        Assert.Equal(2, reopened.Sections.Count);
        Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "Edited after import");
    }

    [Theory]
    [InlineData("HEURE", "doc-word69-fields-unsupported", false)]
    [InlineData("STYLEREF \"Missing Style\"", "doc-word69-field-reference-unresolved", true)]
    public void LegacyWord69_HeaderFieldLoss_RequiresExplicitLoss(string instruction, string code, bool preservesInstruction) {
        using var input = new MemoryStream(CreateWord69StructureFixture(104, unsupportedInstruction: instruction));
        using WordDocument imported = WordDocument.Load(input);
        Assert.Contains(imported.LegacyDocUnsupportedFeatures, feature => feature.Code == code);
        using var output = new MemoryStream();
        Assert.Throws<NotSupportedException>(() => imported.Save(output));
        Assert.Equal(0, output.Length);
        imported.Save(output, new WordSaveOptions { AllowLoss = true });
        output.Position = 0;
        using var package = new ZipArchive(output, ZipArchiveMode.Read);
        using Stream headerStream = package.GetEntry("word/header1.xml")!.Open();
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XDocument header = XDocument.Load(headerStream);
        Assert.Contains("Heading", string.Concat(header.Descendants(w + "t").Select(t => t.Value)));
        Assert.Equal(preservesInstruction, header.Descendants(w + "fldSimple")
            .Any(field => ((string?)field.Attribute(w + "instr"))?.Trim() == instruction));
    }

    [Fact]
    public void LegacyWord69_HeaderRangeOutsideStory_FailsBeforeProjection() {
        using var input = new MemoryStream(CreateWord69StructureFixture(104, invalidHeaderRange: true));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(input));
        Assert.Contains("doc-word69-structure-invalid", error.Message);
        Assert.Contains("header/footer", error.Message);
    }

    private static byte[] CreateWord69StructureFixture(
        ushort version, string? unsupportedInstruction = null, bool invalidHeaderRange = false) {
        OfficeCompoundFileReader.TryRead(CreateWord69Fixture(version), out OfficeCompoundFile? compound, out _);
        byte[] main = new byte[0xB00];
        compound!.Streams["WordDocument"].CopyTo(main, 0);
        const int textStart = 0x300;
        int bodyLength = main[0x34];
        string cell1 = "\u0013 FILENAME \u0014old.doc\u0015\a";
        string instruction = unsupportedInstruction ?? "STYLEREF \"heading 1\"";
        string cell2 = "\u0013 " + instruction + " \u0014Heading\u0015\a";
        string header = cell1 + cell2 + "\a\r\r";
        string footer = "Page \u0013PAGE\u00141\u0015\r\r";
        byte[] headerText = Encoding.ASCII.GetBytes(header + footer + "\r");
        headerText.CopyTo(main, textStart + bodyLength);
        int textLength = bodyLength + headerText.Length;
        WriteInt(main, 0x1C, textStart + textLength);
        WriteInt(main, 0x20, main.Length);
        WriteInt(main, 0x3C, headerText.Length);
        WriteInt(main, 0x88, 0x1A0);
        WriteInt(main, 0x8C, 36);
        WriteInt(main, 0x1A0, 0);
        WriteInt(main, 0x1A4, 8);
        WriteInt(main, 0x1A8, bodyLength);
        WriteInt(main, 0x1AE, 0xA00);
        WriteInt(main, 0x1BA, 0xA40);
        WriteInt(main, 0xB0, 0x1C8);
        WriteInt(main, 0xB4, 16);
        WriteInt(main, 0x1CC, header.Length);
        WriteInt(main, 0x1D0, invalidHeaderRange ? headerText.Length + 100 : header.Length + footer.Length);
        WriteInt(main, 0x1D4, headerText.Length + 2);
        byte[] section1 = { 153, 10, 162, 2, 164, 224, 61, 165, 208, 47,
            166, 132, 3, 167, 232, 3, 156, 88, 2, 157, 188, 2, 150, 1, 161, 3, 0 };
        byte[] section2 = { 142, 0, 144, 1, 0, 145, 104, 1 };
        WriteWord(main, 0xA00, (ushort)section1.Length);
        section1.CopyTo(main, 0xA02);
        WriteWord(main, 0xA40, (ushort)section2.Length);
        section2.CopyTo(main, 0xA42);
        WritePage(0x400, 1, new[] { 0, 7, 13, 18, bodyLength, textLength },
            new[] { new byte[] { 85, 1 }, Array.Empty<byte>(), new byte[] { 86, 1 }, Array.Empty<byte>(), new byte[] { 85, 1 } }, false);
        WritePage(0x600, 7, new[] { 0, 8 },
            new[] { new byte[] { 1, 0, 15, 5, 0, 1, 160, 5, 0, 38, 9, 0, 21, 240, 0, 22, 120, 0 } }, true);
        byte[] inTable = { 0, 0, 24, 1 };
        byte[] rowEnd = { 0, 0, 24, 1, 25, 1, 190, 8, 0, 2, 0, 0, 104, 1, 208, 2 };
        int headerCell1End = bodyLength + cell1.Length;
        int headerCell2End = headerCell1End + cell2.Length;
        WritePage(0x800, 7, new[] { 8, 13, 19, 20, bodyLength, headerCell1End, headerCell2End, headerCell2End + 1, bodyLength + header.Length, textLength },
            new[] { inTable, inTable, rowEnd, new byte[] { 0, 0 }, inTable, inTable, rowEnd, new byte[] { 0, 0 }, new byte[] { 0, 0 } }, true);
        WriteInt(main, 0x204, textStart + textLength);
        WriteInt(main, 0x218, textStart + textLength);
        return OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream("WordDocument", main) });

        void WritePage(int offset, int stride, int[] positions, byte[][] records, bool paragraph) {
            Array.Clear(main, offset, 512);
            int cursor = 510;
            for (int index = 0; index < positions.Length; index++) WriteInt(main, offset + index * 4, textStart + positions[index]);
            for (int index = 0; index < records.Length; index++) {
                byte[] record = records[index];
                if (record.Length == 0) continue;
                int length = paragraph ? (record.Length + 1) & ~1 : record.Length;
                cursor = (cursor - 1 - length) & ~1;
                main[offset + positions.Length * 4 + index * stride] = (byte)(cursor / 2);
                main[offset + cursor] = (byte)(paragraph ? length / 2 : length);
                record.CopyTo(main, offset + cursor + 1);
            }
            main[offset + 511] = (byte)records.Length;
        }
    }
}
