using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(101)]
    [InlineData(104)]
    public void LegacyWord69_FieldLocks_ProjectSourceBackedOldHeaderPlcIntoDocx(int version) {
        // This is a synthetic lock carrier using the qualified Word6/7 grammar,
        // not an independently produced locked legacy document.
        byte[] original = CreateWord69StructureFixture((ushort)version);
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
        byte[] main = compound!.Streams["WordDocument"].ToArray();
        int body = BitConverter.ToInt32(main, 0x34);
        int header = BitConverter.ToInt32(main, 0x3C);
        int text = BitConverter.ToInt32(main, 0x18) + body;
        var records = new List<int>();
        for (int cp = 0; cp < header; cp++) {
            if (main[text + cp] is 0x13 or 0x14 or 0x15) records.Add(cp);
        }
        int start = main.Length;
        int length = 4 + 6 * records.Count;
        Array.Resize(ref main, start + length);
        int properties = start + (records.Count + 1) * 4;
        for (int index = 0; index < records.Count; index++) {
            int cp = records[index];
            byte delimiter = main[text + cp];
            WriteInt(main, start + index * 4, cp);
            main[properties + index * 2] = (byte)(delimiter | 0x80);
            main[properties + index * 2 + 1] = delimiter == 0x15 ? (byte)0x90 : delimiter == 0x14 ? (byte)0xFF : (byte)1;
        }
        WriteInt(main, start + records.Count * 4, int.MaxValue);
        WriteInt(main, 0xE0, start);
        WriteInt(main, 0xE4, length);
        WriteInt(main, 0x20, main.Length);
        byte[] input = OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream("WordDocument", main) });
        using WordDocument document = WordDocument.Load(new MemoryStream(input));
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var package = new ZipArchive(output, ZipArchiveMode.Read);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var fields = new List<XElement>();
        foreach (string part in new[] { "word/header1.xml", "word/footer1.xml" }) {
            using Stream stream = package.GetEntry(part)!.Open();
            fields.AddRange(XDocument.Load(stream).Descendants(w + "fldSimple"));
        }
        Assert.Equal(3, fields.Count);
        Assert.All(fields, field => Assert.Contains((string?)field.Attribute(w + "fldLock"), new[] { "true", "1", "on" }));
        Assert.Equal(new[] { "Heading", "old.doc", "1" }.OrderBy(value => value),
            fields.Select(field => string.Concat(field.Descendants(w + "t").Select(textNode => textNode.Value))).OrderBy(value => value));
    }
}
