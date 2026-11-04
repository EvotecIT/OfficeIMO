using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordFieldFormatNameTests {
    [Theory]
    [InlineData(WordFieldFormat.UpperRoman, "Roman")]
    [InlineData(WordFieldFormat.LowerRoman, "roman")]
    [InlineData(WordFieldFormat.UpperLetter, "ALPHABETICAL")]
    [InlineData(WordFieldFormat.LowerLetter, "Alphabetical")]
    [InlineData(WordFieldFormat.MergeFormat, "Mergeformat")]
    public void PublicFormatNamesRetainWordInstructionTokens(WordFieldFormat format, string token) {
        using var document = WordDocument.Create();
        WordField.AddField(document.AddParagraph(), WordFieldType.Page, format);
        WordField.AddField(document.AddParagraph(), WordFieldType.Page, format, advanced: true);
        using var output = new MemoryStream();
        document.Save(output);
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        using var part = archive.GetEntry("word/document.xml")!.Open();
        var xml = XDocument.Load(part);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var instructions = xml.Descendants(w + "fldSimple").Select(field => field.Attribute(w + "instr")!.Value)
            .Concat(xml.Descendants(w + "instrText").Select(field => field.Value)).ToArray();
        Assert.Equal(2, instructions.Length);
        foreach (string instruction in instructions) {
            Assert.Contains("\\* " + token + " ", instruction, StringComparison.Ordinal);
            var parsed = new WordFieldParser(instruction);
            Assert.Equal(format, parsed.FormatSwitches[0]);
            Assert.Contains(WordFieldFormat.MergeFormat, parsed.FormatSwitches);
        }
    }
}
