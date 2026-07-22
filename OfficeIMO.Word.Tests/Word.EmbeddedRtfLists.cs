using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    private const string NumberedListRtf = """
        {\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}{\*\listtable{\list\listtemplateid77{\listlevel\levelnfc0\levelnfcn0\levelstartat1{\leveltext\'02\'00.;}{\levelnumbers\'01;}\fi-360\li720}{\listname Numbered;}\listid100}}{\*\listoverridetable{\listoverride\listid100\listoverridecount0\ls1}}\pard\plain\f0\fs22\ls1\ilvl0 Existing\par}
        """;

    private const string BulletListRtf = """
        {\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}{\*\listtable{\list\listtemplateid77{\listlevel\levelnfc23\levelnfcn23\levelstartat1{\leveltext\'01\u8226 ?;}{\levelnumbers;}\fi-360\li720}{\listname Bullet;}\listid100}}{\*\listoverridetable{\listoverride\listid100\listoverridecount0\ls1}}\pard\plain\f0\fs22\ls1\ilvl0 TODAY: Today\par\pard\plain\f0\fs22\ls1\ilvl0 YESTERDAY: Yesterday\par\pard Literal \\listid100 stays literal.\par}
        """;

    private const string NestedTableListRtf = """
        {\rtf1\ansi{\*\listtable{\list\listtemplateid77{\listlevel\levelnfc0\levelnfcn0\levelstartat4{\leveltext\'02\'00.;}{\levelnumbers\'01;}\fi-360\li720}{\listlevel\levelnfc23\levelnfcn23\levelstartat1{\leveltext\'01\u8226 ?;}{\levelnumbers;}\fi-360\li1440}{\listname Mixed;}\listid100}}{\*\listoverridetable{\listoverride\listid100\listoverridecount0\ls1}}{\trowd\cellx5000\pard\intbl\ls1\ilvl0 Outer four\par\pard\intbl\ls1\ilvl1 Nested bullet\cell\row}}
        """;

    private const string BinaryPayloadListRtf = """
        {\rtf1\ansi{\*\objdata\bin10 \listid100}{\*\listtable{\list\listtemplateid77{\listlevel\levelnfc23\levelnfcn23\levelstartat1{\leveltext\'01\u8226 ?;}{\levelnumbers;}\fi-360\li720}{\listname Bullet;}\listid100}}{\*\listoverridetable{\listoverride\listid100\listoverridecount0\ls1}}\pard\ls1\ilvl0 Bullet\par}
        """;

    [Fact]
    public void AddEmbeddedFragment_IsolatesCollidingRtfListIdentifiers() {
        using WordDocument document = WordDocument.Create();

        WordEmbeddedDocument numbered = document.AddEmbeddedFragment(
            NumberedListRtf,
            WordAlternativeFormatImportPartType.Rtf);
        WordEmbeddedDocument bullets = document.AddEmbeddedFragment(
            BulletListRtf,
            WordAlternativeFormatImportPartType.Rtf);

        string storedNumbered = ReadRtf(numbered);
        string storedBullets = ReadRtf(bullets);
        Assert.Equal(NumberedListRtf, storedNumbered);

        AssertIdentifierWasRemapped(storedNumbered, storedBullets, "listid");
        AssertIdentifierWasRemapped(storedNumbered, storedBullets, "ls");
        AssertIdentifierWasRemapped(storedNumbered, storedBullets, "listtemplateid");
        Assert.Contains(@"\levelnfc23", storedBullets, StringComparison.Ordinal);
        Assert.Contains(@"\\listid100 stays literal", storedBullets, StringComparison.Ordinal);
    }

    [Fact]
    public void AddEmbeddedFragmentAfter_AppliesRtfListIsolation() {
        using WordDocument document = WordDocument.Create();
        WordParagraph anchor = document.AddParagraph("Anchor");
        WordEmbeddedDocument numbered = document.AddEmbeddedFragment(
            NumberedListRtf,
            WordAlternativeFormatImportPartType.Rtf);

        WordEmbeddedDocument bullets = document.AddEmbeddedFragmentAfter(
            anchor,
            BulletListRtf,
            WordAlternativeFormatImportPartType.Rtf);

        AssertIdentifierWasRemapped(ReadRtf(numbered), ReadRtf(bullets), "listid");
        AssertIdentifierWasRemapped(ReadRtf(numbered), ReadRtf(bullets), "ls");
        AssertIdentifierWasRemapped(ReadRtf(numbered), ReadRtf(bullets), "listtemplateid");
    }

    [Fact]
    public void AddEmbeddedFragment_RemapKeepsNestedTableListReferencesConsistent() {
        using WordDocument document = WordDocument.Create();
        document.AddEmbeddedFragment(NumberedListRtf, WordAlternativeFormatImportPartType.Rtf);

        WordEmbeddedDocument nested = document.AddEmbeddedFragment(
            NestedTableListRtf,
            WordAlternativeFormatImportPartType.Rtf);
        string storedNested = ReadRtf(nested);

        int[] listIds = GetControlValues(storedNested, "listid");
        Assert.Equal(2, listIds.Length);
        Assert.Single(listIds.Distinct());
        Assert.DoesNotContain(100, listIds);

        int[] overrideIds = GetControlValues(storedNested, "ls");
        Assert.Equal(3, overrideIds.Length);
        Assert.Single(overrideIds.Distinct());
        Assert.DoesNotContain(1, overrideIds);
        Assert.Contains(@"\levelnfc0", storedNested, StringComparison.Ordinal);
        Assert.Contains(@"\levelnfc23", storedNested, StringComparison.Ordinal);
    }

    [Fact]
    public void AddEmbeddedFragment_RtfBinaryPayloadIsNotRewrittenAsControlWords() {
        using WordDocument document = WordDocument.Create();
        document.AddEmbeddedFragment(NumberedListRtf, WordAlternativeFormatImportPartType.Rtf);

        WordEmbeddedDocument bullets = document.AddEmbeddedFragment(
            BinaryPayloadListRtf,
            WordAlternativeFormatImportPartType.Rtf);
        string storedBullets = ReadRtf(bullets);

        Assert.Contains(@"\bin10 \listid100", storedBullets, StringComparison.Ordinal);
        Match listTableId = Regex.Match(
            storedBullets,
            @"\\listtable.*?\\listid(-?\d+)",
            RegexOptions.Singleline);
        Assert.True(listTableId.Success);
        Assert.NotEqual(100, int.Parse(listTableId.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void AddEmbeddedFragment_RejectsOversizedExistingRtfPartBeforeReadingItUnbounded() {
        using WordDocument document = WordDocument.Create();
        string oversized = "{\\rtf1 " + new string('A', 16 * 1024 * 1024) + "}";
        document.AddEmbeddedFragment(oversized, WordAlternativeFormatImportPartType.Rtf);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            document.AddEmbeddedFragment(BulletListRtf, WordAlternativeFormatImportPartType.Rtf));

        Assert.Contains("maximum size", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static string ReadRtf(WordEmbeddedDocument embedded) =>
        Encoding.UTF8.GetString(embedded.ToBytes());

    private static void AssertIdentifierWasRemapped(string first, string second, string controlWord) {
        int firstValue = Assert.Single(GetControlValues(first, controlWord).Distinct());
        int secondValue = Assert.Single(GetControlValues(second, controlWord).Distinct());
        Assert.NotEqual(firstValue, secondValue);
    }

    private static int[] GetControlValues(string rtf, string controlWord) {
        return Regex.Matches(rtf, @"(?<!\\)\\" + Regex.Escape(controlWord) + @"(-?\d+)")
            .Cast<Match>()
            .Select(match => int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture))
            .ToArray();
    }
}
