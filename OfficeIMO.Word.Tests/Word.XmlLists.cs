using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordMlListMigrationTests {
    private static readonly XNamespace Word = "http://schemas.microsoft.com/office/word/2003/wordml";

    [Fact]
    public void NativeWordMlListsRetainEditableDefinitionsAndNativeMarkers() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "WordXmlMigration", "word-lists-native-2003.xml");
        using WordDocument document = WordDocument.Load(path, null, out OfficeConversionReport report);
        Assert.Equal(WordFileFormat.WordMl2003, document.SourceFormat);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Location == (Word + "lists").ToString() ||
            diagnostic.Location == (Word + "listPr").ToString() || diagnostic.Location == (Word + "nfc").ToString());
        string[] expected = {
            "Decimal one|0|1.", "Decimal two|0|2.", "Decimal continued|0|3.", "Decimal restarted|0|1.",
            "Decimal starts seven|0|7.", "Bullet one|0|•", "Bullet two|0|•", "Outline first|0|1.",
            "Outline child one|1|1.1)", "Outline child two|1|1.2)", "Outline second|0|2."
        };
        Assert.Equal(expected, Markers(document));
        using WordDocument nativeControl = WordDocument.Load(Path.Combine(Path.GetDirectoryName(path)!, "word-lists-native.docx"));
        Assert.Equal(expected, Markers(nativeControl));

        document.Paragraphs.Single(paragraph => paragraph.Text == "Plain separator").SetText("Edited separator");
        using var output = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(output);
        Assert.Equal(expected, Markers(reopened));
        Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "Edited separator");
        MainDocumentPart main = reopened.OpenXmlDocument.MainDocumentPart!;
        Numbering numbering = main.NumberingDefinitionsPart!.Numbering!;
        Assert.Equal(5, numbering.Elements<AbstractNum>().Count());
        Assert.Equal(5, numbering.Elements<NumberingInstance>().Count());
        AbstractNum outline = numbering.Elements<AbstractNum>().Single(definition => definition.AbstractNumberId!.Value == 0);
        Assert.Equal(Enumerable.Range(0, 9), outline.Elements<Level>().Select(level => level.LevelIndex!.Value));
        Assert.Contains(numbering.Descendants<StartNumberingValue>(), start => start.Val!.Value == 7);
        Level bullet = Assert.Single(numbering.Descendants<Level>(), level => level.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet);
        Assert.Equal("•", bullet.LevelText!.Val!.Value);
        Assert.Equal("Arial", bullet.NumberingSymbolRunProperties!.RunFonts!.Ascii!.Value);
        TabStop bulletTab = Assert.Single(bullet.PreviousParagraphProperties!.Tabs!.Elements<TabStop>());
        Assert.Equal(TabStopValues.Number, bulletTab.Val!.Value);
        Assert.Equal(360, bulletTab.Position!.Value);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(reopened.OpenXmlDocument));
    }

    [Fact]
    public void StrictWordMlListsResolveStylesContinuationCancellationAndStartOverrides() {
        string styles = "<w:styles><w:style w:type=\"paragraph\" w:styleId=\"ListStyle\"><w:name w:val=\"List style\"/>" +
            "<w:pPr><w:listPr><w:ilfo w:val=\"12\"/></w:listPr></w:pPr></w:style></w:styles>";
        string lists = "<w:lists><w:list w:ilfo=\"12\"><w:ilst w:val=\"4\"/><w:lvlOverride w:ilvl=\"0\">" +
            "<w:startOverride w:val=\"7\"/><w:lvl w:ilvl=\"0\"><w:start w:val=\"2\"/><w:nfc w:val=\"0\"/>" +
            "<w:lvlText w:val=\"(%1)\"/></w:lvl></w:lvlOverride></w:list>" + Definition +
            "<w:list w:ilfo=\"11\"><w:ilst w:val=\"4\"/></w:list></w:lists>";
        string body = Paragraph("First", 11) + "<w:p><w:r><w:t>Plain</w:t></w:r></w:p>" + Paragraph("Continued", 11) +
            "<w:p><w:pPr><w:pStyle w:val=\"ListStyle\"/></w:pPr><w:r><w:t>Styled seven</w:t></w:r></w:p>" +
            Paragraph("Eight", 12) + Paragraph("Cancelled", 0);
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(Wrap(styles + lists, body)));
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        Assert.False(report.HasLoss);
        Assert.Equal(new[] { "First|0|1.", "Continued|0|2.", "Styled seven|0|(7)", "Eight|0|(8)" }, Markers(document));
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument package = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
        NumberingInstance restarted = package.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.Elements<NumberingInstance>()
            .Single(instance => instance.NumberID!.Value == 12);
        Assert.Equal(4, restarted.AbstractNumId!.Val!.Value);
        Assert.Equal(7, Assert.Single(restarted.Elements<LevelOverride>()).StartOverrideNumberingValue!.Val!.Value);
        Assert.Equal("(%1)", Assert.Single(restarted.Elements<LevelOverride>()).GetFirstChild<Level>()!.LevelText!.Val!.Value);
        Assert.Equal(0, package.MainDocumentPart.Document!.Body!.Elements<Paragraph>().Last().ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value);
    }

    [Fact]
    public void UnsupportedWordMlPictureListMarkersRemainStrictLosses() {
        string lists = "<w:lists>" + Definition.Replace("<w:lvlText", "<w:lvlPicBulletId w:val=\"3\"/><w:lvlText") +
            "<w:list w:ilfo=\"11\"><w:ilst w:val=\"4\"/></w:list></w:lists>";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(Wrap(lists, Paragraph("Picture list text", 11))));
        using WordDocument document = WordDocument.Load(input, null, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Location == (Word + "lvlPicBulletId").ToString() &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal("Picture list text", document.Paragraphs.Single().Text);
        Assert.Empty(document.OpenXmlDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.Descendants<LevelPictureBulletId>());
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true }));
    }

    [Theory]
    [InlineData("definition")]
    [InlineData("instance")]
    [InlineData("level")]
    public void UnresolvedWordMlListsRemainReportedLossesWithoutInventedBindings(string boundary) {
        string lists = "<w:lists>" + Definition + "<w:list w:ilfo=\"11\"><w:ilst w:val=\"" +
            (boundary == "definition" ? "99" : "4") + "\"/></w:list></w:lists>";
        string xml = Wrap(lists, Paragraph("Retained list text", boundary == "instance" ? 99 : 11, boundary == "level" ? 1 : 0));
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        using WordDocument document = WordDocument.Load(input, null, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Message.Contains("unresolved"));
        Assert.Empty(Markers(document));
        Assert.Equal("Retained list text", document.Paragraphs.Single().Text);
        Assert.Null(document.OpenXmlDocument.MainDocumentPart!.Document!.Body!.GetFirstChild<Paragraph>()!.ParagraphProperties?.NumberingProperties);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true, ReportUnsupportedContent = false }));
    }

    [Theory]
    [InlineData("duplicate-definition")]
    [InlineData("duplicate-instance")]
    [InlineData("duplicate-level")]
    [InlineData("level-out-of-range")]
    [InlineData("mismatched-override")]
    [InlineData("missing-definition-reference")]
    public void AmbiguousOrInvalidWordMlListIdentitiesAreRejected(string boundary) {
        string definition = boundary switch {
            "duplicate-definition" => Definition + Definition,
            "duplicate-level" => Definition.Replace("</w:listDef>", "<w:lvl w:ilvl=\"0\"><w:lvlText w:val=\"%1)\"/></w:lvl></w:listDef>"),
            "level-out-of-range" => Definition.Replace("w:ilvl=\"0\"", "w:ilvl=\"9\""),
            _ => Definition
        };
        string instance = "<w:list w:ilfo=\"11\"><w:ilst w:val=\"4\"/></w:list>";
        if (boundary == "duplicate-instance") instance += instance;
        if (boundary == "missing-definition-reference") instance = "<w:list w:ilfo=\"11\"/>";
        if (boundary == "mismatched-override") instance = "<w:list w:ilfo=\"11\"><w:ilst w:val=\"4\"/>" +
            "<w:lvlOverride w:ilvl=\"0\"><w:lvl w:ilvl=\"1\"><w:lvlText w:val=\"%2.\"/></w:lvl></w:lvlOverride></w:list>";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(Wrap("<w:lists>" + definition + instance + "</w:lists>", Paragraph("Retained", 11))));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(input));
        Assert.Contains("list", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void WordMlNumberingPartAndRelationshipConsumeTheReconstructionBudget(bool parts) {
        string xml = Wrap("<w:lists>" + Definition + "<w:list w:ilfo=\"11\"><w:ilst w:val=\"4\"/></w:list></w:lists>", Paragraph("Budget", 11));
        var policy = new OfficePackageSecurityOptions();
        if (parts) policy.MaxPartCount = 4;
        else policy.MaxRelationships = 1;
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(input, new WordLoadOptions { PackageSecurity = policy }));
        Assert.Contains("relationship or part limit", error.Message);
        if (parts) policy.MaxPartCount++;
        else policy.MaxRelationships++;
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { PackageSecurity = policy, FailOnLoss = true });
        Assert.Equal(new[] { "Budget|0|1." }, Markers(document));
    }

    private static string[] Markers(WordDocument document) => document.Inspect(includeImageBytes: false).Sections
        .SelectMany(section => section.Elements).OfType<WordParagraphInfo>().Where(paragraph => paragraph.IsListItem)
        .Select(paragraph => paragraph.Text + "|" + paragraph.ListLevel + "|" + paragraph.ListMarker).ToArray();

    private const string Definition = "<w:listDef w:listDefId=\"4\"><w:plt w:val=\"SingleLevel\"/><w:lvl w:ilvl=\"0\">" +
        "<w:start w:val=\"1\"/><w:nfc w:val=\"0\"/><w:suff w:val=\"Tab\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/>" +
        "<w:pPr><w:tabs><w:tab w:val=\"list\" w:pos=\"360\"/></w:tabs><w:ind w:left=\"360\" w:hanging=\"360\"/></w:pPr>" +
        "<w:rPr><w:rFonts w:ascii=\"Arial\"/><w:color w:val=\"003366\"/></w:rPr></w:lvl></w:listDef>";

    private static string Paragraph(string text, int number, int level = 0) => "<w:p><w:pPr><w:listPr><w:ilvl w:val=\"" +
        level + "\"/><w:ilfo w:val=\"" + number + "\"/></w:listPr></w:pPr><w:r><w:t>" + text + "</w:t></w:r></w:p>";

    private static string Wrap(string declarations, string body) => "<w:wordDocument xmlns:w=\"" + Word + "\">" + declarations +
        "<w:body>" + body + "</w:body></w:wordDocument>";
}
