using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordMlPropertyContextTests {
    private static readonly XNamespace Word = "http://schemas.microsoft.com/office/word/2003/wordml";
    private static readonly XNamespace Modern = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData("run")]
    [InlineData("paragraph")]
    [InlineData("section")]
    [InlineData("paragraph-section")]
    [InlineData("borders")]
    [InlineData("font")]
    public void MappedPropertyValuesUseTheDestinationSequence(string context) {
        string content = context switch {
            "run" => "<w:p><w:r><w:rPr><w:color w:val=\"112233\"/><w:b/><w:rFonts w:ascii=\"Arial\"/></w:rPr><w:t>Retained</w:t></w:r></w:p>",
            "paragraph" => "<w:p><w:pPr><w:jc w:val=\"center\"/><w:spacing w:before=\"120\"/></w:pPr><w:r><w:t>Retained</w:t></w:r></w:p>",
            "section" => "<w:p><w:r><w:t>Retained</w:t></w:r></w:p><w:sectPr><w:cols w:num=\"2\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>",
            "paragraph-section" => "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:cols w:num=\"2\"/></w:sectPr></w:pPr><w:r><w:t>Retained</w:t></w:r></w:p>",
            "borders" => "<w:p><w:pPr><w:pBdr><w:bottom w:val=\"single\"/><w:top w:val=\"double\"/></w:pBdr></w:pPr><w:r><w:t>Retained</w:t></w:r></w:p>",
            _ => "<w:p><w:r><w:t>Retained</w:t></w:r></w:p>"
        };
        string fonts = context == "font" ? "<w:fonts><w:font w:name=\"Arial\"><w:pitch w:val=\"variable\"/><w:family w:val=\"swiss\"/></w:font></w:fonts>" : "";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(Source(content, fonts)));
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        Assert.False(report.HasLoss);
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        Assert.Equal("Retained", saved.MainDocumentPart!.Document.InnerText);
        XElement root = XElement.Parse(saved.MainDocumentPart.Document.OuterXml);
        if (context == "run") {
            XElement properties = Assert.Single(root.Descendants(Modern + "rPr"));
            Assert.Equal(new[] { "rFonts", "b", "color" }, properties.Elements().Select(element => element.Name.LocalName));
            Assert.Equal("Arial", (string?)properties.Element(Modern + "rFonts")!.Attribute(Modern + "ascii"));
            Assert.Equal("112233", (string?)properties.Element(Modern + "color")!.Attribute(Modern + "val"));
        } else if (context == "paragraph") {
            XElement properties = Assert.Single(root.Descendants(Modern + "pPr"));
            Assert.Equal(new[] { "spacing", "jc" }, properties.Elements().Select(element => element.Name.LocalName));
            Assert.Equal("120", (string?)properties.Element(Modern + "spacing")!.Attribute(Modern + "before"));
            Assert.Equal("center", (string?)properties.Element(Modern + "jc")!.Attribute(Modern + "val"));
        } else if (context is "section" or "paragraph-section") {
            XElement properties = Assert.Single(root.Descendants(Modern + "sectPr"),
                section => section.Parent!.Name == Modern + (context == "paragraph-section" ? "pPr" : "body"));
            Assert.Equal(new[] { "pgSz", "cols" }, properties.Elements().Select(element => element.Name.LocalName));
            Assert.Equal("12240", (string?)properties.Element(Modern + "pgSz")!.Attribute(Modern + "w"));
            Assert.Equal("2", (string?)properties.Element(Modern + "cols")!.Attribute(Modern + "num"));
        } else if (context == "borders") {
            XElement properties = Assert.Single(root.Descendants(Modern + "pBdr"));
            Assert.Equal(new[] { "top", "bottom" }, properties.Elements().Select(element => element.Name.LocalName));
            Assert.Equal("double", (string?)properties.Element(Modern + "top")!.Attribute(Modern + "val"));
            Assert.Equal("single", (string?)properties.Element(Modern + "bottom")!.Attribute(Modern + "val"));
        } else {
            XElement font = XElement.Parse(Assert.Single(saved.MainDocumentPart.FontTablePart!.Fonts.Elements<Font>(), font => font.Name!.Value == "Arial").OuterXml);
            Assert.Equal(new[] { "family", "pitch" }, font.Elements().Select(element => element.Name.LocalName));
            Assert.Equal("swiss", (string?)font.Element(Modern + "family")!.Attribute(Modern + "val"));
            Assert.Equal("variable", (string?)font.Element(Modern + "pitch")!.Attribute(Modern + "val"));
        }
    }

    [Fact]
    public void ChoiceAndRepeatedPropertiesRetainEveryValueAndSourceOrder() {
        const string styles = "<w:styles><w:style w:type=\"table\" w:styleId=\"Example\"><w:name w:val=\"Example\"/>" +
            "<w:tblStylePr w:type=\"first-row\"><w:trPr><w:jc w:val=\"right\"/><w:tblHeader/></w:trPr></w:tblStylePr></w:style></w:styles>";
        const string body = "<w:p><w:pPr><w:tabs><w:tab w:val=\"right\" w:pos=\"1440\"/><w:tab w:val=\"left\" w:pos=\"720\"/></w:tabs></w:pPr></w:p>" +
            "<w:tbl><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr>" +
            "<w:trPr><w:jc w:val=\"center\"/><w:trHeight w:val=\"240\"/><w:cantSplit/></w:trPr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>" +
            "<w:sectPr><w:cols w:equalWidth=\"off\" w:num=\"2\"><w:col w:w=\"2000\" w:space=\"360\"/><w:col w:w=\"4000\" w:space=\"0\"/></w:cols></w:sectPr>";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(Source(body, styles)));
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        Assert.False(report.HasLoss);
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        XElement root = XElement.Parse(saved.MainDocumentPart!.Document.OuterXml);
        Assert.Equal(new[] { "1440", "720" }, root.Descendants(Modern + "tab").Select(tab => (string?)tab.Attribute(Modern + "pos")));
        Assert.Equal(new[] { "1000", "3000" }, root.Descendants(Modern + "gridCol").Select(column => (string?)column.Attribute(Modern + "w")));
        Assert.Equal(new[] { "2000", "4000" }, root.Descendants(Modern + "col").Select(column => (string?)column.Attribute(Modern + "w")));
        XElement row = Assert.Single(root.Descendants(Modern + "trPr"));
        Assert.Equal(new[] { "jc", "trHeight", "cantSplit" }, row.Elements().Select(element => element.Name.LocalName));
        Assert.Equal("center", (string?)row.Element(Modern + "jc")!.Attribute(Modern + "val"));
        Assert.Equal("240", (string?)row.Element(Modern + "trHeight")!.Attribute(Modern + "val"));
        XElement style = XElement.Parse(Assert.Single(saved.MainDocumentPart.StyleDefinitionsPart!.Styles.Elements<Style>(), item => item.StyleId!.Value == "Example").OuterXml);
        Assert.Equal(new[] { "jc", "tblHeader" }, Assert.Single(style.Descendants(Modern + "trPr")).Elements().Select(element => element.Name.LocalName));
    }

    [Theory]
    [InlineData("paragraph-style-run", "rPr", "spacing")]
    [InlineData("paragraph-style-section", "sectPr", "spacing")]
    [InlineData("conditional-style-section", "sectPr", "spacing")]
    [InlineData("character-style-highlight", "highlight", "b")]
    [InlineData("table-style-width", "tblW", "jc")]
    [InlineData("conditional-row-grid", "gridBefore", "jc")]
    [InlineData("body-table-band", "tblStyleRowBandSize", "tblW")]
    [InlineData("duplicate-run", "b", "color")]
    public void UnsupportedDestinationPropertiesProduceLossInsteadOfInvalidXml(string context, string omitted, string retained) {
        string properties = context switch {
            "paragraph-style-run" => "<w:pPr><w:spacing w:before=\"120\"/><w:rPr><w:b/></w:rPr></w:pPr>",
            "paragraph-style-section" => "<w:pPr><w:spacing w:before=\"120\"/><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr>",
            "conditional-style-section" => "<w:tblStylePr w:type=\"firstRow\"><w:pPr><w:spacing w:before=\"120\"/><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr></w:tblStylePr>",
            "character-style-highlight" => "<w:rPr><w:b/><w:highlight w:val=\"yellow\"/></w:rPr>",
            "table-style-width" => "<w:tblPr><w:tblW w:w=\"2000\" w:type=\"dxa\"/><w:jc w:val=\"center\"/></w:tblPr>",
            _ => "<w:tblStylePr w:type=\"first-row\"><w:trPr><w:gridBefore w:val=\"0\"/><w:jc w:val=\"center\"/></w:trPr></w:tblStylePr>"
        };
        string kind = context.StartsWith("paragraph-style", StringComparison.Ordinal) ? "paragraph" : context == "character-style-highlight" ? "character" : "table";
        string styles = "<w:styles><w:style w:type=\"" + kind + "\" w:styleId=\"Example\"><w:name w:val=\"Example\"/>" + properties + "</w:style></w:styles>";
        string body = "<w:p><w:r><w:t>Retained</w:t></w:r></w:p>";
        if (context == "body-table-band") {
            styles = "";
            body += "<w:tbl><w:tblPr><w:tblStyleRowBandSize w:val=\"2\"/><w:tblW w:w=\"2000\" w:type=\"dxa\"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>";
        } else if (context == "duplicate-run") {
            styles = "";
            body = "<w:p><w:r><w:rPr><w:b/><w:b w:val=\"off\"/><w:color w:val=\"112233\"/></w:rPr><w:t>Retained</w:t></w:r></w:p>";
        }
        byte[] source = Encoding.UTF8.GetBytes(Source(body, styles));
        using var input = new MemoryStream(source);
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions(), out OfficeConversionReport report);
        Assert.True(report.HasLoss);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "WORDML2003_OMITTED" && diagnostic.Location == (Word + omitted).ToString());
        if (omitted == "sectPr") Assert.Single(report.Diagnostics);
        using var bytes = new MemoryStream(document.ToBytes(new WordSaveOptions { AllowLoss = true }));
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        Assert.Equal("Retained", saved.MainDocumentPart!.Document.InnerText);
        XElement target = context.IndexOf("style", StringComparison.Ordinal) >= 0 || context == "conditional-row-grid"
            ? XElement.Parse(Assert.Single(saved.MainDocumentPart.StyleDefinitionsPart!.Styles.Elements<Style>(), item => item.StyleId!.Value == "Example").OuterXml)
            : XElement.Parse(saved.MainDocumentPart.Document.OuterXml);
        if (context == "duplicate-run") Assert.Single(target.Descendants(Modern + omitted));
        else Assert.Empty(target.Descendants(Modern + omitted));
        Assert.Single(target.Descendants(Modern + retained));
        if (omitted == "sectPr") Assert.Equal("120", (string?)target.Descendants(Modern + retained).Single().Attribute(Modern + "before"));
        using var strict = new MemoryStream(source);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(strict, new WordLoadOptions { FailOnLoss = true }));
    }

    [Theory]
    [InlineData("cols", "two", "num")]
    [InlineData("cols", "46", "num")]
    [InlineData("tab", "31681", "pos")]
    [InlineData("pgSz", "12pt", "w")]
    [InlineData("pgSz", "31681", "w")]
    [InlineData("color", "ZZZZZZ", "val")]
    public void UnsupportedScalarsAreReportedWithoutDiscardingValidSiblings(string property, string value, string attribute) {
        string body = property switch {
            "cols" => "<w:p><w:r><w:t>Retained</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:cols w:num=\"" + value + "\"/></w:sectPr>",
            "pgSz" => "<w:p><w:r><w:t>Retained</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"" + value + "\" w:h=\"15840\"/><w:cols w:num=\"2\"/></w:sectPr>",
            "tab" => "<w:p><w:pPr><w:pBdr/><w:tabs><w:tab w:val=\"left\" w:pos=\"720\"/><w:tab w:val=\"right\" w:pos=\"" + value + "\"/></w:tabs></w:pPr><w:r><w:t>Retained</w:t></w:r></w:p>",
            _ => "<w:p><w:r><w:rPr><w:rFonts w:ascii=\"  Times New Roman  \"/><w:b/><w:color w:val=\"" + value + "\"/></w:rPr><w:t>Retained</w:t></w:r></w:p>"
        };
        byte[] source = Encoding.UTF8.GetBytes(Source(body));
        using var input = new MemoryStream(source);
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions(), out OfficeConversionReport report);
        Assert.True(report.HasLoss);
        OfficeDiagnostic diagnostic = Assert.Single(report.Diagnostics);
        Assert.Equal("WORDML2003_OMITTED", diagnostic.Code);
        Assert.Equal((Word + property).ToString(), diagnostic.Location);
        Assert.Contains(attribute + "='" + value + "'", diagnostic.Message);
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        Assert.Equal("Retained", saved.MainDocumentPart!.Document.InnerText);
        XElement root = XElement.Parse(saved.MainDocumentPart.Document.OuterXml);
        Assert.DoesNotContain(root.Descendants(Modern + property).Attributes(), item => item.Value == value);
        if (property == "cols") Assert.Equal("12240", (string?)Assert.Single(root.Descendants(Modern + "pgSz")).Attribute(Modern + "w"));
        else if (property == "pgSz") Assert.Equal("2", (string?)Assert.Single(root.Descendants(Modern + "cols")).Attribute(Modern + "num"));
        else if (property == "tab") {
            Assert.Equal("720", (string?)Assert.Single(root.Descendants(Modern + "tab")).Attribute(Modern + "pos"));
            Assert.Empty(Assert.Single(root.Descendants(Modern + "pBdr")).Elements());
        } else {
            Assert.Single(root.Descendants(Modern + "b"));
            Assert.Equal("  Times New Roman  ", (string?)Assert.Single(root.Descendants(Modern + "rFonts")).Attribute(Modern + "ascii"));
        }
        using var strict = new MemoryStream(source);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(strict, new WordLoadOptions { FailOnLoss = true }));
    }

    [Fact]
    public void SupportedScalarLexemesPreserveStringsAndNormalizeOnlyHexWhitespace() {
        const string prefix = "<w:fonts><w:font w:name=\"  Times New Roman  \"><w:panose-1 w:val=\" 02020603050405020304 \"/><w:charset w:val=\" 00 \"/><w:family w:val=\"roman\"/><w:pitch w:val=\"variable\"/></w:font></w:fonts>" +
            "<w:styles><w:style w:type=\"paragraph\" w:styleId=\"ParagraphFormat\"><w:name w:val=\"Paragraph format\"/></w:style><w:style w:type=\"character\" w:styleId=\"CharacterFormat\"><w:name w:val=\"Character format\"/></w:style></w:styles>";
        const string body = "<w:p><w:pPr><w:pStyle w:val=\"ParagraphFormat\"/><w:pBdr><w:top w:val=\"single\" w:sz=\"8\" w:color=\" aBcD12 \"/></w:pBdr>" +
            "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\" 001122 \"/><w:tabs><w:tab w:val=\"left\" w:pos=\"31680\"/></w:tabs><w:spacing w:line=\"-240\"/><w:ind w:left=\"-720\"/></w:pPr>" +
            "<w:r><w:rPr><w:rStyle w:val=\"CharacterFormat\"/><w:rFonts w:ascii=\"  Times New Roman  \"/><w:color w:val=\" FF0000 \"/><w:sz w:val=\" 22 \"/><w:u w:val=\"single\" w:color=\" FF0000 \"/><w:lang w:val=\"en-US\" w:fareast=\"ja-JP\" w:bidi=\"ar-SA\"/></w:rPr><w:t>Retained</w:t></w:r></w:p>" +
            "<w:tbl><w:tblPr><w:tblLook w:val=\" 04A0 \"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:tcPr><w:cnfStyle w:val=\"100000000000\"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>" +
            "<w:sectPr><w:pgSz w:w=\"31680\" w:h=\"15840\"/><w:cols w:num=\"45\"/></w:sectPr>";
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(Source(body, prefix)));
        using WordDocument document = WordDocument.Load(source, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        Assert.False(report.HasLoss);
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        XElement root = XElement.Parse(saved.MainDocumentPart!.Document.OuterXml);
        Assert.Equal("Retained", saved.MainDocumentPart.Document.InnerText);
        Assert.Equal("aBcD12", (string?)Assert.Single(root.Descendants(Modern + "top")).Attribute(Modern + "color"));
        Assert.Equal("FF0000", (string?)Assert.Single(root.Descendants(Modern + "color")).Attribute(Modern + "val"));
        XElement underline = Assert.Single(root.Descendants(Modern + "u"));
        Assert.Equal("single", (string?)underline.Attribute(Modern + "val"));
        Assert.Equal("FF0000", (string?)underline.Attribute(Modern + "color"));
        Assert.Equal("001122", (string?)Assert.Single(root.Descendants(Modern + "shd")).Attribute(Modern + "fill"));
        Assert.Equal("auto", (string?)Assert.Single(root.Descendants(Modern + "shd")).Attribute(Modern + "color"));
        Assert.Equal("  Times New Roman  ", (string?)Assert.Single(root.Descendants(Modern + "rFonts")).Attribute(Modern + "ascii"));
        Assert.Equal("ParagraphFormat", (string?)Assert.Single(root.Descendants(Modern + "pStyle")).Attribute(Modern + "val"));
        Assert.Equal("CharacterFormat", (string?)Assert.Single(root.Descendants(Modern + "rStyle")).Attribute(Modern + "val"));
        Assert.Equal("en-US", (string?)Assert.Single(root.Descendants(Modern + "lang")).Attribute(Modern + "val"));
        Assert.Equal("-240", (string?)Assert.Single(root.Descendants(Modern + "spacing")).Attribute(Modern + "line"));
        Assert.Equal("-720", (string?)Assert.Single(root.Descendants(Modern + "ind")).Attribute(Modern + "start"));
        Assert.Equal("31680", (string?)Assert.Single(root.Descendants(Modern + "tab")).Attribute(Modern + "pos"));
        Assert.Equal("45", (string?)Assert.Single(root.Descendants(Modern + "cols")).Attribute(Modern + "num"));
        Assert.Equal("04A0", (string?)Assert.Single(root.Descendants(Modern + "tblLook")).Attribute(Modern + "val"));
        Assert.Equal("100000000000", (string?)Assert.Single(root.Descendants(Modern + "cnfStyle")).Attribute(Modern + "val"));
        Assert.Single(Assert.Single(root.Descendants(Modern + "tblGrid")).Elements(Modern + "gridCol"));
        XElement font = XElement.Parse(Assert.Single(saved.MainDocumentPart.FontTablePart!.Fonts.Elements<Font>()).OuterXml);
        Assert.Equal("  Times New Roman  ", (string?)font.Attribute(Modern + "name"));
        Assert.Equal("02020603050405020304", (string?)font.Element(Modern + "panose1")!.Attribute(Modern + "val"));
        Assert.Equal("00", (string?)font.Element(Modern + "charset")!.Attribute(Modern + "val"));
    }

    [Fact]
    public void MappedPropertyValidationRejectsACappedInvalidGroupBeforePartialEmission() {
        int count = new OpenXmlValidator(FileFormatVersions.Microsoft365).MaxNumberOfErrors + 1;
        string grid = string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"invalid\"/>", count));
        string xml = Source("<w:tbl><w:tblGrid>" + grid + "</w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>");
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(input));
        Assert.Contains("SDK validation error limit", error.Message);
        Assert.True(input.CanRead);
    }

    private static string Source(string body, string prefix = "") =>
        "<w:wordDocument xmlns:w=\"" + Word + "\">" + prefix + "<w:body>" + body + "</w:body></w:wordDocument>";
}
