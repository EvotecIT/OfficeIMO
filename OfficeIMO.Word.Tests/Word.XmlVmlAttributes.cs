using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordMlVmlAttributeTests {
    private static readonly XNamespace Word = "http://schemas.microsoft.com/office/word/2003/wordml";
    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace Office = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace WordSupplement = "http://schemas.microsoft.com/office/word/2003/wordml/sp2";
    private static readonly XNamespace Auxiliary = "http://schemas.microsoft.com/office/word/2003/auxHint";
    private static readonly byte[] Pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=");

    [Theory]
    [InlineData("shape", "style", "o:style", "width:12pt;height:6pt")]
    [InlineData("shape", "o:spt", "spt", "75")]
    [InlineData("imagedata", "o:title", "title", "Caption")]
    [InlineData("lock", "v:ext", "ext", "edit")]
    [InlineData("shape", null, "behindDoc", "t")]
    [InlineData("shape", null, "o:behindDoc", "t")]
    [InlineData("lock", null, "wsp:rsidR", "00000001")]
    [InlineData("lock", null, "wx:bdrwidth", "10")]
    [InlineData("shape", null, "wsp:rsidR", "00000001")]
    [InlineData("shape", null, "wx:bdrwidth", "10")]
    [InlineData("lock", "xml:space", "space", "preserve")]
    [InlineData("lock", "xml:space", "space", "default")]
    public void SupportedAttributeQNamesSurviveWhileWrongNamespacesAreReportedAndOmitted(
        string element, string? qualified, string wrong, string value) {
        string validSource = Source(element, qualified, value);
        XElement valid = LoadSave(validSource, element, strict: true, out OfficeConversionReport validReport);
        Assert.False(validReport.HasLoss);
        if (qualified != null) Assert.Equal(value, (string?)valid.Attribute(Name(qualified)));

        string unsupportedSource = Source(element, wrong, value);
        XElement projected = LoadSave(unsupportedSource, element, strict: false, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Location == Name(wrong).ToString() &&
            diagnostic.Code == "WORDML2003_OMITTED" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Null(projected.Attribute(Name(wrong)));
        using var strict = new MemoryStream(Encoding.UTF8.GetBytes(unsupportedSource));
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(strict, new WordLoadOptions { FailOnLoss = true }));
    }

    [Theory]
    [InlineData("lock", "v:ext", "edit", "unsupported")]
    [InlineData("stroke", "joinstyle", "miter", "unsupported")]
    [InlineData("stroke", "endcap", "round", "unsupported")]
    [InlineData("path", "o:connecttype", "rect", "unsupported")]
    [InlineData("fill", "type", "solid", "unsupported")]
    [InlineData("fill", "method", "sigma", "unsupported")]
    [InlineData("shape", "filled", "t", "unsupported")]
    [InlineData("shape", "filled", "true", "1")]
    [InlineData("shape", "filled", "false", "0")]
    [InlineData("lock", "aspectratio", "f", "unsupported")]
    [InlineData("shape", "o:spt", "75", "not-a-number")]
    [InlineData("shape", "o:spt", "0", "-1")]
    [InlineData("shape", "o:spt", "75", "2147483648")]
    [InlineData("shapetype", "o:spt", "202", "203")]
    [InlineData("fill", "angle", "45.5", "not-a-number")]
    [InlineData("fill", "angle", "32767", "32768")]
    [InlineData("fill", "angle", "-32767", "-32768")]
    [InlineData("fill", "angle", "45.5", "79228162514264337593543950336")]
    [InlineData("fill", "focus", "100%", "not-a-percent")]
    [InlineData("fill", "focus", "-99%", "101%")]
    [InlineData("fill", "focus", "", "100.1%")]
    [InlineData("imagedata", "blacklevel", "0.5", "0.5001")]
    [InlineData("imagedata", "blacklevel", "-0.5", "1")]
    [InlineData("imagedata", "blacklevel", "32768f", "1")]
    [InlineData("imagedata", "blacklevel", "-0.25", "-1")]
    [InlineData("imagedata", "blacklevel", "-0.5", "-0.5001")]
    [InlineData("imagedata", "blacklevel", "+0.5", "+0.5001")]
    [InlineData("imagedata", "blacklevel", "-32768f", "-1")]
    [InlineData("stroke", "weight", "20116800", "20116801")]
    [InlineData("stroke", "weight", "1pt", "20116801")]
    [InlineData("stroke", "weight", "-0", "-5")]
    [InlineData("stroke", "weight", "+20116800", "+20116801")]
    [InlineData("stroke", "weight", "1pt", "-0.0001")]
    public void SupportedScalarValuesUseTheSdkVocabularyAndDeclaredBoundsBeforeEmission(
        string element, string attribute, string validValue, string invalidValue) {
        XElement valid = LoadSave(Source(element, attribute, validValue), element, strict: true, out OfficeConversionReport validReport);
        Assert.False(validReport.HasLoss);
        Assert.Equal(validValue, (string?)valid.Attribute(Name(attribute)));

        string invalidSource = Source(element, attribute, invalidValue);
        XElement projected = LoadSave(invalidSource, element, strict: false, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Location == Name(attribute).ToString() &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission && diagnostic.Message.Contains(invalidValue, StringComparison.Ordinal));
        Assert.Null(projected.Attribute(Name(attribute)));
        using var strict = new MemoryStream(Encoding.UTF8.GetBytes(invalidSource));
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(strict, new WordLoadOptions { FailOnLoss = true }));
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("shapetype")]
    public void DuplicateVmlIdentifiersWithinOneStoryAreRejected(string element) {
        XElement source = XElement.Parse(Source(element, "id", "Duplicate"));
        XElement original = source.Descendants(Vml + element).Single();
        original.AddAfterSelf(new XElement(original));
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(source.ToString(SaveOptions.DisableFormatting)));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(input));
        Assert.Contains("duplicate VML " + element + " identifier within one story", error.Message);
    }

    [Theory]
    [InlineData("shape", 0, false)]
    [InlineData("shape", 128, false)]
    [InlineData("shape", 129, false)]
    [InlineData("shape", 128, true)]
    [InlineData("shapetype", 0, false)]
    [InlineData("shapetype", 128, false)]
    [InlineData("shapetype", 129, false)]
    [InlineData("shapetype", 128, true)]
    public void VmlFormulaGroupsRespectTheQualified128FormulaLimit(string element, int count, bool unknownChild) {
        XNamespace unknown = "urn:fixture:unknown";
        XElement source = XElement.Parse(Source(element, null, string.Empty));
        var formulas = new XElement(Vml + "formulas", Enumerable.Range(0, count)
            .Select(index => new XElement(Vml + "f", new XAttribute("eqn", "val " + index))));
        if (unknownChild) formulas.Add(new XElement(unknown + "f", new XAttribute("eqn", "val 999")));
        source.Descendants(Vml + element).Single().Add(formulas);
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(source.ToString(SaveOptions.DisableFormatting)));
        if (count > 128) {
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(input));
            Assert.Contains("at most 128 VML formulas", error.Message);
            return;
        }

        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { FailOnLoss = !unknownChild }, out OfficeConversionReport report);
        Assert.Equal(unknownChild, report.HasLoss);
        if (unknownChild) Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Location == (unknown + "f").ToString() &&
            diagnostic.Code == "WORDML2003_OMITTED");
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        XElement output = Assert.Single(XElement.Parse(saved.MainDocumentPart!.Document.OuterXml).Descendants(Vml + "formulas"));
        Assert.Equal(Enumerable.Range(0, count).Select(index => "val " + index), output.Elements(Vml + "f").Select(formula => (string?)formula.Attribute("eqn")));
        Assert.Empty(output.Elements(unknown + "f"));
    }

    [Fact]
    public void VmlIdentifierUniquenessAllowsSeparateStoriesTypesAndEmptyIds() {
        XElement source = XElement.Parse(Source("shape", "id", "Shared"));
        XElement shape = source.Descendants(Vml + "shape").Single();
        shape.AddBeforeSelf(new XElement(Vml + "shapetype", new XAttribute("id", "Shared")));
        shape.AddAfterSelf(new XElement(Vml + "shape", new XAttribute("id", "")),
            new XElement(Vml + "shape", new XAttribute("id", "")));
        source.Element(Word + "body")!.Add(new XElement(Word + "sectPr", new XElement(Word + "hdr", new XAttribute(Word + "type", "odd"),
            new XElement(Word + "p", new XElement(Word + "r", new XElement(Word + "pict", new XElement(shape)))))));
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(source.ToString(SaveOptions.DisableFormatting)));
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        Assert.False(report.HasLoss);
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        Assert.Equal(new[] { "Shared", "", "" }, saved.MainDocumentPart!.Document.Descendants<DocumentFormat.OpenXml.Vml.Shape>().Select(s => s.Id!.Value));
        Assert.Equal("Shared", Assert.Single(saved.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Vml.Shapetype>()).Id!.Value);
        Assert.Equal("Shared", Assert.Single(Assert.Single(saved.MainDocumentPart.HeaderParts).Header!.Descendants<DocumentFormat.OpenXml.Vml.Shape>()).Id!.Value);
    }

    private static XElement LoadSave(string source, string element, bool strict, out OfficeConversionReport report) {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(source));
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { FailOnLoss = strict }, out report);
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        if (element == "imagedata") {
            using Stream image = Assert.Single(saved.MainDocumentPart!.ImageParts).GetStream();
            using var imageBytes = new MemoryStream();
            image.CopyTo(imageBytes);
            Assert.Equal(Pixel, imageBytes.ToArray());
        } else Assert.Empty(saved.MainDocumentPart!.ImageParts);
        XName target = (element == "lock" ? Office : Vml) + element;
        XElement result = Assert.Single(XElement.Parse(saved.MainDocumentPart!.Document.OuterXml).Descendants(target));
        Assert.Equal("urn:fixture:local", (string?)result.Attribute(XNamespace.Xmlns + "local"));
        if (element == "imagedata") {
            Assert.Null(result.Attribute("src"));
            Assert.NotNull(result.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")));
        }
        return result;
    }

    private static string Source(string element, string? attribute, string value) {
        var shape = new XElement(Vml + "shape", new XAttribute("id", "Picture1"), new XAttribute("style", "width:12pt;height:6pt"));
        XElement target = element == "shape" ? shape : new XElement((element == "lock" ? Office : Vml) + element);
        target.SetAttributeValue(XNamespace.Xmlns + "local", "urn:fixture:local");
        if (element == "imagedata") target.SetAttributeValue("src", "wordml://pixel.png");
        if (element == "lock") target.SetAttributeValue("aspectratio", "t");
        if (element == "shapetype") target.SetAttributeValue("id", "FixtureType");
        if (attribute != null) target.SetAttributeValue(Name(attribute), value);
        if (target != shape && element != "shapetype") shape.Add(target);
        return new XElement(Word + "wordDocument", new XAttribute(XNamespace.Xmlns + "w", Word),
            new XAttribute(XNamespace.Xmlns + "v", Vml), new XAttribute(XNamespace.Xmlns + "o", Office),
            new XElement(Word + "body", new XElement(Word + "p", new XElement(Word + "r", new XElement(Word + "pict",
                new XElement(Word + "binData", new XAttribute(Word + "name", "wordml://pixel.png"), Convert.ToBase64String(Pixel)),
                element == "shapetype" ? target : null, shape)))))
            .ToString(SaveOptions.DisableFormatting);
    }

    private static XName Name(string name) => name.StartsWith("o:", StringComparison.Ordinal) ? Office + name.Substring(2)
        : name.StartsWith("v:", StringComparison.Ordinal) ? Vml + name.Substring(2)
        : name.StartsWith("wsp:", StringComparison.Ordinal) ? WordSupplement + name.Substring(4)
        : name.StartsWith("wx:", StringComparison.Ordinal) ? Auxiliary + name.Substring(3)
        : name.StartsWith("xml:", StringComparison.Ordinal) ? XNamespace.Xml + name.Substring(4) : XName.Get(name);
}
