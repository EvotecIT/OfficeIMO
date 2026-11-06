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
    private static readonly byte[] Pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=");

    [Theory]
    [InlineData("shape", "style", "o:style", "width:12pt;height:6pt")]
    [InlineData("shape", "o:spt", "spt", "75")]
    [InlineData("imagedata", "o:title", "title", "Caption")]
    [InlineData("lock", "v:ext", "ext", "edit")]
    [InlineData("shape", null, "behindDoc", "t")]
    [InlineData("shape", null, "o:behindDoc", "t")]
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
    public void SupportedEnumAndBooleanTokensUseTheSdkVocabularyBeforeEmission(
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
        if (element == "imagedata") {
            Assert.Null(result.Attribute("src"));
            Assert.NotNull(result.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")));
        }
        return result;
    }

    private static string Source(string element, string? attribute, string value) {
        var shape = new XElement(Vml + "shape", new XAttribute("id", "Picture1"), new XAttribute("style", "width:12pt;height:6pt"));
        XElement target = element == "shape" ? shape : new XElement((element == "lock" ? Office : Vml) + element);
        if (element == "imagedata") target.SetAttributeValue("src", "wordml://pixel.png");
        if (element == "lock") target.SetAttributeValue("aspectratio", "t");
        if (attribute != null) target.SetAttributeValue(Name(attribute), value);
        if (target != shape) shape.Add(target);
        return new XElement(Word + "wordDocument", new XAttribute(XNamespace.Xmlns + "w", Word),
            new XAttribute(XNamespace.Xmlns + "v", Vml), new XAttribute(XNamespace.Xmlns + "o", Office),
            new XElement(Word + "body", new XElement(Word + "p", new XElement(Word + "r", new XElement(Word + "pict",
                new XElement(Word + "binData", new XAttribute(Word + "name", "wordml://pixel.png"), Convert.ToBase64String(Pixel)), shape)))))
            .ToString(SaveOptions.DisableFormatting);
    }

    private static XName Name(string name) => name.StartsWith("o:", StringComparison.Ordinal) ? Office + name.Substring(2)
        : name.StartsWith("v:", StringComparison.Ordinal) ? Vml + name.Substring(2) : XName.Get(name);
}
