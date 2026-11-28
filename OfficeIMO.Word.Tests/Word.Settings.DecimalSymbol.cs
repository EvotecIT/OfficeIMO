using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void DocumentSettingsIgnoreExtensionAndAlternateContentChildren() {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        XNamespace extension = "urn:officeimo:settings-extension-test";
        using var source = CreatePackageWithSettingsMarkup(settings => {
            settings.Element(w + "decimalSymbol")?.Remove();
            settings.Add(new XAttribute(XNamespace.Xmlns + "ex", extension.NamespaceName));
            settings.Add(
                new XElement(mc + "AlternateContent",
                    new XElement(mc + "Choice", new XAttribute("Requires", "ex"),
                        new XElement(w + "decimalSymbol", new XAttribute(w + "val", "!"))),
                    new XElement(mc + "Fallback")),
                new XElement(extension + "extLst",
                    new XElement(w + "decimalSymbol", new XAttribute(w + "val", "x"))));
            XElement compatibility = settings.Element(w + "compat")!;
            compatibility.AddBeforeSelf(new XElement(extension + "extLst"));
            compatibility.Element(w + "noColumnBalance")?.Remove();
            compatibility.Add(new XElement(extension + "extLst", new XElement(w + "noColumnBalance")));
        });
        using WordDocument document = WordDocument.Load(source);
        Assert.Null(document.Settings.DecimalSymbol);
        Assert.False(document.CompatibilitySettings.DoNotBalanceTextColumns);
        document.Settings.DecimalSymbol = ",";
        document.CompatibilitySettings.DoNotBalanceTextColumns = true;
        using var saved = new MemoryStream(document.ToBytes());
        Assert.Empty(document.OpenXmlSdkFallbacks);

        using var reopened = new ZipArchive(saved, ZipArchiveMode.Read);
        using Stream part = reopened.GetEntry("word/settings.xml")!.Open();
        XElement written = XElement.Load(part);
        Assert.Equal(",", (string?)written.Element(w + "decimalSymbol")?.Attribute(w + "val"));
        Assert.Equal("x", (string?)written.Elements(extension + "extLst").Last().Element(w + "decimalSymbol")?.Attribute(w + "val"));
        Assert.Equal("!", (string?)written.Element(mc + "AlternateContent")?.Element(mc + "Choice")?.Element(w + "decimalSymbol")?.Attribute(w + "val"));
        Assert.Equal(extension + "extLst", written.Elements().Last().Name);
        XElement compat = written.Element(w + "compat")!;
        Assert.Single(compat.ElementsAfterSelf(w + "decimalSymbol"));
        Assert.Equal("true", (string?)compat.Element(w + "noColumnBalance")?.Attribute(w + "val"));
        Assert.Equal(extension + "extLst", compat.Elements().Last().Name);
        Assert.NotNull(compat.Element(extension + "extLst")?.Element(w + "noColumnBalance"));
    }

    [Fact]
    public void CompatibilityModeIsInsertedBeforeLaterSettings() {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var source = CreatePackageWithSettingsMarkup(settings => {
            settings.Element(w + "compat")!.Remove();
            settings.Element(w + "themeFontLang")!.AddBeforeSelf(new XElement(w + "uiCompat97To2003"));
        });
        using WordDocument document = WordDocument.Load(source);
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2010;
        using var saved = new MemoryStream(document.ToBytes());
        using var package = new ZipArchive(saved, ZipArchiveMode.Read);
        using Stream part = package.GetEntry("word/settings.xml")!.Open();
        XElement written = XElement.Load(part);
        Assert.Single(written.Element(w + "compat")!.ElementsAfterSelf(w + "uiCompat97To2003"));
    }

    private static MemoryStream CreatePackageWithSettingsMarkup(Action<XElement> edit) {
        using WordDocument document = WordDocument.Create();
        var source = new MemoryStream();
        document.Save(source);
        using (var archive = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry("word/settings.xml")!;
            XElement settings;
            using (Stream input = entry.Open()) settings = XElement.Load(input);
            entry.Delete();
            edit(settings);
            using Stream output = archive.CreateEntry("word/settings.xml").Open();
            settings.Save(output);
        }
        source.Position = 0;
        return source;
    }

    [Theory]
    [InlineData(".")]
    [InlineData(",")]
    public void Decimal_symbol_survives_save_load_and_can_be_cleared_without_invalid_values(string value) {
        using var document = WordDocument.Create();
        document.AddParagraph("One\t123.45").AddTabStop(OfficeLength.Points(36), WordTabAlignment.Decimal);
        document.Settings.DecimalSymbol = value;
        foreach (string invalid in new[] { "", "..", "\t", "\uD800" })
            Assert.Throws<ArgumentException>(() => document.Settings.DecimalSymbol = invalid);
        Assert.Equal(value, document.Settings.DecimalSymbol);
        using var output = new MemoryStream(); document.Save(output); output.Position = 0;
        using (var xml = WordprocessingDocument.Open(output, false)) Assert.Empty(new OpenXmlValidator().Validate(xml));
        output.Position = 0;
        using var reopened = WordDocument.Load(output);
        Assert.Equal(value, reopened.Settings.DecimalSymbol);
        reopened.Settings.DecimalSymbol = null;
        using var cleared = new MemoryStream(); reopened.Save(cleared); cleared.Position = 0;
        using var without = WordDocument.Load(cleared);
        Assert.Null(without.Settings.DecimalSymbol);
    }
}
