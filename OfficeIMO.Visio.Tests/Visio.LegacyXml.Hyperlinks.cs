using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioLegacyHyperlinkTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";
    private static readonly string[] Cells = {
        "Description", "Address", "SubAddress", "ExtraInfo", "Frame", "NewWindow", "Default", "Invisible", "SortKey"
    };

    [Theory]
    [InlineData("http://schemas.microsoft.com/visio/2003/core", VisioPackageType.Drawing)]
    [InlineData("urn:schemas-microsoft-com:office:visio", VisioPackageType.Template)]
    [InlineData("http://schemas.microsoft.com/visio/2003/core", VisioPackageType.Stencil)]
    public void UneditedHyperlinkRowsKeepIdentityFormulaUnitNullAndErrorState(string ns, VisioPackageType family) {
        var document = Load(ns, family);
        Assert.Equal("https://example.org/original", document.Pages[0].Shapes[0].Hyperlinks.Single().Address);
        Assert.Equal("", document.Pages[0].Connectors.Single().Hyperlinks.Single().SubAddress);
        Assert.True(document.Masters.Single().Shape.Hyperlinks.Single().NewWindow);
        AssertRoundTrips(document, xml => {
            Assert.Equal(3, xml.Descendants(Legacy + "Hyperlink").Count());
            foreach (XElement row in xml.Descendants(Legacy + "Hyperlink")) AssertOriginal(row);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TypedHyperlinkWritesReplaceSourceFormulasErrorsAndNullMarkersEvenWhenUnchanged(bool sameValue) {
        var document = Load();
        foreach (VisioHyperlink hyperlink in new[] {
                     document.Pages[0].Shapes[0].Hyperlinks.Single(),
                     document.Pages[0].Connectors.Single().Hyperlinks.Single(),
                     document.Masters.Single().Shape.Hyperlinks.Single()
                 }) AssignAll(hyperlink, sameValue);

        AssertRoundTrips(document, xml => {
            foreach (XElement row in xml.Descendants(Legacy + "Hyperlink")) {
                AssertIdentity(row);
                foreach (string name in Cells) {
                    XElement cell = row.Element(Legacy + name)!;
                    Assert.Equal(Expected(name, sameValue), cell.Value);
                    Assert.Equal(IsFlag(name) ? "BOOL" : "STR", (string?)cell.Attribute("Unit"));
                    Assert.Null(cell.Attribute("F"));
                    Assert.Null(cell.Attribute("Err"));
                    Assert.Null(cell.Attribute("V"));
                }
            }
        });

        using var archive = new ZipArchive(new MemoryStream(document.ToBytes()), ZipArchiveMode.Read);
        foreach (string part in new[] { "visio/pages/page1.xml", "visio/masters/master1.xml" }) {
            using Stream stream = archive.GetEntry(part)!.Open();
            foreach (XElement row in XDocument.Load(stream).Descendants(Modern + "Section")
                         .Where(section => (string?)section.Attribute("N") == "Hyperlink").Elements(Modern + "Row")) {
                Assert.Equal("Link", (string?)row.Attribute("N"));
                Assert.Equal("7", (string?)row.Attribute("IX"));
                foreach (XElement cell in row.Elements(Modern + "Cell")) {
                    Assert.Null(cell.Attribute("F")); Assert.Null(cell.Attribute("E")); Assert.Null(cell.Attribute("Err"));
                }
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CopiesKeepHyperlinkAssignmentIntentAndUntouchedSourceState(bool beforeCopy) {
        var document = Load(); var page = document.Pages[0];
        if (beforeCopy) AssignDescriptions(page);
        VisioPage copy = document.DuplicatePage(page, "Copy");
        if (!beforeCopy) AssignDescriptions(copy);
        document.DuplicatePage(copy, "Second");
        AssertRoundTrips(document, xml => {
            foreach (XElement exportedPage in xml.Descendants(Legacy + "Page")) {
                bool assigned = beforeCopy || (string?)exportedPage.Attribute("Name") != "Source";
                foreach (XElement row in exportedPage.Descendants(Legacy + "Hyperlink")) {
                    AssertIdentity(row);
                    XElement description = row.Element(Legacy + "Description")!;
                    Assert.Equal("", description.Value);
                    Assert.Equal(assigned ? null : "null", (string?)description.Attribute("V"));
                    Assert.Null(description.Attribute("F"));
                    Assert.Equal(assigned ? null : "producer description error", (string?)description.Attribute("Err"));
                    Assert.Equal("STR", (string?)description.Attribute("Unit"));
                    AssertOriginalCell(row, "Address");
                    AssertOriginalCell(row, "SubAddress");
                    AssertOriginalCell(row, "ExtraInfo");
                }
            }
            AssertOriginal(xml.Descendants(Legacy + "Master").Single().Descendants(Legacy + "Hyperlink").Single());
        });
        static void AssignDescriptions(VisioPage page) {
            page.Shapes[0].Hyperlinks.Single().Description = "";
            page.Connectors.Single().Hyperlinks.Single().Description = "";
        }
    }

    [Fact]
    public void ImportedMasterHyperlinkEditsKeepUnitsAndLeaveSourcePackageStateIndependent() {
        var source = Load(family: VisioPackageType.Stencil);
        var destination = VisioDocument.Create(VisioPackageType.Stencil);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
        try {
            File.WriteAllBytes(path, source.ToBytes());
            VisioMaster master = destination.ImportStencilMastersAndGet(path, new[] { "Link master" }).Single();
            master.Shape.Hyperlinks.Single().SubAddress = "Detail";
            master.Shape.Hyperlinks.Single().ExtraInfo = master.Shape.Hyperlinks.Single().ExtraInfo;
            AssertRoundTrips(destination, xml => {
                XElement row = xml.Descendants(Legacy + "Master").Single().Descendants(Legacy + "Hyperlink").Single();
                foreach (string name in new[] { "SubAddress", "ExtraInfo" }) {
                    XElement cell = row.Element(Legacy + name)!;
                    Assert.Equal(name == "SubAddress" ? "Detail" : "?q=1", cell.Value);
                    Assert.Equal("STR", (string?)cell.Attribute("Unit"));
                    Assert.Null(cell.Attribute("V")); Assert.Null(cell.Attribute("F")); Assert.Null(cell.Attribute("Err"));
                }
                AssertOriginalCell(row, "Description");
                AssertOriginalCell(row, "Address");
            });
            AssertOriginal(Export(source).Descendants(Legacy + "Master").Single().Descendants(Legacy + "Hyperlink").Single());
        } finally { File.Delete(path); }
    }

    private static VisioDocument Load(string ns = "http://schemas.microsoft.com/visio/2003/core", VisioPackageType family = VisioPackageType.Drawing) {
        string row = "<Hyperlink ID='7' NameU='Link' Name='Localized link'>"
            + "<Description V='null' Unit='STR' Err='producer description error'/><Address Unit='STR' F='&quot;https://example.org/original&quot;' Err='#VALUE!'>https://example.org/original</Address>"
            + "<SubAddress V='null' Unit='STR' Err='producer target error'/><ExtraInfo Unit='STR' Err='producer query error'>?q=1</ExtraInfo>"
            + "<Frame Unit='STR' F='Inh'>main</Frame><NewWindow Unit='BOOL' F='Inh'>1</NewWindow><Default Unit='BOOL' F='Inh'>0</Default>"
            + "<Invisible Unit='BOOL' F='Inh'>0</Invisible><SortKey Unit='STR' F='Inh'>10</SortKey></Hyperlink>";
        string shape = "<Shape ID='1' NameU='Source'><XForm><Width>2</Width><Height>1</Height></XForm>" + row + "</Shape>";
        string xml = "<VisioDocument xmlns='" + ns + "'><Masters><Master ID='0' NameU='Link master'><Shapes>" + shape
            + "</Shapes></Master></Masters><Pages><Page ID='0' Name='Source'><Shapes>" + shape
            + "<Shape ID='2' NameU='Target' Master='0'><XForm><Width>1</Width><Height>1</Height></XForm></Shape>"
            + "<Shape ID='3' NameU='Edge'><XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>2</EndX><EndY>0</EndY></XForm1D>" + row + "<Text>Edge</Text></Shape>"
            + "</Shapes><Connects><Connect FromSheet='3' FromCell='BeginX' ToSheet='1' ToCell='PinX'/><Connect FromSheet='3' FromCell='EndX' ToSheet='2' ToCell='PinX'/></Connects>"
            + "</Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml)), family).Value;
    }

    private static void AssignAll(VisioHyperlink link, bool sameValue) {
        link.Description = sameValue ? link.Description : "Edited link";
        link.Address = sameValue ? link.Address : "https://example.org/edited";
        link.SubAddress = sameValue ? link.SubAddress : "Detail";
        link.ExtraInfo = sameValue ? link.ExtraInfo : "?q=2";
        link.Frame = sameValue ? link.Frame : "detail";
        link.NewWindow = sameValue ? link.NewWindow : !link.NewWindow;
        link.Default = sameValue ? link.Default : !link.Default;
        link.Invisible = sameValue ? link.Invisible : !link.Invisible;
        link.SortKey = sameValue ? link.SortKey : "20";
    }

    private static string Expected(string cell, bool original) => cell switch {
        "Description" => original ? "" : "Edited link", "Address" => original ? "https://example.org/original" : "https://example.org/edited",
        "SubAddress" => original ? "" : "Detail", "ExtraInfo" => original ? "?q=1" : "?q=2", "Frame" => original ? "main" : "detail",
        "NewWindow" => original ? "1" : "0", "Default" or "Invisible" => original ? "0" : "1", "SortKey" => original ? "10" : "20",
        _ => throw new ArgumentException("Unknown hyperlink cell", nameof(cell))
    };
    private static bool IsFlag(string name) => name is "NewWindow" or "Default" or "Invisible";
    private static void AssertOriginal(XElement row) {
        AssertIdentity(row);
        foreach (string cell in Cells) AssertOriginalCell(row, cell);
    }
    private static void AssertIdentity(XElement row) {
        Assert.Equal("7", (string?)row.Attribute("ID")); Assert.Equal("Link", (string?)row.Attribute("NameU"));
        Assert.Equal("Localized link", (string?)row.Attribute("Name"));
    }
    private static void AssertOriginalCell(XElement row, string name) {
        XElement cell = row.Element(Legacy + name)!;
        Assert.Equal(Expected(name, true), cell.Value);
        Assert.Equal(name is "Description" or "SubAddress" ? "null" : null, (string?)cell.Attribute("V"));
        Assert.Equal(name switch {
            "Address" => "\"https://example.org/original\"", "Description" or "SubAddress" or "ExtraInfo" => null, _ => "Inh"
        }, (string?)cell.Attribute("F"));
        Assert.Equal(IsFlag(name) ? "BOOL" : "STR", (string?)cell.Attribute("Unit"));
        Assert.Equal(name switch {
            "Description" => "producer description error", "Address" => "#VALUE!", "SubAddress" => "producer target error", "ExtraInfo" => "producer query error", _ => null
        }, (string?)cell.Attribute("Err"));
    }
    private static void AssertRoundTrips(VisioDocument document, Action<XDocument> assert) {
        assert(Export(document));
        VisioDocument native = VisioDocument.Load(new MemoryStream(document.ToBytes())); assert(Export(native));
        VisioDocument legacy = VisioDocument.LoadLegacyXml(new MemoryStream(native.ToLegacyXmlResult().Value), document.PackageType).Value;
        assert(Export(legacy));
    }
    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
}
