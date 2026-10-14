using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioNativeFontWireTests {
    private static readonly XNamespace Native = "http://schemas.microsoft.com/office/visio/2012/main";
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private const string Transform = "<XForm><PinX>2</PinX><PinY>2</PinY><Width>2</Width><Height>1</Height><LocPinX>1</LocPinX><LocPinY>0.5</LocPinY></XForm>";

    [Fact]
    public void NamedNativeFontsLoadIntoEditableShapeConnectorAndMasterStyles() {
        var seed = VisioDocument.Create();
        var master = seed.RegisterMaster("Native font master", new VisioShape("1", 1, .5, 2, 1, "native master"));
        var page = seed.AddPage("Fonts");
        var shape = page.AddRectangle(2, 2, 2, 1, "native shape");
        var target = page.AddRectangle(6, 2, 2, 1, "target");
        page.AddConnector(shape, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).Label = "native connector";
        page.AddShape("instance", master, 2, 5, 2, 1);

        byte[] input = RewritePackage(seed.ToBytes(), archive => {
            RewritePart(archive, "visio/document.xml", xml => {
                xml.Root!.Element(Native + "FaceNames")?.Remove();
                xml.Root.Add(new XElement(Native + "FaceNames",
                    new XElement(Native + "FaceName", new XAttribute("NameU", "Arial")),
                    new XElement(Native + "FaceName", new XAttribute("NameU", "Consolas")),
                    new XElement(Native + "FaceName", new XAttribute("NameU", "Calibri"))));
            });
            RewritePart(archive, "visio/pages/page1.xml", xml => {
                SetNamedFont(FindTextShape(xml, "native shape"), "Arial");
                SetNamedFont(FindTextShape(xml, "native connector"), "Consolas");
            });
            RewritePart(archive, "visio/masters/master1.xml", xml => SetNamedFont(FindTextShape(xml, "native master"), "Calibri"));
        });

        VisioDocument loaded = LoadNative(input);
        Assert.Equal("Arial", loaded.Pages[0].Shapes.Single(candidate => candidate.Text == "native shape").TextStyle!.FontFamily);
        Assert.Equal("Consolas", Assert.Single(loaded.Pages[0].Connectors).TextStyle!.FontFamily);
        Assert.Equal("Calibri", loaded.GetMaster("Native font master").Shape.TextStyle!.FontFamily);

        loaded.Pages[0].Shapes.Single(candidate => candidate.Text == "native shape").TextStyle!.FontFamily = "Aptos";
        byte[] saved = loaded.ToBytes();
        Assert.Equal("Aptos", Cell(CharacterRow(FindTextShape(ReadPart(saved, "visio/pages/page1.xml"), "native shape")), "Font"));
        Assert.Equal("Consolas", Cell(CharacterRow(FindTextShape(ReadPart(saved, "visio/pages/page1.xml"), "native connector")), "Font"));
        Assert.Equal("Calibri", Cell(CharacterRow(FindTextShape(ReadPart(saved, "visio/masters/master1.xml"), "native master")), "Font"));
        AssertNativeFaceNames(ReadPart(saved, "visio/document.xml").Root!.Element(Native + "FaceNames")!);

        VisioDocument reopened = LoadNative(saved);
        Assert.Equal("Aptos", reopened.Pages[0].Shapes.Single(candidate => candidate.Text == "native shape").TextStyle!.FontFamily);
        Assert.Equal("Consolas", Assert.Single(reopened.Pages[0].Connectors).TextStyle!.FontFamily);
        Assert.Equal("Calibri", reopened.GetMaster("Native font master").Shape.TextStyle!.FontFamily);
    }

    [Fact]
    public void NativePackagesWithoutFontEntriesOmitTheEmptyFaceNamesTable() {
        var document = VisioDocument.Create();
        document.AddPage("No fonts").AddRectangle(2, 2, 2, 1, "plain");
        byte[] saved = document.ToBytes();
        Assert.Null(ReadPart(saved, "visio/document.xml").Root!.Element(Native + "FaceNames"));
        Assert.Null(ReadPart(LoadNative(saved).ToBytes(), "visio/document.xml").Root!.Element(Native + "FaceNames"));
    }

    [Fact]
    public void LiteralFontEditsInSourceRowsDeclareTheNativeFamilyAndExportALegacyIdentity() {
        var document = LoadLegacy("<VisioDocument xmlns='" + Legacy + "'><FaceNames><FaceName ID='5' Name='Arial'/></FaceNames>"
            + "<Pages><Page ID='0'><Shapes>" + Shape("1", "changed", "<Char IX='0'><Font>5</Font></Char><Char IX='1'><Font>5</Font></Char>")
            + "</Shapes></Page></Pages></VisioDocument>");
        VisioShape shape = document.Pages[0].Shapes.Single();
        VisioShapeSheetSection section = shape.GetShapeSheetSections().Single(s => s.Name == "Character");
        section.Rows.Single(row => row.Index == 1).SetCell("Font", "Consolas");
        shape.SetShapeSheetSection(section);
        byte[] package = document.ToBytes();
        Assert.Contains(ReadPart(package, "visio/document.xml").Root!.Element(Native + "FaceNames")!.Elements(Native + "FaceName"),
            face => (string?)face.Attribute("NameU") == "Consolas");
        foreach (VisioDocument candidate in new[] { document, LoadNative(package) }) {
            XDocument xml = XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value));
            string id = xml.Descendants(Legacy + "Char").Single(row => (string?)row.Attribute("IX") == "1").Element(Legacy + "Font")!.Value;
            Assert.Equal("Consolas", (string?)xml.Descendants(Legacy + "FaceName").Single(face => (string?)face.Attribute("ID") == id).Attribute("Name"));
        }
        Assert.Equal("Consolas", shape.GetShapeSheetSections().Single(s => s.Name == "Character").Rows.Single(row => row.Index == 1).Cells.Single(cell => cell.Name == "Font").Value);
    }

    [Fact]
    public void NamedAuxiliaryFontsDoNotBecomeZeroFallbackAfterAnExternalNativeEdit() {
        var document = LoadLegacy("<VisioDocument xmlns='" + Legacy + "'><FaceNames><FaceName ID='0' Name='Arial'/></FaceNames>"
            + "<Pages><Page ID='0'><Shapes>" + Shape("1", "edited", "<Char IX='0'><Font>0</Font></Char>") + "</Shapes></Page></Pages></VisioDocument>");
        byte[] changed = RewritePackage(document.ToBytes(), archive => RewritePart(archive, "visio/pages/page1.xml", xml => {
            XElement shape = FindTextShape(xml, "edited"), row = CharacterRow(shape);
            foreach (string name in new[] { "AsianFont", "ComplexScriptFont" }) row.Add(new XElement(Native + "Cell", new XAttribute("N", name), new XAttribute("V", "Arial")));
            shape.Add(new XElement(Native + "Section", new XAttribute("N", "Paragraph"), new XElement(Native + "Row", new XAttribute("IX", "0"),
                new XElement(Native + "Cell", new XAttribute("N", "BulletFont"), new XAttribute("V", "Arial")))));
        }));
        VisioDocument reopened = LoadNative(changed);
        XDocument legacy = XDocument.Load(new MemoryStream(reopened.ToLegacyXmlResult().Value));
        foreach (string name in new[] { "AsianFont", "ComplexScriptFont", "BulletFont" }) {
            string id = legacy.Descendants(Legacy + name).Single().Value;
            Assert.NotEqual("0", id);
            Assert.Equal("Arial", (string?)legacy.Descendants(Legacy + "FaceName").Single(face => (string?)face.Attribute("ID") == id).Attribute("Name"));
        }
        AssertFontReferences(FindTextShape(ReadPart(reopened.ToBytes(), "visio/pages/page1.xml"), "edited"), "Arial", "Arial", "Arial", "Arial");
    }

    [Fact]
    public void LegacyFontReferencesProjectAcrossNativePartsWithoutResolvingFallbackSentinels() {
        const string faces = "<FaceNames><FaceName ID='0' Name='Arial'/><FaceName ID='5' Name='Consolas'/>"
            + "<FaceName ID='6' Name='SimSun'/><FaceName ID='7' Name='Mangal'/><FaceName ID='8' Name='Wingdings'/></FaceNames>";
        const string positive = "<Char IX='0'><Font F='GUARD(5)'>5</Font><AsianFont>6</AsianFont><ComplexScriptFont>7</ComplexScriptFont></Char>"
            + "<Para IX='0'><Bullet>1</Bullet><BulletStr>o.</BulletStr><BulletFont>8</BulletFont></Para>";
        const string zero = "<Char IX='0'><Font>0</Font><AsianFont>0</AsianFont><ComplexScriptFont>0</ComplexScriptFont></Char>"
            + "<Para IX='0'><Bullet>1</Bullet><BulletStr>o.</BulletStr><BulletFont>0</BulletFont></Para>";
        const string empty = "<Char IX='0'><Font>0</Font><AsianFont/><ComplexScriptFont/></Char>"
            + "<Para IX='0'><Bullet>1</Bullet><BulletStr>o.</BulletStr><BulletFont/></Para>";
        string source = "<VisioDocument xmlns='" + Legacy + "'>" + faces
            + "<StyleSheets><StyleSheet ID='10' NameU='Font style'>" + positive + "</StyleSheet></StyleSheets>"
            + "<Masters><Master ID='8' NameU='Font master'><Shapes>" + Shape("1", "master", positive) + "</Shapes></Master></Masters>"
            + "<Pages><Page ID='0'><Shapes>" + Shape("1", "positive", positive, "Master='8' TextStyle='10'")
            + Shape("2", "zero", zero) + Shape("3", "empty", empty) + "</Shapes></Page></Pages></VisioDocument>";
        byte[] package = LoadLegacy(source).ToBytes();
        XDocument documentXml = ReadPart(package, "visio/document.xml");
        AssertNativeFaceNames(documentXml.Root!.Element(Native + "FaceNames")!);
        XElement style = documentXml.Descendants(Native + "StyleSheet").Single(element => (string?)element.Attribute("ID") == "10");
        XElement nativeMaster = FindTextShape(ReadPart(package, "visio/masters/master1.xml"), "master");
        XDocument pageXml = ReadPart(package, "visio/pages/page1.xml");
        foreach (XElement sheet in new[] { style, nativeMaster, FindTextShape(pageXml, "positive") })
        {
            AssertFontReferences(sheet, "Consolas", "SimSun", "Mangal", "Wingdings");
            Assert.Equal("GUARD(FONT(\"Consolas\"))", (string?)CharacterRow(sheet).Elements(Native + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Font").Attribute("F"));
        }
        AssertFontReferences(FindTextShape(pageXml, "zero"), "Arial", "0", "0", "0");
        AssertFontReferences(FindTextShape(pageXml, "empty"), "Arial", "", "", "");
    }

    [Fact]
    public void NativeReopeningRestoresDistinctLegacyFontIdentitiesTablesAndGuardedCells() {
        const string tables = "<Fonts><FontEntry ID='3' Name='Consolas' CharSet='0' PitchAndFamily='49' Attributes='123' Weight='400' Unicode='1'/>"
            + "<FontEntry ID='7' Name='Consolas' CharSet='128' PitchAndFamily='49' Attributes='125' Weight='400' Unicode='1'/></Fonts>"
            + "<FaceNames><FaceName ID='3' Name='Consolas' UnicodeRanges='0-255' CharSets='0' Panos='2 11 6 9 3 5 4 4 2 4' Flags='325'/>"
            + "<FaceName ID='7' Name='Consolas' UnicodeRanges='0-65535' CharSets='128' Panos='2 11 6 9 3 5 4 4 2 4' Flags='421'/></FaceNames>";
        const string rows = "<Char IX='0'><Font F='GUARD(3)' Err='first font error'>3</Font><Size Unit='PT'>0.1666666666666667</Size></Char>"
            + "<Char IX='4'><Font F='7' Err='second font error'>7</Font><AsianFont F='Inh'>7</AsianFont><ComplexScriptFont F='GUARD(3)'>3</ComplexScriptFont></Char>"
            + "<Para IX='0'><Bullet>1</Bullet><BulletStr>o.</BulletStr><BulletFont F='7' Err='bullet font error'>7</BulletFont></Para>";
        string source = "<VisioDocument xmlns='" + Legacy + "'>" + tables
            + "<Pages><Page ID='0'><Shapes>" + Shape("1", "duplicate family", rows) + "</Shapes></Page></Pages></VisioDocument>";
        XDocument original = XDocument.Parse(source);
        VisioDocument document = LoadLegacy(source);

        // Both font IDs have the same native family value; only guarded source identity
        // can restore the distinct legacy charset entries and formulas.
        for (int round = 0; round < 2; round++) {
            byte[] package = document.ToBytes();
            AssertNativeFaceNames(ReadPart(package, "visio/document.xml").Root!.Element(Native + "FaceNames")!);
            XElement nativeShape = FindTextShape(ReadPart(package, "visio/pages/page1.xml"), "duplicate family");
            Assert.Equal(new[] { "Consolas", "Consolas" }, nativeShape.Elements(Native + "Section")
                .Single(section => (string?)section.Attribute("N") == "Character").Descendants(Native + "Cell")
                .Where(cell => (string?)cell.Attribute("N") == "Font").Select(cell => (string?)cell.Attribute("V")).ToArray());
            document = LoadNative(package);
            XDocument restored = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
            foreach (string name in new[] { "Fonts", "FaceNames" })
                Assert.True(XNode.DeepEquals(LogicalXml(original.Root!.Element(Legacy + name)!), LogicalXml(restored.Root!.Element(Legacy + name)!)), name);
            Assert.Equal(FontCells(original), FontCells(restored));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExternalNativeFontTableEditsInvalidateObsoleteLegacyCellIdentities(bool removeOriginalDeclaration) {
        string rows = "<Char IX='0'><Font F='GUARD(5)'>5</Font></Char>";
        var source = LoadLegacy("<VisioDocument xmlns='" + Legacy + "'><FaceNames><FaceName ID='5' Name='Arial'/></FaceNames>"
            + "<StyleSheets><StyleSheet ID='9'>" + rows + "</StyleSheet></StyleSheets>"
            + "<Masters><Master ID='1' NameU='Fonts'><Shapes>" + Shape("1", "master", rows) + "</Shapes></Master></Masters>"
            + "<Pages><Page ID='0'><Shapes>" + Shape("2", "page", rows) + "</Shapes></Page></Pages></VisioDocument>", VisioPackageType.Stencil);
        byte[] changed = RewritePackage(source.ToBytes(), archive => RewritePart(archive, "visio/document.xml", xml => {
            XElement faces = xml.Root!.Element(Native + "FaceNames")!;
            if (removeOriginalDeclaration) faces.Elements().Remove();
            faces.Add(new XElement(Native + "FaceName", new XAttribute("NameU", "Consolas")));
        }));
        var loaded = LoadNative(changed);
        if (!removeOriginalDeclaration) {
            Assert.Equal("Arial", Assert.IsType<VisioTextStyle>(loaded.Pages[0].Shapes.Single().TextStyle).FontFamily);
            Assert.Equal("Arial", loaded.GetMaster("Fonts").Shape.TextStyle!.FontFamily);
        }
        for (int round = 0; round < 2; round++) {
            byte[] saved = loaded.ToBytes();
            foreach (string part in new[] { "visio/document.xml", "visio/pages/page1.xml", "visio/masters/master1.xml" })
                Assert.All(ReadPart(saved, part).Descendants(Native + "Cell").Where(cell => (string?)cell.Attribute("N") == "Font"),
                    cell => Assert.Equal("Arial", (string?)cell.Attribute("V")));
            loaded = LoadNative(saved);
            Assert.Equal("Arial", Assert.IsType<VisioTextStyle>(loaded.Pages[0].Shapes.Single().TextStyle).FontFamily);
        }
    }

    [Fact]
    public void StencilImportMapsAuxiliaryFontAliasesCreatedAfterDocumentDecoding() {
        string rows = "<Char IX='0'><Font>0</Font><AsianFont>0</AsianFont><ComplexScriptFont>0</ComplexScriptFont></Char>"
            + "<Para IX='0'><Bullet>1</Bullet><BulletFont>0</BulletFont></Para>";
        var stencil = LoadLegacy("<VisioDocument xmlns='" + Legacy + "'><FaceNames><FaceName ID='0' Name='Arial'/></FaceNames>"
            + "<Masters><Master ID='1' NameU='Fonts'><Shapes>" + Shape("1", "master", rows) + "</Shapes></Master></Masters></VisioDocument>", VisioPackageType.Stencil);
        byte[] changed = RewritePackage(stencil.ToBytes(), archive => RewritePart(archive, "visio/masters/master1.xml", xml => {
            foreach (XElement cell in xml.Descendants(Native + "Cell").Where(cell => (string?)cell.Attribute("N") is "AsianFont" or "ComplexScriptFont" or "BulletFont"))
                cell.SetAttributeValue("V", "Arial");
        }));
        var destination = LoadLegacy("<VisioDocument xmlns='" + Legacy + "'><FaceNames><FaceName ID='0' Name='Calibri'/>"
            + "<FaceName ID='1' Name='Consolas'/></FaceNames><Pages><Page ID='0'><Shapes/></Page></Pages></VisioDocument>");
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
        try {
            File.WriteAllBytes(path, changed);
            destination.ImportStencilMasters(path);
            destination.Pages[0].AddShape("instance", destination.GetMaster("Fonts"), 2, 2, 2, 1);
            for (int round = 0; round < 2; round++) {
                byte[] saved = destination.ToBytes();
                XElement master = FindTextShape(ReadPart(saved, "visio/masters/master1.xml"), "master");
                AssertFontReferences(master, "Arial", "Arial", "Arial", "Arial");
                destination = LoadNative(saved);
            }
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 4)]
    [InlineData(true, 4)]
    public void SameFamilyAssignmentWithoutFormulaClearsFontErrorsForShapesAndConnectors(bool connector, int rowIndex) {
        string rows = "<Char IX='" + rowIndex + "'><Font Err='font error'>5</Font><Size Err='size error'>0.1666666666666667</Size></Char>";
        string geometry = connector ? "<XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>3</EndX><EndY>1</EndY></XForm1D>" : "";
        var document = LoadLegacy("<VisioDocument xmlns='" + Legacy + "'><FaceNames><FaceName ID='5' Name='Arial'/></FaceNames>"
            + "<Pages><Page ID='0'><Shapes>" + Shape("1", "errors", rows + geometry, connector ? "OneD='1'" : "")
            + "</Shapes></Page></Pages></VisioDocument>");
        XDocument untouched = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        Assert.Equal("font error", (string?)untouched.Descendants(Legacy + "Font").Single().Attribute("Err"));
        VisioTextStyle style = connector ? document.Pages[0].Connectors.Single().TextStyle! : document.Pages[0].Shapes.Single().TextStyle!;
        style.FontFamily = "Arial";
        for (int round = 0; round < 2; round++) {
            XDocument legacy = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
            Assert.Null(legacy.Descendants(Legacy + "Font").Single().Attribute("Err"));
            Assert.Equal("size error", (string?)legacy.Descendants(Legacy + "Size").Single().Attribute("Err"));
            document = LoadNative(document.ToBytes());
        }
    }

    private static void AssertNativeFaceNames(XElement table) {
        Assert.NotEmpty(table.Elements(Native + "FaceName"));
        Assert.All(table.Elements(Native + "FaceName"), face => {
            Assert.False(string.IsNullOrWhiteSpace((string?)face.Attribute("NameU")));
            Assert.Null(face.Attribute("ID"));
            Assert.Null(face.Attribute("Name"));
        });
    }

    private static void AssertFontReferences(XElement sheet, string font, string asian, string complex, string bullet) {
        XElement character = CharacterRow(sheet);
        Assert.Equal(font, Cell(character, "Font"));
        Assert.Equal(asian, Cell(character, "AsianFont"));
        Assert.Equal(complex, Cell(character, "ComplexScriptFont"));
        XElement paragraph = sheet.Elements(Native + "Section").Single(section => (string?)section.Attribute("N") == "Paragraph").Elements(Native + "Row").Single();
        Assert.Equal(bullet, Cell(paragraph, "BulletFont"));
    }

    private static XElement CharacterRow(XElement shape) => shape.Elements(Native + "Section")
        .Single(section => (string?)section.Attribute("N") == "Character").Elements(Native + "Row").Single();

    private static string? Cell(XElement row, string name) => (string?)row.Elements(Native + "Cell")
        .Single(cell => (string?)cell.Attribute("N") == name).Attribute("V");

    private static XElement FindTextShape(XDocument xml, string text) => xml.Descendants(Native + "Shape")
        .Single(shape => shape.Element(Native + "Text")?.Value == text);

    private static void SetNamedFont(XElement shape, string family) {
        shape.Elements(Native + "Section").Where(section => (string?)section.Attribute("N") is "Character" or "Char").Remove();
        shape.Add(new XElement(Native + "Section", new XAttribute("N", "Character"),
            new XElement(Native + "Row", new XAttribute("IX", "0"),
                new XElement(Native + "Cell", new XAttribute("N", "Font"), new XAttribute("V", family)),
                new XElement(Native + "Cell", new XAttribute("N", "Size"), new XAttribute("V", "0.1666666666666667")))));
    }

    private static string Shape(string id, string text, string rows, string attributes = "") =>
        "<Shape ID='" + id + "' " + attributes + ">" + Transform + rows + "<Text>" + text + "</Text></Shape>";

    private static string[] FontCells(XDocument xml) => xml.Descendants().Where(element => element.Name.Namespace == Legacy &&
        element.Name.LocalName is "Font" or "AsianFont" or "ComplexScriptFont" or "BulletFont")
        .Select(element => LogicalXml(element).ToString(SaveOptions.DisableFormatting)).ToArray();

    private static XElement LogicalXml(XElement element) {
        var copy = new XElement(element);
        copy.DescendantsAndSelf().Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
        copy.DescendantNodes().OfType<XText>().Where(text => string.IsNullOrWhiteSpace(text.Value)).Remove();
        return copy;
    }

    private static VisioDocument LoadLegacy(string source, VisioPackageType type = VisioPackageType.Drawing) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source)), type).Value;
    private static VisioDocument LoadNative(byte[] bytes) => VisioDocument.Load(new MemoryStream(bytes, writable: false));

    private static XDocument ReadPart(byte[] package, string path) {
        using var stream = new MemoryStream(package, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        using Stream input = archive.GetEntry(path)!.Open();
        return XDocument.Load(input);
    }

    private static byte[] RewritePackage(byte[] package, Action<ZipArchive> rewrite) {
        using var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) rewrite(archive);
        return stream.ToArray();
    }

    private static void RewritePart(ZipArchive archive, string path, Action<XDocument> rewrite) {
        ZipArchiveEntry entry = archive.GetEntry(path)!;
        XDocument xml;
        using (Stream input = entry.Open()) xml = XDocument.Load(input);
        rewrite(xml);
        entry.Delete();
        using Stream output = archive.CreateEntry(path).Open();
        xml.Save(output);
    }
}
