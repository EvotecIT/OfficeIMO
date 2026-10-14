using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioMasterInstancesTests {
    [Fact]
    public void GeneratedMasterChildLinksSurviveLaterInsertionAndReordering() {
        var document = VisioDocument.Create();
        var root = new VisioShape("root", 1, 1, 2, 2, "");
        root.Children.Add(new VisioShape("child", 1, 1, 1, 1, "original"));
        var master = document.RegisterMaster("Strings", root);
        var instance = document.AddPage("Page").AddShape("instance", master, 3, 3, 2, 2);
        string assigned = instance.Children[0].MasterShapeId!;
        root.Children.Insert(0, new VisioShape(assigned, 1, 1, 1, 1, "inserted"));
        foreach (var loaded in new[] { VisioDocument.Load(new MemoryStream(document.ToBytes())), VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value }) {
            var child = loaded.Pages[0].Shapes[0].Children[0];
            Assert.Equal(assigned, child.MasterShapeId);
            Assert.Equal("original", child.MasterShape!.Text);
            Assert.Equal("child", child.MasterShape.Id);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MasterFormattingDoesNotDuplicatePreservedLocalTextSections(bool multipleRows) {
        string extraRows = multipleRows ? "<Char IX='1'><Size>0.2</Size></Char><Para IX='1'><HorzAlign>2</HorzAlign></Para>" : "";
        string xml = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Masters><Master ID='1' NameU='Text'><Shapes><Shape ID='1'><XForm><Width>2</Width><Height>1</Height></XForm><Char IX='0'><Size>0.1666666666666667</Size></Char><Para IX='0'><HorzAlign>1</HorzAlign></Para></Shape></Shapes></Master></Masters><Pages><Page ID='1'><Shapes><Shape ID='1' Master='1'><Char IX='0'><Size>0.1</Size><FontScale>1</FontScale></Char><Para IX='0'><HorzAlign>0</HorzAlign><IndLeft>0.1</IndLeft></Para>" + extraRows + "<Text><cp IX='0'/><pp IX='0'/>Local text</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
        foreach (byte[] bytes in new[] { document.ToLegacyXmlResult().Value, VisioDocument.Load(new MemoryStream(document.ToBytes())).ToLegacyXmlResult().Value }) {
            var saved = XDocument.Load(new MemoryStream(bytes)); XNamespace ns = saved.Root!.Name.Namespace;
            var shape = saved.Descendants(ns + "Page").Descendants(ns + "Shape").Single();
            Assert.Equal(multipleRows ? 2 : 1, shape.Elements(ns + "Char").Count());
            Assert.Equal(multipleRows ? 2 : 1, shape.Elements(ns + "Para").Count());
            var character = shape.Elements(ns + "Char").Single(row => (string?)row.Attribute("IX") == "0");
            Assert.Equal(0.1, (double)character.Element(ns + "Size")!);
            Assert.Equal(1, (int)character.Element(ns + "FontScale")!);
            var paragraph = shape.Elements(ns + "Para").Single(row => (string?)row.Attribute("IX") == "0");
            Assert.Equal(0, (int)paragraph.Element(ns + "HorzAlign")!);
            Assert.Equal(0.1, (double)paragraph.Element(ns + "IndLeft")!);
        }
    }

    [Fact]
    public void ImportedSymbolTextIsAvailableImmediatelyAndCanBeCleared() {
        var document = VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "nxbre-ie3.vsx")).Value;
        var page = document.AddPage("Symbols");
        var master = document.GetMaster("Equivalent");
        var symbol = page.AddShape("symbol", master, 3, 3, 2, 2);
        Assert.Equal(master.Shape.Text, symbol.Text);
        Assert.Contains("=", page.ToSvg());
        foreach (string? empty in new string?[] { "", null }) {
            symbol.Text = empty;
            foreach (var saved in new[] { VisioDocument.Load(new MemoryStream(document.ToBytes())), VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value }) {
                Assert.Equal("", saved.Pages.Single(p => p.Name == "Symbols").Shapes[0].Text);
                Assert.Equal(master.Shape.Text, saved.GetMaster("Equivalent").Shape.Text);
            }
        }
    }

    [Fact]
    public void ExplicitZeroLocalPinsSurviveStandaloneAndMasterBackedSaves() {
        var document = VisioDocument.Create();
        var page = document.AddPage("Pins");
        var master = document.RegisterMaster("Box", new VisioShape("1", 1, 1, 2, 2, ""));
        page.AddRectangle(1, 1, 2, 2);
        page.AddShape("master", master, 5, 5, 2, 2);
        foreach (var shape in page.Shapes) { shape.LocPinX = 0; shape.LocPinY = 0; }
        foreach (bool deltasOnly in new[] { true, false }) {
            document.WriteMasterDeltasOnly = deltasOnly;
            foreach (var saved in new[] { VisioDocument.Load(new MemoryStream(document.ToBytes())), VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value }) {
                Assert.All(saved.Pages[0].Shapes, shape => { Assert.Equal(0, shape.LocPinX); Assert.Equal(0, shape.LocPinY); });
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportIntoDocumentWithExistingFontsKeepsBothFontFamilies(bool guardedFont) {
        var stencil = VisioDocument.Create(VisioPackageType.Stencil);
        var root = new VisioShape("1", 1, 1, 2, 2, "");
        root.Children.Add(new VisioShape("2", 1, 1, 1, 1, "Imported") { TextStyle = new VisioTextStyle { FontFamily = "Consolas", Size = 12 } });
        stencil.RegisterMaster("Group", root);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
        try {
            File.WriteAllBytes(path, stencil.ToBytes());
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Update)) {
                var entry = archive.GetEntry("visio/masters/master1.xml")!;
                XDocument xml;
                using (var stream = entry.Open()) xml = XDocument.Load(stream);
                var font = xml.Descendants().Single(cell => (string?)cell.Attribute("N") == "Font");
                string family = (string)font.Attribute("V")!;
                string formula = "FONT(\"" + family + "\")";
                font.SetAttributeValue("F", guardedFont ? "GUARD(" + formula + ")" : formula);
                entry.Delete();
                using var output = archive.CreateEntry("visio/masters/master1.xml").Open();
                xml.Save(output);
            }
            var document = VisioDocument.Create(); var page = document.AddPage("Page");
            page.AddRectangle(1, 1, 1, 1, "Existing").TextStyle = new VisioTextStyle { FontFamily = "Arial", Size = 12 };
            _ = document.ToBytes();
            document.ImportStencilMasters(path);
            var importedFont = document.GetMaster("Group").RawMasterContentXml!.Descendants().Single(cell => (string?)cell.Attribute("N") == "Font");
            string importedId = (string)importedFont.Attribute("V")!;
            Assert.Equal("1", importedId);
            Assert.Equal(guardedFont ? "GUARD(FONT(\"Consolas\"))" : "FONT(\"Consolas\")", (string?)importedFont.Attribute("F"));
            var instance = page.AddShape("group", "Group", 4, 4, 2, 2);
            Assert.Equal("Consolas", instance.Children[0].TextStyle!.FontFamily);
            var loaded = VisioDocument.Load(new MemoryStream(document.ToBytes()));
            Assert.Equal("Arial", loaded.Pages[0].Shapes[0].TextStyle!.FontFamily);
            Assert.Equal("Consolas", loaded.Pages[0].FindShapeById("group")!.Children[0].TextStyle!.FontFamily);
            Assert.Equal("Consolas", loaded.GetMaster("Group").Shape.Children[0].TextStyle!.FontFamily);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void MinimalNativeInstanceInheritsChildPlacementTextAndStyleBeforeEditing() {
        const string xml = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Fonts><FontEntry ID='0' Name='Arial'/></Fonts><Masters><Master ID='1' NameU='Group'><Shapes><Shape ID='7' Type='Group'><XForm><Width>4</Width><Height>2</Height><LocPinX>2</LocPinX><LocPinY>1</LocPinY></XForm><Shapes><Shape ID='8'><XForm><PinX>1</PinX><PinY>0.5</PinY><Width>2</Width><Height>1</Height><LocPinX>1</LocPinX><LocPinY>0.5</LocPinY></XForm><Fill><FillForegnd>#FF0000</FillForegnd></Fill><Char IX='0'><Font>0</Font><Size>0.1666666666666667</Size></Char><Text>Inherited child</Text></Shape></Shapes></Shape></Shapes></Master></Masters><Pages><Page ID='0'><Shapes><Shape ID='10' Master='1' Type='Group'><XForm><PinX>5</PinX><PinY>5</PinY><Width>8</Width><Height>4</Height><LocPinX>4</LocPinX><LocPinY>2</LocPinY></XForm><Shapes><Shape ID='11' MasterShape='8'/></Shapes></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
        var child = document.Pages[0].Shapes[0].Children[0];
        Assert.Equal(2, child.PinX); Assert.Equal(1, child.PinY); Assert.Equal(4, child.Width); Assert.Equal(2, child.Height);
        Assert.Equal("Inherited child", child.Text);
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.Red, child.FillColor);
        Assert.Equal("Arial", child.TextStyle!.FontFamily); Assert.Equal(12, child.TextStyle.Size!.Value, 6);
        child.PinX = 3; child.Text = "edited";
        foreach (var saved in new[] { VisioDocument.Load(new MemoryStream(document.ToBytes())), VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value }) {
            var reopened = saved.Pages[0].Shapes[0].Children[0];
            Assert.Equal(3, reopened.PinX); Assert.Equal(4, reopened.Width); Assert.Equal("edited", reopened.Text);
            Assert.Equal(OfficeIMO.Drawing.OfficeColor.Red, reopened.FillColor);
            Assert.Equal("Inherited child", saved.Masters.First().Shape.Children[0].Text);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IndependentGroupedStencilCreatesEditableInstancesThatReopen(bool importPackage) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "nxbre-ie3.vsx");
        var document = VisioDocument.LoadLegacyXml(path).Value;
        if (importPackage) {
            string package = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            try {
                File.WriteAllBytes(package, document.ToBytes());
                document = VisioDocument.Create();
                document.ImportStencilMasters(package, new[] { "Atom" });
            } finally { File.Delete(package); }
        }
        var page = document.AddPage("Instances", 12, 10);
        var master = document.GetMaster("Atom");
        var first = page.AddShape("first", master, 3, 6, master.Shape.Width * 2, master.Shape.Height * 2);
        var second = page.AddShape("second", "Atom", 8, 6, master.Shape.Width, master.Shape.Height);
        Assert.Equal(2, first.Children.Count);
        Assert.Equal("Group", first.Type);
        Assert.Equal(master.Shape.Children[0].Width * 2, first.Children[0].Width, 8);
        Assert.Equal(master.Shape.Children[0].Text, first.Children[0].Text);
        Assert.Equal(master.Shape.Children[0].FillColor, first.Children[0].FillColor);
        first.Children[0].Text = "Edited relation";
        first.Children[0].SetUserCell("Review", "approved");
        first.PinX += 1;
        var connector = page.AddConnector(first, second);
        Assert.NotEqual(first.Children[0].Text, second.Children[0].Text);
        Assert.NotEqual(first.Children[0].Text, master.Shape.Children[0].Text);
        Assert.Contains("Edited relation", page.ToSvg());
        foreach (var loaded in new[] {
            VisioDocument.Load(new MemoryStream(document.ToBytes())),
            VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value
        }) {
            var savedPage = loaded.Pages.Single(p => p.Name == "Instances");
            var instance = savedPage.FindShapeById("first")!;
            Assert.Equal(2, instance.Children.Count);
            Assert.Equal("Edited relation", instance.Children[0].Text);
            Assert.Equal("approved", instance.Children[0].GetUserCellValue("Review"));
            Assert.Equal(first.Children[0].Width, instance.Children[0].Width, 8);
            Assert.Equal(first.Children[0].PinX, instance.Children[0].PinX, 8);
            for (int index = 0; index < first.Children.Count; index++) {
                Assert.Equal(first.Children[index].PinY, instance.Children[index].PinY, 8);
                Assert.Equal(first.Children[index].LocPinX, instance.Children[index].LocPinX, 8);
                Assert.Equal(first.Children[index].LocPinY, instance.Children[index].LocPinY, 8);
            }
            Assert.NotNull(instance.Children[0].MasterShape);
            Assert.Equal("first", savedPage.Connectors.Single().From?.Id);
            Assert.Equal("second", savedPage.Connectors.Single().To?.Id);
            Assert.Contains("Edited relation", savedPage.ToSvg());
        }
    }

    [Fact]
    public void ImportedGeometryCachesFollowRequestedSizeIncludingCircularArcs() {
        var document = VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "nxbre-ie3.vsx")).Value;
        var page = document.AddPage("Scaled");
        var master = document.GetMaster("Atom");
        var instance = page.AddShape("scaled", master, 4, 4, master.Shape.Width * 3, master.Shape.Height * 2);
        var xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        XNamespace ns = xml.Root!.Name.Namespace;
        var group = xml.Descendants(ns + "Page").Single(p => (string?)p.Attribute("Name") == "Scaled").Element(ns + "Shapes")!.Elements(ns + "Shape").Single();
        var line = group.Element(ns + "Shapes")!.Elements(ns + "Shape").First().Element(ns + "Geom")!.Elements(ns + "LineTo").First();
        Assert.Equal(instance.Children[0].Width, (double)line.Element(ns + "X")!, 8);
        var warning = document.GetMaster("Warning sign");
        page.AddShape("warning", warning, 5, 5, warning.Shape.Width * 2, warning.Shape.Height * 3);
        Assert.Equal(2, page.Shapes.Count);
        Assert.Equal("scaled", page.Shapes[0].Id);
        xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        Assert.NotEmpty(xml.Descendants(ns + "Page").Single(p => (string?)p.Attribute("Name") == "Scaled").Descendants(ns + "EllipticalArcTo"));
    }

    [Fact]
    public void NestedGroupCreationRemapsLocalFormulaReferencesAndPreservesQuotedText() {
        var document = VisioDocument.Create();
        var blueprint = new VisioShape("7", 2, 1, 4, 2, "root");
        var child = new VisioShape("8", 1, .5, 2, 1, "child");
        child.SetUserCell("WidthCopy", "4", formula: "Sheet.7!Width");
        child.SetUserCell("Literal", "Sheet.7!Width", formula: "\"Sheet.7!Width\"");
        child.Children.Add(new VisioShape("9", .5, .25, 1, .5, "nested"));
        blueprint.Children.Add(child);
        var master = document.RegisterMaster("Nested", blueprint);
        var page = document.AddPage("Page");
        page.AddRectangle(1, 1, 1, 1, "existing");
        var first = page.AddShape("instance", master, 5, 5, 8, 4);
        Assert.Single(first.Children);
        Assert.Equal(4, first.Children[0].Width);
        Assert.Equal(2, first.Children[0].PinX);
        Assert.Equal(2, first.Children[0].Children[0].Width);
        var xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        XNamespace ns = xml.Root!.Name.Namespace;
        var root = xml.Descendants(ns + "Page").Descendants(ns + "Shape").Single(s => (string?)s.Element(ns + "Text") == "root");
        string nativeId = (string)root.Attribute("ID")!;
        Assert.Equal("Sheet." + nativeId + "!Width", (string?)root.Descendants(ns + "User").Single(r => (string?)r.Attribute("NameU") == "WidthCopy").Element(ns + "Value")!.Attribute("F"));
        Assert.Equal("\"Sheet.7!Width\"", (string?)root.Descendants(ns + "User").Single(r => (string?)r.Attribute("NameU") == "Literal").Element(ns + "Value")!.Attribute("F"));
        page.Shapes.Insert(0, new VisioShape(nativeId, 1, 1, 1, 1, "new"));
        var loaded = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        var instance = loaded.Pages[0].FindShapeById("instance")!;
        Assert.Equal(nativeId, instance.PersistedId);
        Assert.Equal("Sheet." + nativeId + "!Width", instance.Children[0].UserCells.Single(r => r.Name == "WidthCopy").Formula);
        Assert.Single(loaded.Masters.Single().Shape.Children);
        Assert.Single(instance.Children[0].Children);
    }
}
