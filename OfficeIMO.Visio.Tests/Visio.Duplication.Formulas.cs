using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioDuplicationFormulaTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    public void CopiedGraphsUseStableNativeReferencesWithoutChangingOriginals(bool pageCopy, bool namedIds, bool package) {
        XNamespace legacy = "http://schemas.microsoft.com/visio/2003/core";
        string source = $"<VisioDocument xmlns='{legacy}'><Pages><Page ID='0' Name='Source'><Shapes>"
            + "<Shape ID='1' NameU='Root' Type='Group'><XForm><PinX>2</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm>"
            + "<Char IX='7'><Size Unit='PT' F='GUARD(Sheet.2!Width*6 pt)'>0.16666666666666667</Size></Char>"
            + "<User ID='0' NameU='Link'><Value F='Sheet.2!Width'>1</Value></User>"
            + "<User ID='1' NameU='Quoted'><Value F='&quot;Sheet.2!Width&quot;'/></User>"
            + "<Scratch IX='0'><X F='Sheet.4!Width+OtherPage!Sheet.2!Width+Sheet.999!Width'>1</X></Scratch>"
            + "<Text><cp IX='7'/>Root</Text><Shapes><Shape ID='002' NameU='Child'><XForm><Width>1</Width><Height>1</Height></XForm>"
            + "<Prop ID='0' NameU='Parent'><Value F='Sheet.1!Width'>2</Value></Prop><Text>Child</Text></Shape></Shapes></Shape>"
            + "<Shape ID='3' NameU='Other'><XForm><PinX>4</PinX><PinY>1</PinY><Width>1</Width><Height>1</Height></XForm><Text>Other</Text></Shape>"
            + "<Shape ID='4'><XForm><Width>2</Width><Height>0</Height></XForm><XForm1D><BeginX>2</BeginX><BeginY>1</BeginY><EndX>4</EndX><EndY>1</EndY></XForm1D>"
            + "<Char IX='7'><Size Unit='PT' F='GUARD(Sheet.2!Width*6 pt)'>0.16666666666666667</Size></Char>"
            + "<Para IX='4'><HorzAlign F='Sheet.1!User.Align'>0</HorzAlign></Para>"
            + "<Prop ID='0' NameU='Target'><Value F='Sheet.3!Width'>1</Value></Prop>"
            + "<Scratch IX='0'><X F='Sheet.1!Width'>2</X></Scratch><Text><cp IX='7'/><pp IX='4'/>Edge</Text></Shape>"
            + "</Shapes><Connects><Connect FromSheet='4' FromCell='BeginX' ToSheet='1' ToCell='PinX'/><Connect FromSheet='4' FromCell='EndX' ToSheet='3' ToCell='PinX'/></Connects></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        var originalPage = document.Pages[0];
        VisioPage target;
        VisioShape root, other;
        if (pageCopy) {
            target = document.DuplicatePage(originalPage, "Copy");
            root = target.Shapes[0]; other = target.Shapes[1];
        } else {
            target = originalPage;
            var copies = target.DuplicateShapes(target.Shapes.ToArray(), new VisioShapeDuplicationOptions {
                IdSuffix = namedIds ? "-copy" : null, OffsetX = 0, OffsetY = 0
            });
            root = copies[0]; other = copies[1];
        }
        root.NameU = "CopiedRoot"; root.Children[0].NameU = "CopiedChild"; other.NameU = "CopiedOther";
        var edge = target.Connectors.Last(); edge.Label = "CopiedEdge";
        // A later insertion/reorder must not reuse a native ID already referenced by copied formulas.
        target.Shapes.Insert(0, new VisioShape("inserted", 0, 0, 1, 1, "Inserted"));
        target.Shapes.Remove(root); target.Shapes.Insert(0, root);
        var candidate = package ? VisioDocument.Load(new MemoryStream(document.ToBytes())) : document;
        var xml = XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value));
        var copyPage = xml.Descendants(legacy + "Page").Single(p => (string?)p.Attribute("Name") == target.Name);
        var copiedRoot = Shape("CopiedRoot"); var copiedChild = Shape("CopiedChild"); var copiedOther = Shape("CopiedOther");
        var copiedEdge = copyPage.Descendants(legacy + "Shape").Single(s => s.Element(legacy + "Text")?.Value == "CopiedEdge");
        string rootId = Id(copiedRoot), childId = Id(copiedChild), otherId = Id(copiedOther), edgeId = Id(copiedEdge);
        Assert.Equal($"GUARD(Sheet.{childId}!Width*6 pt)", F(copiedRoot.Element(legacy + "Char")!.Element(legacy + "Size")!));
        Assert.Equal($"GUARD(Sheet.{childId}!Width*6 pt)", F(copiedEdge.Element(legacy + "Char")!.Element(legacy + "Size")!));
        Assert.Equal($"Sheet.{rootId}!User.Align", F(copiedEdge.Element(legacy + "Para")!.Element(legacy + "HorzAlign")!));
        Assert.Equal($"Sheet.{childId}!Width", F(copiedRoot.Elements(legacy + "User").Single(r => (string?)r.Attribute("NameU") == "Link").Element(legacy + "Value")!));
        Assert.Equal("\"Sheet.2!Width\"", F(copiedRoot.Elements(legacy + "User").Single(r => (string?)r.Attribute("NameU") == "Quoted").Element(legacy + "Value")!));
        Assert.Equal($"Sheet.{rootId}!Width", F(copiedChild.Element(legacy + "Prop")!.Element(legacy + "Value")!));
        Assert.Equal($"Sheet.{otherId}!Width", F(copiedEdge.Element(legacy + "Prop")!.Element(legacy + "Value")!));
        Assert.Equal($"Sheet.{edgeId}!Width+OtherPage!Sheet.2!Width+Sheet.999!Width", F(copiedRoot.Element(legacy + "Scratch")!.Element(legacy + "X")!));
        Assert.Equal($"Sheet.{rootId}!Width", F(copiedEdge.Element(legacy + "Scratch")!.Element(legacy + "X")!));
        var original = xml.Descendants(legacy + "Page").Single(p => (string?)p.Attribute("Name") == "Source")
            .Descendants(legacy + "Shape").Single(s => (string?)s.Attribute("NameU") == "Root");
        Assert.Equal("GUARD(Sheet.2!Width*6 pt)", F(original.Element(legacy + "Char")!.Element(legacy + "Size")!));

        XElement Shape(string name) => copyPage.Descendants(legacy + "Shape").Single(s => (string?)s.Attribute("NameU") == name);
        string Id(XElement element) => (string)element.Attribute("ID")!;
        string? F(XElement element) => (string?)element.Attribute("F");
    }
}
