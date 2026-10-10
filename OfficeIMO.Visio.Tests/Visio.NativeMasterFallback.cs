using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioNativeMasterFallbackTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void RootMastersWithoutGeometryRetainTheirPrimitiveRendering(bool reopenPackage) {
            const string xml = """
                <VisioDocument xmlns="http://schemas.microsoft.com/visio/2003/core">
                  <Masters><Master ID="1" NameU="Rectangle"><Shapes><Shape ID="1" NameU="Rectangle">
                    <XForm><Width>2</Width><Height>1</Height><LocPinX>1</LocPinX><LocPinY>0.5</LocPinY></XForm>
                    <Fill><FillForegnd>#00A000</FillForegnd><FillPattern>1</FillPattern></Fill>
                  </Shape></Shapes></Master></Masters>
                  <Pages><Page ID="0"><PageSheet><PageProps><PageWidth>5</PageWidth><PageHeight>3</PageHeight></PageProps></PageSheet>
                    <Shapes><Shape ID="1" Master="1"><XForm><PinX>2.5</PinX><PinY>1.5</PinY></XForm></Shape></Shapes>
                  </Page></Pages>
                </VisioDocument>
                """;
            using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            VisioDocument document = VisioDocument.LoadLegacyXml(input).Value;
            if (reopenPackage) {
                using MemoryStream package = new MemoryStream();
                document.Save(package);
                document = VisioDocument.Load(package);
            }
            VisioPage page = Assert.Single(document.Pages);
            VisioShape shape = Assert.Single(page.Shapes);
            Assert.Same(shape.Master!.Shape, shape.MasterShape);
            Assert.Empty(shape.PreservedGeometrySections);
            Assert.Empty(shape.Master.Shape.PreservedGeometrySections);
            XDocument svg = XDocument.Parse(page.ToSvg());
            XNamespace ns = svg.Root!.Name.Namespace;
            XElement group = Assert.Single(svg.Descendants(ns + "g"), element =>
                (string?)element.Attribute("data-visio-shape-id") == "1");
            Assert.Contains(group.Elements(), element => element.Name.LocalName is "rect" or "path");
            Assert.NotEmpty(page.ToDrawing().Value.Elements);
        }
    }
}
