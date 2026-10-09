using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgImportTests {
    [Fact]
    public void Supported256DefinitionGraphicParentChainPreservesInheritedFillAcrossImportAndBothContainers() {
        const int definitionCount = 256;
        var source = OdgDocument.Create(); var page = source.AddPage();
        var shape = page.Shapes.AddRectangle(Rect(0, 0, 40, 30));
        for (int index = 0; index < definitionCount; index++) {
            var style = source.Styles.CreateNamed("Depth" + index, OdfStyleFamily.Graphic,
                index == definitionCount - 1 ? null : "Depth" + (index + 1));
            if (index == definitionCount - 1) SetFill(style, "#2468AC");
        }
        // Bind directly so the 256 definitions are the entire dependency chain.
        shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "Depth0");
        source.MarkPartDirty("content.xml");
        var expected = OdfColor.Parse("2468AC");
        Assert.Equal(expected, shape.FillColor);
        string[] sourceBefore = Parts(source);

        var destination = OdgDocument.Create();
        var imported = destination.ImportPage(source, 0);
        Assert.Equal(expected, imported.Shapes[0].FillColor);
        foreach (var read in new[] { destination }.Concat(RoundTrips(destination))) {
            var persisted = read.Pages[0].Shapes[0];
            string styleName = (string)persisted.Element.Attribute(OdfNamespaces.Draw + "style-name")!;
            var style = read.Styles.FindInPart(OdfStyleFamily.Graphic, styleName, "content.xml")!;
            Assert.Equal(definitionCount, read.Styles.Resolve(style).Count);
            Assert.Equal(expected, persisted.FillColor);
        }
        Assert.Equal(sourceBefore, Parts(source));
    }
}
