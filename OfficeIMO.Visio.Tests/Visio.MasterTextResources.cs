using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioMasterTextResourceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LoadedMasterInstancesSharePreservedTextUntilItIsEdited(bool package) {
        var document = Load(package);
        VisioShape master = document.Masters.Single().Shape;
        var shapes = document.Pages[0].Shapes;
        Assert.NotNull(master.PreservedTextElement);
        Assert.Equal(16, shapes.Count);
        Assert.Equal(32, master.PreservedTextElement!.Elements().Count());
        // Retaining one XML tree per master prevents instance-count × text-node expansion.
        Assert.All(shapes, shape => Assert.Same(master.PreservedTextElement, shape.PreservedTextElement));
        Assert.All(shapes, shape => Assert.Equal(new string('x', 32), shape.Text));

        shapes[0].Text = "Edited instance";
        foreach (var reopened in new[] {
            VisioDocument.Load(new MemoryStream(document.ToBytes())),
            VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value
        }) {
            Assert.Equal("Edited instance", reopened.Pages[0].Shapes[0].Text);
            Assert.Equal(new string('x', 32), reopened.Masters.Single().Shape.Text);
            Assert.All(reopened.Pages[0].Shapes.Skip(1), shape => Assert.Equal(new string('x', 32), shape.Text));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FormulaRemappingDetachesEitherTheMasterOrAnInstance(bool master) {
        var document = Load(false);
        var source = document.Masters.Single().Shape;
        var first = document.Pages[0].Shapes[0];
        var peer = document.Pages[0].Shapes[1];
        VisioShape changed = master ? source : first;

        VisioShapeFormulaReferences.Remap(changed, new Dictionary<string, string> { ["7"] = "70" });

        Assert.Equal(32, changed.PreservedTextElement!.Descendants().Attributes("F").Count());
        Assert.All(changed.PreservedTextElement!.Descendants().Attributes("F"), formula => Assert.Equal("Sheet.70!Width", formula.Value));
        foreach (var untouched in new[] { source, first, peer }.Where(shape => !ReferenceEquals(shape, changed))) {
            Assert.Equal(32, untouched.PreservedTextElement!.Descendants().Attributes("F").Count());
            Assert.All(untouched.PreservedTextElement!.Descendants().Attributes("F"), formula => Assert.Equal("Sheet.7!Width", formula.Value));
        }
    }

    private static VisioDocument Load(bool package) {
        const string ns = "http://schemas.microsoft.com/visio/2003/core";
        string text = string.Concat(Enumerable.Repeat("<cp IX='0' F='Sheet.7!Width'/>x", 32));
        string shapes = string.Concat(Enumerable.Range(10, 16).Select(id => "<Shape ID='" + id + "' Master='1'/>"));
        string xml = "<VisioDocument xmlns='" + ns + "'><Masters><Master ID='1' NameU='Text'><Shapes><Shape ID='7'>" +
            "<XForm><Width>2</Width><Height>1</Height></XForm><Text>" + text + "</Text></Shape></Shapes></Master></Masters>" +
            "<Pages><Page ID='1'><Shapes>" + shapes + "</Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
        return package ? VisioDocument.Load(new MemoryStream(document.ToBytes())) : document;
    }
}
