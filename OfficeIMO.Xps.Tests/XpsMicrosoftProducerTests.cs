using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsMicrosoftProducerTests {
    [Theory]
    [InlineData("PrintingDrt.xps", 450, 450, "")]
    [InlineData("Test_Document.xps", 816, 1056, "This is a test XPS file.")]
    [InlineData("word.xps", 816, 1056, "These are just some words printed")]
    public void MicrosoftPagesAndOpaqueResourcesSurviveNativePageEdits(string file, double width, double height, string text) {
        var document = XpsDocument.Load(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "MicrosoftWpf", file)));
        Assert.Single(document.Documents);
        var page = Assert.Single(document.Pages);
        Assert.Equal(width, page.Width);
        Assert.Equal(height, page.Height);
        if (text.Length == 0) Assert.Empty(page.ExtractText());
        else Assert.Contains(text, page.ExtractText());
        var markup = page.GetMarkup();
        var svg = page.ToSvg();
        Assert.Empty(svg.Diagnostics);
        // These are independently serialized embedded fonts, print tickets and
        // thumbnail resources; synthetic round trips do not cover their bytes.
        var resources = document.PartNames.Where(name =>
            name.EndsWith(".odttf", StringComparison.OrdinalIgnoreCase) ||
            name.EndsWith("_PT.xml", StringComparison.OrdinalIgnoreCase) ||
            name.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(name => name, document.GetPartBytes);
        document.Documents[0].InsertPage(0, 100, 80).AddPath("M0,0H20V20Z");
        var reopened = XpsDocument.Load(document.Save());
        Assert.Equal(2, reopened.Pages.Count);
        Assert.True(XNode.DeepEquals(markup, reopened.Pages[1].GetMarkup()));
        Assert.Equal(page.ExtractText(), reopened.Pages[1].ExtractText());
        Assert.Equal(svg.Svg, reopened.Pages[1].ToSvg().Svg);
        foreach (var resource in resources) Assert.Equal(resource.Value, reopened.GetPartBytes(resource.Key));
    }
}
