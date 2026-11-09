using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsCreatedVisualTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void CreatedVisualsEditTheNativePageAndKeepSiblingHandlesLive(XpsFormat format) {
        var document = XpsDocument.Create(format);
        var page = document.AddPage(100, 100);
        XpsVisual first = page.AddPath("M0,0H20V20Z");
        XpsVisual second = page.AddPath("M30,30H50V50Z");
        Assert.Same(page, first.Page);
        var markup = first.GetMarkup();
        markup.SetAttributeValue("Fill", "#FFFF0000");
        markup.SetAttributeValue("Name", "edited");
        first.ReplaceMarkup(markup);
        Assert.Equal("#FFFF0000", (string?)first.GetMarkup().Attribute("Fill"));
        Assert.Equal("M30,30H50V50Z", (string?)second.GetMarkup().Attribute("Data"));
        var loaded = XpsDocument.Load(document.ToBytes());
        var saved = loaded.Pages[0].GetMarkup().Elements().First();
        Assert.Equal("edited", (string?)saved.Attribute("Name"));
        Assert.Equal("#FFFF0000", (string?)saved.Attribute("Fill"));
        Assert.Throws<ArgumentException>(() => first.ReplaceMarkup(new XElement("Path")));
        page.ReplaceMarkup(page.GetMarkup());
        Assert.Throws<InvalidOperationException>(() => first.GetMarkup());
    }
}
