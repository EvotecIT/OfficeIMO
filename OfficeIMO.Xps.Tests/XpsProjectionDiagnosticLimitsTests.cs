using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Xps;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsProjectionDiagnosticLimitsTests {
    [Theory]
    [InlineData("model")]
    [InlineData("document")]
    [InlineData("chunks")]
    public void RepeatedPageLinkDiagnosticsShareTheProjectionTextLimit(string entryPoint) {
        var document = WithRepeatedLink("data:application/octet-stream," + new string('A', 1_000_000), 17);
        byte[] bytes = document.Save();
        var loaded = XpsDocument.Load(bytes);
        Assert.Equal(17, loaded.Pages.Count);
        Assert.Same(loaded.Pages[0], loaded.Pages[16]);
        var reader = new OfficeDocumentReaderBuilder().AddXpsHandler().Build();

        var error = Assert.Throws<InvalidDataException>(() => {
            if (entryPoint == "model") loaded.ToOfficeDocumentModel();
            else if (entryPoint == "document") reader.ReadDocument(bytes, "links.xps");
            else reader.Read(bytes, "links.xps").ToArray();
        });
        Assert.Contains("content-structure text limit", error.Message);
    }

    [Fact]
    public void OrdinaryUnresolvedLinksRetainTheirMessageAndPageLocations() {
        const string target = "data:text/plain,unresolved";
        var model = WithRepeatedLink(target, 2).ToOfficeDocumentModel();
        var diagnostics = model.Diagnostics.Where(item => item.Code == "XpsNavigationUnavailable").ToArray();

        Assert.Equal(new int?[] { 1, 2 }, diagnostics.Select(item => item.Location!.Page));
        Assert.All(diagnostics, item => Assert.Equal("Unresolved or unsupported native link: " + target, item.Message));
    }

    private static XpsDocument WithRepeatedLink(string target, int pageCount) {
        var document = XpsDocument.Create();
        var page = document.AddPage(100, 100);
        var markup = page.GetMarkup();
        markup.Add(new XElement(markup.Name.Namespace + "Path", new XAttribute("Data", "M0,0 L10,10"),
            new XAttribute("FixedPage.NavigateUri", target)));
        page.ReplaceMarkup(markup);
        for (int index = 1; index < pageCount; index++) document.Documents[0].InsertPage(index, page);
        return document;
    }
}
