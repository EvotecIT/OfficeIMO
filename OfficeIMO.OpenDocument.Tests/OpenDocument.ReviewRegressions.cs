using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentReviewRegressionTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public void RejectsDecodedSpaceRunsBeyondTheTextSafetyLimit() {
        using OdtDocument document = OdtDocument.Create();
        document.AddParagraph("placeholder");
        XDocument content = document.Package.GetXml("content.xml");
        XElement paragraph = content.Descendants(OdfNamespaces.Text + "p").Single();
        paragraph.ReplaceNodes(new XElement(OdfNamespaces.Text + "s",
            new XAttribute(OdfNamespaces.Text + "c", int.MaxValue)));
        document.Package.MarkXmlDirty("content.xml");

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => reopened.Paragraphs.Single().Text);
        Assert.Contains("safety limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ExistingHeaderParagraphEditsRewriteStylesPart() {
        using OdtDocument source = OdtDocument.Create();
        source.PageLayout.Header.AddParagraph("Before");
        byte[] original = source.ToBytes();

        using OdtDocument edited = OdtDocument.Open(new MemoryStream(original));
        edited.PageLayout.Header.Paragraphs.Single().Text = "After";
        byte[] output = edited.ToBytes();

        Assert.Contains("styles.xml", edited.LastSaveReport!.RewrittenEntries);
        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(output));
        Assert.Equal("After", reopened.PageLayout.Header.Paragraphs.Single().Text);
    }

    [Fact]
    public void DirectFormattingClonesSharedAutomaticStyles() {
        using OdtDocument document = OdtDocument.Create();
        OdtParagraph first = document.AddParagraph("First");
        OdtParagraph second = document.AddParagraph("Second");
        OdfStyle shared = document.Styles.CreateAutomatic(OdfStyleFamily.Paragraph, "shared");
        shared.Bold = false;
        first.StyleName = shared.Name;
        second.StyleName = shared.Name;

        first.Bold = true;

        Assert.True(first.Bold);
        Assert.False(second.Bold);
        Assert.NotEqual(first.StyleName, second.StyleName);
        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));
        Assert.True(reopened.Paragraphs[0].Bold);
        Assert.False(reopened.Paragraphs[1].Bold);
    }

    [Fact]
    public void ManifestValidationIgnoresUnlistedZipDirectoryEntries() {
        using OdtDocument document = OdtDocument.Create();
        document.Package.AddOrReplaceEntry("Configurations2/accelerator/", Array.Empty<byte>(), string.Empty);

        OdfValidationResult result = document.Validate();

        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Id == "ODF103");
    }

    [Fact]
    public void ImageBytesResolveRelativeAndEscapedPackageHrefs() {
        using OdtDocument document = OdtDocument.Create();
        OdtImage image = document.AddParagraph("Image").AddImage(TinyPng, "pixel.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        XElement imageElement = document.Package.GetXml("content.xml").Descendants(OdfNamespaces.Draw + "image").Single();
        imageElement.SetAttributeValue(OdfNamespaces.XLink + "href", "./" + image.Path.Replace(".", "%2E"));
        document.Package.MarkXmlDirty("content.xml");

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));

        Assert.Equal(TinyPng, reopened.Paragraphs.Single().Images.Single().GetImageBytes());
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void HeaderStylesResolveWithinStylesPartWhenNamesCollide() {
        using OdtDocument document = OdtDocument.Create();
        OdtParagraph body = document.AddParagraph("Body");
        OdtParagraph header = document.PageLayout.Header.AddParagraph("Header");
        body.StyleName = "P1";
        header.StyleName = "P1";
        AddParagraphStyle(document.Package.GetXml("content.xml"), "P1", "normal");
        AddParagraphStyle(document.Package.GetXml("styles.xml"), "P1", "bold");
        document.Package.MarkXmlDirty("content.xml");
        document.Package.MarkXmlDirty("styles.xml");

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));

        Assert.False(reopened.Paragraphs.Single().Bold);
        Assert.True(reopened.PageLayout.Header.Paragraphs.Single().Bold);
    }

    private static void AddParagraphStyle(XDocument document, string name, string weight) {
        XElement automatic = document.Root!.Element(OdfNamespaces.Office + "automatic-styles")!;
        automatic.Add(new XElement(OdfNamespaces.Style + "style",
            new XAttribute(OdfNamespaces.Style + "name", name),
            new XAttribute(OdfNamespaces.Style + "family", "paragraph"),
            new XElement(OdfNamespaces.Style + "text-properties",
                new XAttribute(OdfNamespaces.Fo + "font-weight", weight))));
    }

}
