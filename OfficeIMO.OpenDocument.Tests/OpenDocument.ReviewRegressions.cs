using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentReviewRegressionTests {
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

}
