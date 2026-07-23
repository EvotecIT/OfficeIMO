using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReadStreamTests {
    [Fact]
    public void NameTreeBudgetCountsAnIndirectDictionaryOnce() {
        var page = new PdfDictionary();
        page.Items["Type"] = new PdfName("Page");

        var destination = new PdfArray();
        destination.Items.Add(new PdfReference(3, 0));
        destination.Items.Add(new PdfName("Fit"));

        var names = new PdfArray();
        names.Items.Add(new PdfStringObj("Chapter1"));
        names.Items.Add(destination);

        var tree = new PdfDictionary();
        tree.Items["Names"] = names;
        var objects = new Dictionary<int, PdfIndirectObject> {
            [3] = new PdfIndirectObject(3, 0, page),
            [5] = new PdfIndirectObject(5, 0, tree)
        };

        bool supported = PdfPageExtractor.TryBuildFlattenedNamedDestinationNameTree(
            objects,
            new PdfReference(5, 0),
            copiedPageObjectIds: null,
            out PdfDictionary flattened,
            maximumNodes: 1);

        Assert.True(supported);
        Assert.True(flattened.Items.ContainsKey("Names"));
    }

    [Fact]
    public void NameTreeBudgetDoesNotCountADirectRootDictionary() {
        var page = new PdfDictionary();
        page.Items["Type"] = new PdfName("Page");

        var destination = new PdfArray();
        destination.Items.Add(new PdfReference(3, 0));
        destination.Items.Add(new PdfName("Fit"));

        var names = new PdfArray();
        names.Items.Add(new PdfStringObj("Chapter1"));
        names.Items.Add(destination);

        var leaf = new PdfDictionary();
        leaf.Items["Names"] = names;
        var kids = new PdfArray();
        kids.Items.Add(new PdfReference(5, 0));
        var root = new PdfDictionary();
        root.Items["Kids"] = kids;
        var objects = new Dictionary<int, PdfIndirectObject> {
            [3] = new PdfIndirectObject(3, 0, page),
            [5] = new PdfIndirectObject(5, 0, leaf)
        };

        bool supported = PdfPageExtractor.TryBuildFlattenedNamedDestinationNameTree(
            objects,
            root,
            copiedPageObjectIds: null,
            out PdfDictionary flattened,
            maximumNodes: 1);

        Assert.True(supported);
        Assert.True(flattened.Items.ContainsKey("Names"));
    }

    [Fact]
    public void RewriteApis_PreserveDirectNamedDestinationsForCopiedPages() {
        byte[] namedDestinationPdf = BuildNamedDestinationPdf();
        byte[] twoPageNamedDestinationPdf = BuildTwoPageNamedDestinationPdf();

        AssertNamedDestinations(PdfPageExtractor.ExtractPages(namedDestinationPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(namedDestinationPdf);
        Assert.Single(splitPages);
        AssertNamedDestinations(splitPages[0]);
        AssertNamedDestinations(PdfPageEditor.DeletePages(twoPageNamedDestinationPdf, 2), containsSecondDestination: false);
        AssertNamedDestinations(PdfPageEditor.ReorderPages(twoPageNamedDestinationPdf, 2, 1), containsSecondDestination: true);
        AssertNamedDestinations(PdfPageEditor.RotatePages(namedDestinationPdf, 90));
        AssertNamedDestinations(PdfMetadataEditor.UpdateMetadata(namedDestinationPdf, title: "Updated"));
        AssertNamedDestinations(PdfMerger.Merge(namedDestinationPdf));
        AssertNamedDestinations(PdfStamper.StampText(namedDestinationPdf, "STAMP"));

        static void AssertNamedDestinations(byte[] output, bool containsSecondDestination = false) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Dests ", text, StringComparison.Ordinal);
            Assert.Contains("/Chapter1 [", text, StringComparison.Ordinal);
            Assert.Contains("/XYZ 0 200 0", text, StringComparison.Ordinal);
            if (containsSecondDestination) {
                Assert.Contains("/Chapter2 [", text, StringComparison.Ordinal);
                Assert.Contains("/Fit", text, StringComparison.Ordinal);
            } else {
                Assert.DoesNotContain("/Chapter2 [", text, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void RewriteApis_PreserveNamedDestinationNameTreesForCopiedPages() {
        byte[] namedDestinationPdf = BuildNamedDestinationNameTreePdf();
        byte[] twoPageNamedDestinationPdf = BuildTwoPageNamedDestinationNameTreePdf();

        AssertNamedDestinations(PdfPageExtractor.ExtractPages(namedDestinationPdf, 1));
        var splitPages = PdfPageExtractor.SplitPages(namedDestinationPdf);
        Assert.Single(splitPages);
        AssertNamedDestinations(splitPages[0]);
        AssertNamedDestinations(PdfPageEditor.DeletePages(twoPageNamedDestinationPdf, 2), containsSecondDestination: false);
        AssertNamedDestinations(PdfPageEditor.ReorderPages(twoPageNamedDestinationPdf, 2, 1), containsSecondDestination: true);
        AssertNamedDestinations(PdfPageEditor.RotatePages(namedDestinationPdf, 90));
        AssertNamedDestinations(PdfMetadataEditor.UpdateMetadata(namedDestinationPdf, title: "Updated"));
        AssertNamedDestinations(PdfMerger.Merge(namedDestinationPdf));
        AssertNamedDestinations(PdfStamper.StampText(namedDestinationPdf, "STAMP"));

        static void AssertNamedDestinations(byte[] output, bool containsSecondDestination = false) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Names << /Dests << /Names [", text, StringComparison.Ordinal);
            Assert.Contains("(Chapter1)", text, StringComparison.Ordinal);
            Assert.Contains("/XYZ 0 200 0", text, StringComparison.Ordinal);
            if (containsSecondDestination) {
                Assert.Contains("(Chapter2)", text, StringComparison.Ordinal);
                Assert.Contains("/Fit", text, StringComparison.Ordinal);
            } else {
                Assert.DoesNotContain("(Chapter2)", text, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void RewriteApis_PreserveNamedDestinationNameTreeKidsForCopiedPages() {
        byte[] namedDestinationPdf = BuildNamedDestinationNameTreeWithKidsPdf();
        byte[] twoPageNamedDestinationPdf = BuildTwoPageNamedDestinationNameTreeWithKidsPdf();

        AssertNamedDestinations(PdfPageExtractor.ExtractPages(namedDestinationPdf, 1));
        AssertNamedDestinations(PdfPageEditor.DeletePages(twoPageNamedDestinationPdf, 2), containsSecondDestination: false);
        AssertNamedDestinations(PdfPageEditor.ReorderPages(twoPageNamedDestinationPdf, 2, 1), containsSecondDestination: true, chapter1Page: 2, chapter2Page: 1);
        AssertNamedDestinations(PdfMetadataEditor.UpdateMetadata(namedDestinationPdf, title: "Updated"));

        static void AssertNamedDestinations(byte[] output, bool containsSecondDestination = false, int chapter1Page = 1, int chapter2Page = 2) {
            string text = System.Text.Encoding.ASCII.GetString(output);
            Assert.Contains("/Names << /Dests << /Names [", text, StringComparison.Ordinal);
            Assert.DoesNotContain("/Dests << /Kids", text, StringComparison.Ordinal);

            PdfDocumentInfo info = PdfInspector.Inspect(output);
            AssertDestination(info, "Chapter1", chapter1Page, 200);
            if (containsSecondDestination) {
                AssertDestination(info, "Chapter2", chapter2Page, null, PdfOpenActionDestinationMode.Fit);
            } else {
                Assert.DoesNotContain(info.NamedDestinations, destination => destination.Name == "Chapter2");
            }
        }

        static void AssertDestination(PdfDocumentInfo info, string name, int pageNumber, double? top, PdfOpenActionDestinationMode? mode = PdfOpenActionDestinationMode.Xyz) {
            PdfNamedDestination destination = Assert.Single(info.NamedDestinations, item => item.Name == name);
            Assert.Equal(pageNumber, destination.PageNumber);
            Assert.Equal(top, destination.DestinationTop);
            Assert.Equal(mode, destination.DestinationMode);
        }
    }

    [Fact]
    public void ReadApis_ResolveDirectGoToLinkAnnotationDestinations() {
        byte[] linkedPdf = BuildGoToActionLinkAnnotationPdf();

        PdfReadDocument document = PdfReadDocument.Open(new MemoryStream(linkedPdf));
        PdfLinkAnnotation pageLink = Assert.Single(document.Pages[0].GetLinkAnnotations());
        Assert.Null(pageLink.PageNumber);
        Assert.Null(pageLink.DestinationPageNumber);
        Assert.Null(pageLink.DestinationName);
        Assert.Null(pageLink.Uri);
        Assert.True(pageLink.IsInternalDestinationLink);
        Assert.False(pageLink.IsNamedDestinationLink);
        Assert.False(pageLink.IsUriLink);
        Assert.Equal(144d, pageLink.DestinationTop);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, pageLink.DestinationMode);

        PdfDocumentInfo info = PdfInspector.Inspect(linkedPdf);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
        Assert.Equal(1, link.PageNumber);
        Assert.Equal(1, link.DestinationPageNumber);
        Assert.Null(link.DestinationName);
        Assert.Null(link.Uri);
        Assert.True(link.IsInternalDestinationLink);
        Assert.False(link.IsNamedDestinationLink);
        Assert.False(link.IsUriLink);
        Assert.Equal("Jump to top", link.Contents);
        Assert.Equal(144d, link.DestinationTop);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, link.DestinationMode);
    }

    [Fact]
    public void ReadApis_ResolveExtendedDirectGoToLinkAnnotationDestinations() {
        AssertLinkDestination(
            BuildGoToActionLinkAnnotationPdf("[3 0 R /FitV 24]"),
            PdfOpenActionDestinationMode.FitVertical,
            destinationLeft: 24D);

        AssertLinkDestination(
            BuildGoToActionLinkAnnotationPdf("[3 0 R /FitR 10 20 90 144]"),
            PdfOpenActionDestinationMode.FitRectangle,
            destinationLeft: 10D,
            destinationBottom: 20D,
            destinationRight: 90D,
            destinationTop: 144D);

        AssertLinkDestination(
            BuildGoToActionLinkAnnotationPdf("[3 0 R /FitB]"),
            PdfOpenActionDestinationMode.FitBoundingBox);

        AssertLinkDestination(
            BuildGoToActionLinkAnnotationPdf("[3 0 R /FitBH 155]"),
            PdfOpenActionDestinationMode.FitBoundingBoxHorizontal,
            destinationTop: 155D);

        AssertLinkDestination(
            BuildGoToActionLinkAnnotationPdf("[3 0 R /FitBV 33]"),
            PdfOpenActionDestinationMode.FitBoundingBoxVertical,
            destinationLeft: 33D);

        static void AssertLinkDestination(
            byte[] pdf,
            PdfOpenActionDestinationMode expectedMode,
            double? destinationLeft = null,
            double? destinationBottom = null,
            double? destinationRight = null,
            double? destinationTop = null) {
            PdfDocumentInfo info = PdfInspector.Inspect(pdf);
            PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
            Assert.Equal(1, link.DestinationPageNumber);
            Assert.Equal(expectedMode, link.DestinationMode);
            Assert.Equal(destinationLeft, link.DestinationLeft);
            Assert.Equal(destinationBottom, link.DestinationBottom);
            Assert.Equal(destinationRight, link.DestinationRight);
            Assert.Equal(destinationTop, link.DestinationTop);
        }
    }

    [Fact]
    public void RewriteApis_PreserveSimpleLinkAnnotationsWithContentsMetadata() {
        byte[] linkedPdf = BuildTwoPageLinkAnnotationPdf();

        AssertSinglePageLink(PdfPageExtractor.ExtractPages(linkedPdf, 1), "https://evotec.xyz/first", "First link metadata");
        var splitPages = PdfPageExtractor.SplitPages(linkedPdf);
        Assert.Equal(2, splitPages.Count);
        AssertSinglePageLink(splitPages[0], "https://evotec.xyz/first", "First link metadata");
        AssertSinglePageLink(splitPages[1], "https://evotec.xyz/second", "Second link metadata");

        AssertSinglePageLink(PdfPageEditor.DeletePages(linkedPdf, 2), "https://evotec.xyz/first", "First link metadata");
        AssertTwoPageLinks(
            PdfPageEditor.ReorderPages(linkedPdf, 2, 1),
            ("https://evotec.xyz/second", "Second link metadata"),
            ("https://evotec.xyz/first", "First link metadata"));
        AssertTwoPageLinks(
            PdfPageEditor.RotatePages(linkedPdf, 90),
            ("https://evotec.xyz/first", "First link metadata"),
            ("https://evotec.xyz/second", "Second link metadata"));
        AssertTwoPageLinks(
            PdfMetadataEditor.UpdateMetadata(linkedPdf, title: "Updated"),
            ("https://evotec.xyz/first", "First link metadata"),
            ("https://evotec.xyz/second", "Second link metadata"));
        AssertTwoPageLinks(
            PdfMerger.Merge(linkedPdf),
            ("https://evotec.xyz/first", "First link metadata"),
            ("https://evotec.xyz/second", "Second link metadata"));
        AssertTwoPageLinks(
            PdfStamper.StampText(linkedPdf, "STAMP"),
            ("https://evotec.xyz/first", "First link metadata"),
            ("https://evotec.xyz/second", "Second link metadata"));

        static void AssertSinglePageLink(byte[] output, string uri, string contents) {
            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.Single(info.Pages);
            AssertPageLink(info.Pages[0], uri, contents);
        }

        static void AssertTwoPageLinks(byte[] output, (string Uri, string Contents) first, (string Uri, string Contents) second) {
            PdfDocumentInfo info = PdfInspector.Inspect(output);
            Assert.Equal(2, info.PageCount);
            Assert.Equal(2, info.LinkAnnotationCount);
            Assert.Equal(2, info.LinkAnnotations.Count);
            Assert.Equal(2, info.LinkUriCount);
            Assert.Equal(new[] { first.Uri, second.Uri }, info.LinkUris);
            AssertPageLink(info.Pages[0], first.Uri, first.Contents);
            AssertPageLink(info.Pages[1], second.Uri, second.Contents);
            Assert.Equal(first.Uri, info.LinkAnnotations[0].Uri);
            Assert.Equal(1, info.LinkAnnotations[0].PageNumber);
            Assert.Equal(second.Uri, info.LinkAnnotations[1].Uri);
            Assert.Equal(2, info.LinkAnnotations[1].PageNumber);
        }

        static void AssertPageLink(PdfPageInfo page, string uri, string contents) {
            var link = Assert.Single(page.LinkAnnotations);
            Assert.Equal(page.PageNumber, link.PageNumber);
            Assert.Equal(uri, link.Uri);
            Assert.Equal(contents, link.Contents);
            Assert.True(link.Width > 0);
            Assert.True(link.Height > 0);
            Assert.InRange(link.X1, 0, page.Width);
            Assert.InRange(link.X2, 0, page.Width);
            Assert.InRange(link.Y1, 0, page.Height);
            Assert.InRange(link.Y2, 0, page.Height);
        }
    }

    [Fact]
    public void RewriteApis_RejectComplexNamedDestinationNameTreesWithClearUnsupportedDiagnostic() {
        byte[] namedDestinationPdf = BuildComplexNamedDestinationNameTreePdf();

        static void AssertNamedDestinations(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }

        AssertNamedDestinations(() => PdfPageExtractor.ExtractPages(namedDestinationPdf, 1));
        AssertNamedDestinations(() => PdfPageExtractor.SplitPages(namedDestinationPdf));
        AssertNamedDestinations(() => PdfPageEditor.DeletePages(namedDestinationPdf, 1));
        AssertNamedDestinations(() => PdfMetadataEditor.UpdateMetadata(namedDestinationPdf, title: "Updated"));
        AssertNamedDestinations(() => PdfMerger.Merge(namedDestinationPdf));
        AssertNamedDestinations(() => PdfStamper.StampText(namedDestinationPdf, "STAMP"));
    }


}
