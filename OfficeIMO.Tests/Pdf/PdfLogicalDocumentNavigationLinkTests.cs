using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfLogicalDocumentTests {
    [Fact]
    public void Load_ExposesDocumentNavigationObjects() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildNavigationPdf());

        Assert.Equal("FullScreen", logical.CatalogPageMode);
        Assert.Equal("TwoColumnLeft", logical.CatalogPageLayout);
        Assert.Equal("1.7", logical.CatalogVersion);
        Assert.Equal("en-US", logical.CatalogLanguage);

        Assert.True(logical.HasOutlines);
        PdfOutlineItem outline = Assert.Single(logical.Outlines);
        Assert.Equal("Logical outline", outline.Title);
        Assert.Equal(1, outline.PageNumber);

        Assert.True(logical.HasReadablePageLabels);
        PdfPageLabel label = Assert.Single(logical.PageLabels);
        Assert.Equal(0, label.StartPageIndex);
        Assert.Equal("D", label.Style);
        Assert.Equal("A-", label.Prefix);
        Assert.Equal(3, label.StartNumber);

        Assert.True(logical.HasNamedDestinations);
        PdfNamedDestination destination = Assert.Single(logical.NamedDestinations);
        Assert.Equal("Chapter1", destination.Name);
        Assert.Equal(1, destination.PageNumber);

        Assert.True(logical.HasReadableOpenAction);
        Assert.NotNull(logical.OpenAction);
        Assert.Equal("Destination", logical.OpenAction!.ActionType);
        Assert.Equal(1, logical.OpenAction.PageNumber);

        Assert.True(logical.HasReadableViewerPreferences);
        Assert.NotNull(logical.ViewerPreferences);
        Assert.True(logical.ViewerPreferences!.GetBoolean("HideToolbar"));
        Assert.True(logical.ViewerPreferences.GetBoolean("DisplayDocTitle"));
    }

    [Fact]
    public void Load_ExposesLinkAnnotationsAsLogicalElements() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 240,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .H1("Linked heading", linkUri: "https://evotec.xyz/logical-link", linkContents: "Logical link metadata")
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        PdfLogicalPage page = Assert.Single(logical.Pages);

        PdfLogicalLinkAnnotation link = Assert.Single(page.Links);
        Assert.Equal(1, link.PageNumber);
        Assert.True(link.IsUriLink);
        Assert.False(link.IsNamedDestinationLink);
        Assert.Equal("https://evotec.xyz/logical-link", link.Uri);
        Assert.Equal("Logical link metadata", link.Contents);
        Assert.True(link.Width > 0);
        Assert.True(link.Height > 0);
        Assert.Equal(1, link.SourceLink.PageNumber);
        Assert.True(logical.HasLinks);
        Assert.Same(link, Assert.Single(logical.Links));
        Assert.Same(link, Assert.Single(logical.LinksByUri["https://evotec.xyz/logical-link"]));
        Assert.Same(link, Assert.Single(logical.GetLinksByUri("https://evotec.xyz/logical-link")));
        Assert.Empty(logical.GetLinksByUri("https://evotec.xyz/missing"));
        Assert.Empty(logical.GetLinksByDestinationName("Missing"));
        Assert.Contains(page.Elements, element => element.Kind == PdfLogicalElementKind.LinkAnnotation);
        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.LinkAnnotation);
    }

    [Fact]
    public void Load_ExposesHeadingBookmarkLinksAsLogicalElements() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 240,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .H1("Jump to details", linkDestinationName: "Details", linkContents: "Heading jump metadata")
            .Spacer(18)
            .Bookmark("Details")
            .H2("Details")
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        Assert.Contains(logical.Headings, heading => heading.Text == "Jump to details");
        Assert.Contains(logical.NamedDestinations, destination => destination.Name == "Details");
        PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByDestinationName("Details"));
        Assert.False(link.IsUriLink);
        Assert.True(link.IsNamedDestinationLink);
        Assert.Null(link.Uri);
        Assert.Equal("Details", link.DestinationName);
        Assert.Equal("Heading jump metadata", link.Contents);
        Assert.True(link.Width > 0);
        Assert.True(link.Height > 0);
    }

    [Fact]
    public void Load_ExposesDirectDestinationLinkCoordinatesAsLogicalElements() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildDirectDestinationLinkPdf());

        PdfLogicalLinkAnnotation link = Assert.Single(logical.Links);
        Assert.True(link.IsInternalDestinationLink);
        Assert.False(link.IsUriLink);
        Assert.False(link.IsNamedDestinationLink);
        Assert.Equal(1, link.PageNumber);
        Assert.Equal(1, link.DestinationPageNumber);
        Assert.Equal(PdfOpenActionDestinationMode.FitRectangle, link.DestinationMode);
        Assert.Equal(10D, link.DestinationLeft);
        Assert.Equal(20D, link.DestinationBottom);
        Assert.Equal(90D, link.DestinationRight);
        Assert.Equal(144D, link.DestinationTop);
        Assert.Equal("Direct destination link", link.Contents);
        Assert.Same(link, Assert.Single(logical.LinksByDestinationPageNumber[1]));
        Assert.Same(link, Assert.Single(logical.GetLinksByDestinationPageNumber(1)));
        Assert.Empty(logical.GetLinksByDestinationPageNumber(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.GetLinksByDestinationPageNumber(0));
    }

    [Fact]
    public void Load_ExposesNamedActionLinksAsLogicalElements() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildNamedActionLinkPdf());

        PdfLogicalLinkAnnotation link = Assert.Single(logical.Links);
        Assert.True(link.IsNamedActionLink);
        Assert.False(link.IsUriLink);
        Assert.False(link.IsNamedDestinationLink);
        Assert.False(link.IsInternalDestinationLink);
        Assert.Equal(1, link.PageNumber);
        Assert.Null(link.Uri);
        Assert.Null(link.DestinationName);
        Assert.Null(link.DestinationPageNumber);
        Assert.Equal("NextPage", link.NamedAction);
        Assert.Equal("Next page action", link.Contents);
        Assert.Same(link, Assert.Single(logical.LinksByNamedAction["NextPage"]));
        Assert.Same(link, Assert.Single(logical.GetLinksByNamedAction("NextPage")));
        Assert.Empty(logical.GetLinksByNamedAction("PrevPage"));
        Assert.True(link.Width > 0);
        Assert.True(link.Height > 0);
    }

    [Fact]
    public void Load_ExposesRemoteGoToLinksAsLogicalElements() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildRemoteGoToLinkPdf());

        PdfLogicalLinkAnnotation link = Assert.Single(logical.Links);
        Assert.True(link.IsRemoteGoToLink);
        Assert.False(link.IsUriLink);
        Assert.False(link.IsNamedDestinationLink);
        Assert.False(link.IsInternalDestinationLink);
        Assert.False(link.IsNamedActionLink);
        Assert.Equal(1, link.PageNumber);
        Assert.Null(link.Uri);
        Assert.Null(link.DestinationName);
        Assert.Null(link.DestinationPageNumber);
        Assert.Null(link.NamedAction);
        Assert.Equal("remote-report.pdf", link.RemoteFile);
        Assert.Null(link.RemoteDestinationName);
        Assert.Equal(2, link.RemoteDestinationPageNumber);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, link.RemoteDestinationMode);
        Assert.Equal(144D, link.RemoteDestinationTop);
        Assert.Equal("Remote report link", link.Contents);
        Assert.Same(link, Assert.Single(logical.LinksByRemoteFile["remote-report.pdf"]));
        Assert.Same(link, Assert.Single(logical.GetLinksByRemoteFile("remote-report.pdf")));
        Assert.Empty(logical.GetLinksByRemoteFile("missing.pdf"));
        Assert.True(link.Width > 0);
        Assert.True(link.Height > 0);
    }

    [Fact]
    public void Load_ExposesTableCellNamedDestinationLinksAsLogicalElements() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 240,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .Table(new[] {
                new[] {
                    PdfTableCell.TextCell("Jump to target", linkDestinationName: "TargetCell", linkContents: "Table cell jump"),
                    PdfTableCell.TextCell("Target cell", namedDestinationName: "TargetCell")
                }
            })
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        PdfNamedDestination destination = Assert.Single(logical.NamedDestinations);
        Assert.Equal("TargetCell", destination.Name);
        PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByDestinationName("TargetCell"));
        Assert.True(link.IsNamedDestinationLink);
        Assert.Equal("Table cell jump", link.Contents);
        Assert.Equal("TargetCell", link.DestinationName);
    }
}
