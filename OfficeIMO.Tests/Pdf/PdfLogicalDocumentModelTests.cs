using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfLogicalDocumentTests {
    [Fact]
    public void Load_BuildsLogicalPagesWithTextTablesAndImages() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "Logical sample", author: "OfficeIMO")
            .H1("Logical Heading")
            .Paragraph(p => p.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .Image(CreateMinimalRgbPng(), 18, 18)
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        PdfLogicalPage page = Assert.Single(logical.Pages);
        Assert.Equal("Logical sample", logical.Metadata.Title);
        Assert.True(logical.HasSourcePage(1));
        Assert.Same(page, Assert.Single(logical.PagesBySourcePageNumber[1]));
        Assert.Same(page, Assert.Single(logical.GetPages(1)));
        Assert.Empty(logical.GetPages(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.HasSourcePage(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.GetPages(0));
        PdfLogicalHeading heading = Assert.Single(page.Headings);
        Assert.Equal("Logical Heading", heading.Text);
        Assert.Equal(1, heading.Level);
        Assert.Equal(PdfLogicalElementKind.Heading, heading.Line.Kind);
        Assert.Same(heading, Assert.Single(logical.Headings));
        Assert.Contains(page.TextBlocks, block => Normalize(block.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(logical.TextBlocks, block => Normalize(block.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(page.TextBlocks, block =>
            block.Kind == PdfLogicalElementKind.ListItem &&
            Normalize(block.Text).Contains("Detectedlogicalbullet", StringComparison.Ordinal));
        PdfLogicalListItem listItem = Assert.Single(page.ListItems);
        Assert.Equal("Detected logical bullet", listItem.Text);
        Assert.Equal(1, listItem.Level);
        Assert.NotEmpty(listItem.Marker);
        Assert.Equal(PdfLogicalElementKind.ListItem, listItem.Line.Kind);
        Assert.Same(listItem, Assert.Single(logical.ListItems));
        Assert.Contains(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(logical.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.DoesNotContain(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("A-100", StringComparison.Ordinal));

        PdfLogicalTable table = Assert.Single(page.Tables, item => item.Rows.Count >= 3 && item.Columns.Count >= 3);
        Assert.Same(table, Assert.Single(logical.Tables, item => item.Rows.Count >= 3 && item.Columns.Count >= 3));
        Assert.Contains(table.Rows, row => row.Count >= 3 &&
            Normalize(row[0]) == "A-100" &&
            Normalize(row[1]) == "Alpha" &&
            Normalize(row[2]) == "2");
        Assert.Contains(table.Cells, cell =>
            cell.PageNumber == 1 &&
            cell.RowIndex == 1 &&
            cell.ColumnIndex == 0 &&
            Normalize(cell.Text) == "A-100" &&
            cell.Column is not null &&
            cell.Column.From < cell.Column.To);
        Assert.Contains(table.Cells, cell =>
            cell.RowIndex == 2 &&
            cell.ColumnIndex == 2 &&
            Normalize(cell.Text) == "14");

        PdfLogicalImage image = Assert.Single(page.Images);
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal("image/png", image.MimeType);
        PdfImagePlacement placement = Assert.Single(image.Placements);
        Assert.True(image.HasPlacements);
        Assert.Equal(1, placement.PageNumber);
        Assert.Equal(image.ResourceName, placement.ResourceName);
        Assert.True(placement.Width > 0);
        Assert.True(placement.Height > 0);
        Assert.True(placement.IsAxisAligned);
        Assert.Same(image, Assert.Single(logical.Images));

        Assert.True(logical.HasElementKind(PdfLogicalElementKind.Table));
        Assert.True(logical.HasElementKind(PdfLogicalElementKind.Image));
        Assert.True(page.HasElementKind(PdfLogicalElementKind.Heading));
        Assert.True(page.HasElementKind(PdfLogicalElementKind.Image));
        Assert.Same(heading.Line, Assert.Single(page.GetElements(PdfLogicalElementKind.Heading)));
        Assert.Same(table, Assert.Single(logical.GetElements(PdfLogicalElementKind.Table)));
        Assert.Same(image, Assert.Single(logical.ElementsByKind[PdfLogicalElementKind.Image]));
        Assert.Equal(page.Elements, logical.ElementsByPageNumber[1]);
        Assert.Equal(page.Elements, logical.GetElements(1));
        Assert.Empty(logical.GetElements(PdfLogicalElementKind.LinkAnnotation));
        Assert.Empty(page.GetElements(PdfLogicalElementKind.LinkAnnotation));
        Assert.Empty(logical.GetElements(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.GetElements(0));
        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.Table);
        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.Image);
    }
}
