using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfLogicalDocumentTests {
    [Fact]
    public void ToMarkdown_RendersLogicalHeadingsParagraphsListsTablesAndImages() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
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

        string markdown = logical.ToMarkdown();
        string normalizedMarkdown = Normalize(markdown);

        Assert.Contains("# Logical Heading", markdown, StringComparison.Ordinal);
        Assert.Contains("Logicalreadbackmarker.", normalizedMarkdown, StringComparison.Ordinal);
        Assert.Contains("-Detectedlogicalbullet", normalizedMarkdown, StringComparison.Ordinal);
        Assert.Contains("| Code | Name | Qty |", markdown, StringComparison.Ordinal);
        Assert.Contains("| --- | --- | ---: |", markdown, StringComparison.Ordinal);
        Assert.Contains("| A-100 | Alpha | 2 |", markdown, StringComparison.Ordinal);
        Assert.Contains("[Image: page 1", markdown, StringComparison.Ordinal);
        AssertContainsInOrder(normalizedMarkdown,
            "#LogicalHeading",
            "Logicalreadbackmarker.",
            "-Detectedlogicalbullet",
            "|Code|Name|Qty|",
            "[Image:page1");

        string withoutImages = logical.ToMarkdown(new PdfLogicalMarkdownOptions {
            IncludeImagePlaceholders = false
        });
        Assert.DoesNotContain("[Image:", withoutImages, StringComparison.Ordinal);
    }

    [Fact]
    public void ToMarkdown_RightAlignsNumericTableColumns() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Item", "Qty", "Amount" },
                new[] { "Service", "2", "$125.50" },
                new[] { "Discount", "1", "(10.00)" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 130, 60, 90 },
                HeaderRowCount = 1
            })
            .ToBytes();

        string markdown = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).ToMarkdown();

        Assert.Contains("| Item | Qty | Amount |", markdown, StringComparison.Ordinal);
        Assert.Contains("| --- | ---: | ---: |", markdown, StringComparison.Ordinal);
        Assert.Contains("| Service | 2 | $125.50 |", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void TableAnalysis_ExposesColumnProfilesForAdapters() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Qty", "Mixed" },
                new[] { "A-100", "2", "123" },
                new[] { "B-200", "14", "n/a" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 100, 60, 90 },
                HeaderRowCount = 1
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).Pages[0].Tables);

        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Equal(3, data.ColumnProfiles.Count);
        Assert.Equal(PdfLogicalTableColumnKind.Text, data.ColumnProfiles[0].Kind);
        Assert.Equal(PdfLogicalTableColumnKind.Numeric, data.ColumnProfiles[1].Kind);
        Assert.Equal(PdfLogicalTableColumnKind.Mixed, data.ColumnProfiles[2].Kind);
        Assert.False(data.IsNumericColumn(0));
        Assert.True(data.IsNumericColumn(1));
        Assert.False(data.IsNumericColumn(2));
        Assert.Equal(2, data.ColumnProfiles[1].NonEmptyCellCount);
        Assert.Equal(2, data.ColumnProfiles[1].NumericCellCount);
        Assert.Equal(0.5d, data.ColumnProfiles[2].Confidence);
    }

    [Fact]
    public void TableAnalysis_PreservesOrdinaryTwoColumnTableHeaders() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 320,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Name", "Age" },
                new[] { "Alice", "42" },
                new[] { "Bob", "37" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 120, 80 },
                HeaderRowCount = 1
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).Pages[0].Tables);

        PdfLogicalTableStructure structure = PdfLogicalTableAnalysis.Analyze(table);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);

        Assert.True(structure.HasHeaderRow);
        Assert.False(structure.IsKeyValueTable);
        Assert.Equal(new[] { "Name", "Age" }, structure.Columns);
        Assert.Equal(1, structure.BodyStartRowIndex);
        Assert.Equal(new[] { "Name", "Age" }, data.Columns);
        Assert.Equal(new[] { "Alice", "42" }, data.Rows[0]);
        Assert.True(data.IsNumericColumn(1));
    }

    [Fact]
    public void TableAnalysis_IdentifiesHeaderlessKeyValueTableShape() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .KeyValueTable(new[] {
                PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfKeyValueRow.Text("Due", "2026-06-30")
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 120, 170 },
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        PdfLogicalTable table = Assert.Single(logical.Pages[0].Tables);
        PdfLogicalTableStructure structure = PdfLogicalTableAnalysis.Analyze(table);

        Assert.Equal(2, structure.ColumnCount);
        Assert.Equal(new[] { "Key", "Value" }, structure.Columns);
        Assert.Equal(0, structure.BodyStartRowIndex);
        Assert.Equal(3, structure.TotalBodyRowCount);
        Assert.False(structure.HasHeaderRow);
        Assert.True(structure.IsKeyValueTable);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table, maxRows: 2);
        Assert.Equal(new[] { "Key", "Value" }, data.Columns);
        Assert.Equal(2, data.Rows.Count);
        Assert.Equal(3, data.TotalRowCount);
        Assert.True(data.Truncated);
        Assert.Equal(new[] { "InvoiceId", "INV-001" }, data.Rows[0]);
        PdfLogicalTableExtraction extraction = Assert.Single(PdfLogicalTableAnalysis.ExtractTables(logical, maxRows: 2));
        Assert.Equal(0, extraction.PageIndex);
        Assert.Equal(1, extraction.PageNumber);
        Assert.Equal(0, extraction.TableIndex);
        Assert.Equal(table.DetectionKind, extraction.DetectionKind);
        Assert.True(extraction.Data.Truncated);

        string markdown = logical.ToMarkdown();
        Assert.Contains("| Key | Value |", markdown, StringComparison.Ordinal);
        Assert.Contains("| InvoiceId | INV-001 |", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void ToMarkdown_EscapesMarkdownControlSyntaxFromPdfText() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("# Literal heading marker"))
            .Paragraph(p => p.Text("[not a link](https://example.test)"))
            .ToBytes();

        string markdown = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).ToMarkdown();

        string normalized = Normalize(markdown);
        Assert.Contains("\\#Literalheadingmarker", normalized, StringComparison.Ordinal);
        Assert.Contains("\\[notalink\\](https://example.test)", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void ToMarkdown_DoesNotRenderLeaderRowsTwiceWhenTableAlreadyContainsThem() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Chapter One ........ 3"))
            .ToBytes();

        string markdown = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).ToMarkdown();

        Assert.Equal(1, CountOccurrences(markdown, "Chapter One"));
    }

    [Fact]
    public void ToMarkdown_RendersDirectDestinationLinkAnnotations() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildDirectDestinationLinkPdf());

        string markdown = logical.ToMarkdown(new PdfLogicalMarkdownOptions {
            IncludeLinkAnnotations = true
        });

        Assert.Contains("[Link: Direct destination link -> page 1, FitRectangle, left 10, bottom 20, right 90, top 144]", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void ToMarkdown_RendersNamedActionLinkAnnotations() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildNamedActionLinkPdf());

        string markdown = logical.ToMarkdown(new PdfLogicalMarkdownOptions {
            IncludeLinkAnnotations = true
        });

        Assert.Contains("[Link: Next page action -> named action NextPage]", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void ToMarkdown_RendersRemoteGoToLinkAnnotations() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildRemoteGoToLinkPdf());

        string markdown = logical.ToMarkdown(new PdfLogicalMarkdownOptions {
            IncludeLinkAnnotations = true
        });

        Assert.Contains("[Link: Remote report link -> remote file remote-report.pdf, page 2, FitHorizontal, top 144]", markdown, StringComparison.Ordinal);
    }
}
