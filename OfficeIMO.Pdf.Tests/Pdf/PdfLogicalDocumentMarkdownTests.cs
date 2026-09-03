using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentReadResultTests {
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

        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf, new PdfTextLayoutOptions {
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

        string markdown = PdfDocumentReadResult.Load(pdf, new PdfTextLayoutOptions {
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

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).Pages[0].Tables);

        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Assert.Equal(3, data.ColumnProfiles.Count);
        Assert.Equal(new[] { "Code", "Qty", "Mixed" }, data.Columns);
        Assert.Equal(PdfLogicalTableColumnKind.Text, data.ColumnProfiles[0].Kind);
        Assert.Equal(PdfLogicalTableColumnKind.Numeric, data.ColumnProfiles[1].Kind);
        Assert.Equal(PdfLogicalTableColumnKind.Mixed, data.ColumnProfiles[2].Kind);
        Assert.Equal("Qty", data.ColumnProfiles[1].Name);
        Assert.False(data.IsNumericColumn(0));
        Assert.True(data.IsNumericColumn(1));
        Assert.False(data.IsNumericColumn(2));
        Assert.Equal(2, data.ColumnProfiles[1].NonEmptyCellCount);
        Assert.Equal(2, data.ColumnProfiles[1].NumericCellCount);
        Assert.Equal(0.5d, data.ColumnProfiles[2].Confidence);
    }

    [Fact]
    public void TableAnalysis_RecoversGeneratedHeaderBandForColumnProfiles() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 460,
                PageHeight = 380,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("Revenue Readback Diagnostics")
            .Paragraph(paragraph => paragraph.Text("Image geometry and table confidence marker."))
            .Table(new[] {
                new[] { "Metric", "Score", "Owner" },
                new[] { "Renewal quality", "97", "Finance" },
                new[] { "Pipeline coverage", "84", "Sales" },
                new[] { "Risk burn-down", "76", "Operations" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 150, 70, 110 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).Pages[0].Tables);

        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        PdfLogicalTableColumnProfile scoreProfile = Assert.Single(data.ColumnProfiles, profile => profile.Name == "Score");

        Assert.Equal(new[] { "Metric", "Score", "Owner" }, data.Columns);
        Assert.Equal(3, data.Rows.Count);
        Assert.Equal(new[] { "Renewal quality", "97", "Finance" }, data.Rows[0]);
        Assert.True(data.Diagnostics.HasGeometry);
        Assert.True(data.Diagnostics.Width > 0);
        Assert.True(data.Diagnostics.Height > 0);
        Assert.True(data.Diagnostics.Confidence >= 0.95D);
        Assert.NotEmpty(data.Diagnostics.Evidence);
        Assert.Equal(0.95D, data.Diagnostics.SchemaConfidence, 3);
        Assert.Equal(1D, data.Diagnostics.CellCompleteness, 3);
        Assert.Equal(1D, data.Diagnostics.ColumnGeometryConfidence, 3);
        Assert.Equal(PdfLogicalTableColumnKind.Numeric, scoreProfile.Kind);
        Assert.Equal(3, scoreProfile.NumericCellCount);
        Assert.Equal(1D, scoreProfile.Confidence, 3);
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

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).Pages[0].Tables);

        PdfLogicalTableStructure structure = PdfLogicalTableAnalysis.Analyze(table);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);

        Assert.True(structure.HasHeaderRow);
        Assert.Equal(PdfLogicalTableSchemaKind.HeaderRow, structure.SchemaKind);
        Assert.InRange(structure.SchemaConfidence, 0.8D, 1D);
        Assert.NotEmpty(structure.SchemaEvidence);
        Assert.Equal(new[] { "Name", "Age" }, structure.Columns);
        Assert.Equal(1, structure.BodyStartRowIndex);
        Assert.Equal(new[] { "Name", "Age" }, data.Columns);
        Assert.Equal(new[] { "Alice", "42" }, data.Rows[0]);
        Assert.True(data.IsNumericColumn(1));
    }

    [Fact]
    public void TableAnalysis_KeepsMixedNumericFirstRowsAsBodyRows() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 460,
                PageHeight = 320,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Product A", "10", "20" },
                new[] { "Product B", "11", "21" },
                new[] { "Product C", "12", "22" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 150, 70, 70 },
                HeaderRowCount = 0
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).Pages[0].Tables);

        PdfLogicalTableStructure structure = PdfLogicalTableAnalysis.Analyze(table);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);

        Assert.False(structure.HasHeaderRow);
        Assert.Equal(0, structure.BodyStartRowIndex);
        Assert.Equal(new[] { "", "", "" }, data.Columns);
        Assert.Equal(new[] { "Product A", "10", "20" }, data.Rows[0]);
    }

    [Fact]
    public void TableAnalysis_DoesNotPromoteAnAllTextDataRowWithoutStructuralHeaderEvidence() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 460,
                PageHeight = 320,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Sverige", "Stockholm" },
                new[] { "Norge", "Oslo" },
                new[] { "Suomi", "Helsinki" }
            }, style: new PdfTableStyle {
                HeaderRowCount = 0,
                ColumnWidthPoints = new List<double?> { 150, 150 }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);

        Assert.False(data.Structure.HasHeaderRow);
        Assert.Equal(0, data.Structure.BodyStartRowIndex);
        Assert.Equal(new[] { "", "" }, data.Columns);
        Assert.Equal(new[] { "Sverige", "Stockholm" }, data.Rows[0]);
    }

    [Fact]
    public void TableAnalysis_PreservesStructurallyEstablishedNumericDuplicateAndBlankHeaders() {
        PdfLogicalTable table = PdfLogicalTable.From(
            1,
            new PdfUnderstandingTableCandidate(
                "test-geometry",
                100D,
                20D,
                new[] {
                    new PdfUnderstandingTableColumn(0D, 100D),
                    new PdfUnderstandingTableColumn(100D, 200D),
                    new PdfUnderstandingTableColumn(200D, 300D)
                },
                new IReadOnlyList<string>[] {
                    new[] { "1", "1", "" },
                    new[] { "A", "B", "C" },
                    new[] { "D", "E", "F" }
                },
                Array.Empty<PdfUnderstandingLine>(),
                evidence: new[] {
                    new PdfInferenceEvidence(
                        "table.header-emphasis",
                        "The first row has distinct source typography.",
                        0.9D)
                }));
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);

        Assert.True(data.Structure.HasHeaderRow);
        Assert.Equal(PdfLogicalTableSchemaKind.HeaderRow, data.Structure.SchemaKind);
        Assert.Equal(new[] { "1", "1", "" }, data.Columns);
        Assert.Equal(new[] { "A", "B", "C" }, data.Rows[0]);
        Assert.Contains(data.Structure.SchemaEvidence, static evidence => evidence.Code == "table.header-emphasis");
    }

    [Fact]
    public void TableAnalysis_UsesDistinctHeaderTypographyWithoutHeaderVocabulary() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 460,
                PageHeight = 320,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Område", "Ansvarig" },
                new[] { "Norr", "Linnea" },
                new[] { "Söder", "Mikael" }
            }, style: new PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 150, 150 }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfDocumentReadResult.Load(pdf).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);

        Assert.True(data.Structure.HasHeaderRow);
        Assert.Equal(new[] { "Område", "Ansvarig" }, data.Columns);
        Assert.Equal(new[] { "Norr", "Linnea" }, data.Rows[0]);
    }

    [Fact]
    public void TableAnalysis_ReportsHeaderlessTwoColumnShapeAsUnknown() {
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

        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        PdfLogicalTable table = Assert.Single(logical.Pages[0].Tables);
        PdfLogicalTableStructure structure = PdfLogicalTableAnalysis.Analyze(table);

        Assert.Equal(2, structure.ColumnCount);
        Assert.Equal(new[] { "", "" }, structure.Columns);
        Assert.Equal(0, structure.BodyStartRowIndex);
        Assert.Equal(3, structure.TotalBodyRowCount);
        Assert.False(structure.HasHeaderRow);
        Assert.Equal(PdfLogicalTableSchemaKind.Unknown, structure.SchemaKind);
        Assert.Equal(0D, structure.SchemaConfidence);
        Assert.Contains(structure.SchemaEvidence, static evidence => evidence.Code == "table.schema-unknown");
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table, maxRows: 2);
        Assert.Equal(new[] { "", "" }, data.Columns);
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
        Assert.Contains("|  |  |", markdown, StringComparison.Ordinal);
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

        string markdown = PdfDocumentReadResult.Load(pdf, new PdfTextLayoutOptions {
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

        string markdown = PdfDocumentReadResult.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).ToMarkdown();

        Assert.Equal(1, CountOccurrences(markdown, "Chapter One"));
    }

    [Fact]
    public void ToMarkdown_RendersDirectDestinationLinkAnnotations() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(BuildDirectDestinationLinkPdf());

        string markdown = logical.ToMarkdown(new PdfLogicalMarkdownOptions {
            IncludeLinkAnnotations = true
        });

        Assert.Contains("[Link: Direct destination link -> page 1, FitRectangle, left 10, bottom 20, right 90, top 144]", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void ToMarkdown_RendersNamedActionLinkAnnotations() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(BuildNamedActionLinkPdf());

        string markdown = logical.ToMarkdown(new PdfLogicalMarkdownOptions {
            IncludeLinkAnnotations = true
        });

        Assert.Contains("[Link: Next page action -> named action NextPage]", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void ToMarkdown_RendersRemoteGoToLinkAnnotations() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(BuildRemoteGoToLinkPdf());

        string markdown = logical.ToMarkdown(new PdfLogicalMarkdownOptions {
            IncludeLinkAnnotations = true
        });

        Assert.Contains("[Link: Remote report link -> remote file remote-report.pdf, page 2, FitHorizontal, top 144]", markdown, StringComparison.Ordinal);
    }
}
