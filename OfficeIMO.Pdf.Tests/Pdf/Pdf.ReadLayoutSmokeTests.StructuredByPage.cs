using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class PdfReadLayoutSmokeTests {
    [Fact]
    public void PdfTextExtractor_ExtractStructuredByPage_DetectsGeneratedTableRows() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 320,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Structured table readback marker."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha,Inc", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new System.Collections.Generic.List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        var pages = PdfTextExtractor.ExtractStructuredByPage(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        var page = Assert.Single(pages);
        Assert.Contains(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Structuredtablereadbackmarker", StringComparison.Ordinal));
        Assert.DoesNotContain(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("A-100", StringComparison.Ordinal));

        var table = Assert.Single(page.TablesDetailed, t => t.Rows.Count >= 3 && t.Columns.Count >= 3);
        Assert.Contains(table.Rows, row => row.Length >= 3 &&
            Normalize(row[0]) == "Code" &&
            Normalize(row[1]) == "Name" &&
            Normalize(row[2]) == "Qty");
        Assert.Contains(table.Rows, row => row.Length >= 3 &&
            Normalize(row[0]) == "A-100" &&
            Normalize(row[1]) == "Alpha,Inc" &&
            Normalize(row[2]) == "2");
        Assert.Contains(table.Rows, row => row.Length >= 3 &&
            Normalize(row[0]) == "B-200" &&
            Normalize(row[1]) == "Beta" &&
            Normalize(row[2]) == "14");

        var tablePages = PdfTextExtractor.ExtractTablesByPage(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        var tablePage = Assert.Single(tablePages);
        Assert.Equal(1, tablePage.PageNumber);
        var extractedTable = Assert.Single(tablePage.Tables, t => t.Rows.Count >= 3 && t.Columns.Count >= 3);
        Assert.Contains(extractedTable.Rows, row => row.Length >= 3 &&
            Normalize(row[0]) == "A-100" &&
            Normalize(row[1]) == "Alpha,Inc" &&
            Normalize(row[2]) == "2");

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-table-csv-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputDirectory = Path.Combine(directory, "tables");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, bytes);

            var paths = PdfTextExtractor.ExtractTablesByPage(inputPath, outputDirectory, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });

            string csvPath = Assert.Single(paths);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0001-table-0001.csv"), csvPath);
            string csv = File.ReadAllText(csvPath);
            Assert.Contains("Code,Name,Qty", csv, StringComparison.Ordinal);
            Assert.Contains("A-100,\"Alpha,Inc\",2", csv, StringComparison.Ordinal);

            using var stream = new MemoryStream(bytes);
            string streamOutputDirectory = Path.Combine(directory, "stream-tables");
            var streamPaths = PdfTextExtractor.ExtractTablesByPage(stream, streamOutputDirectory, "stream-source.pdf", new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });

            string streamCsvPath = Assert.Single(streamPaths);
            Assert.Equal(Path.Combine(streamOutputDirectory, "stream-source-page-0001-table-0001.csv"), streamCsvPath);
            Assert.Contains("A-100,\"Alpha,Inc\",2", File.ReadAllText(streamCsvPath), StringComparison.Ordinal);

            string byteOutputDirectory = Path.Combine(directory, "byte-tables");
            var bytePaths = PdfTextExtractor.ExtractTablesByPage(bytes, byteOutputDirectory, "byte-source.pdf", new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });

            string byteCsvPath = Assert.Single(bytePaths);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-source-page-0001-table-0001.csv"), byteCsvPath);
            Assert.Contains("A-100,\"Alpha,Inc\",2", File.ReadAllText(byteCsvPath), StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void PdfTextExtractor_ExtractStructuredByPage_GroupsParagraphLines() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("This structured paragraph should wrap across multiple nearby PDF text lines so wrappers can start from paragraph-like objects."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "P-100", "Paragraph table text", "2" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new System.Collections.Generic.List<double?> { 50, 100, 30 },
                HeaderRowCount = 1
            })
            .ToBytes();

        StructuredPage page = Assert.Single(PdfTextExtractor.ExtractStructuredByPage(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }));

        StructuredParagraph paragraph = Assert.Single(page.Paragraphs, item => item.Text.Contains("structured paragraph", StringComparison.Ordinal));
        Assert.True(paragraph.Lines.Count > 1);
        Assert.Contains("structured paragraph", paragraph.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("P-100", paragraph.Text, StringComparison.Ordinal);

        StructuredParagraphPage paragraphPage = Assert.Single(PdfTextExtractor.ExtractParagraphsByPage(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }));
        StructuredParagraph extractedParagraph = Assert.Single(paragraphPage.Paragraphs, item => item.Text.Contains("structured paragraph", StringComparison.Ordinal));
        Assert.Equal(1, paragraphPage.PageNumber);
        Assert.True(extractedParagraph.Lines.Count > 1);
        Assert.DoesNotContain("P-100", extractedParagraph.Text, StringComparison.Ordinal);

        using var stream = new MemoryStream(bytes);
        StructuredParagraphPage streamParagraphPage = Assert.Single(PdfTextExtractor.ExtractParagraphsByPage(stream, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }));
        Assert.Equal(1, streamParagraphPage.PageNumber);
        Assert.Contains(streamParagraphPage.Paragraphs, item => item.Text.Contains("structured paragraph", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfTextExtractor_ExtractStructuredByPage_DoesNotDuplicateFallbackTableRowsAsParagraphs() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 180,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10,
                DefaultParagraphStyle = new PdfParagraphStyle {
                    DefaultTabStopWidth = 160,
                    SpacingAfter = 0
                }
            })
            .Paragraph(p => p.Text("Fallback table row").Tab().Text("12345"))
            .ToBytes();

        StructuredPage page = Assert.Single(PdfTextExtractor.ExtractStructuredByPage(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }));

        Assert.Contains(page.Tables, row => row.Length >= 2 &&
            Normalize(row[0]).Contains("Fallbacktablerow", StringComparison.Ordinal) &&
            Normalize(row[1]).Contains("12345", StringComparison.Ordinal));
        Assert.DoesNotContain(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Fallbacktablerow", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfTextExtractor_ExtractStructuredByPage_DetectsHeadingLines() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("Structured Heading")
            .Paragraph(p => p.Text("Body copy for heading detection."))
            .ToBytes();

        StructuredPage page = Assert.Single(PdfTextExtractor.ExtractStructuredByPage(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }));

        StructuredHeading heading = Assert.Single(page.Headings);
        Assert.Equal("Structured Heading", heading.Text);
        Assert.Equal(1, heading.Level);
        Assert.True(heading.FontSize > page.Paragraphs[0].Lines[0].FontSize);
        Assert.DoesNotContain(page.Paragraphs, paragraph => paragraph.Text.Contains("Structured Heading", StringComparison.Ordinal));

        StructuredHeadingPage headingPage = Assert.Single(PdfTextExtractor.ExtractHeadingsByPage(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }));
        StructuredHeading extractedHeading = Assert.Single(headingPage.Headings);
        Assert.Equal(1, headingPage.PageNumber);
        Assert.Equal("Structured Heading", extractedHeading.Text);

        using var stream = new MemoryStream(bytes);
        StructuredHeadingPage streamHeadingPage = Assert.Single(PdfTextExtractor.ExtractHeadingsByPage(stream, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }));
        Assert.Equal(1, streamHeadingPage.PageNumber);
        Assert.Contains(streamHeadingPage.Headings, item => item.Text == "Structured Heading");

        var rangeHeadingPages = PdfTextExtractor.ExtractHeadingsByPageRanges(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }, PdfPageRange.ParseMany("1,1"));
        Assert.Equal(2, rangeHeadingPages.Count);
        Assert.All(rangeHeadingPages, item => Assert.Contains(item.Headings, heading => heading.Text == "Structured Heading"));

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-heading-ranges-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "headings.pdf");
        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, bytes);

            StructuredHeadingPage pathHeadingPage = Assert.Single(PdfTextExtractor.ExtractHeadingsByPageRanges(inputPath, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }, PdfPageRange.ParseMany("1")));
            Assert.Contains(pathHeadingPage.Headings, item => item.Text == "Structured Heading");
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void PdfTextExtractor_ExtractStructuredByPage_DetectsListItems() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Bullets(new[] { "First bullet", "Second bullet" })
            .Numbered(new[] { "First numbered", "Second numbered" }, startNumber: 3)
            .ToBytes();

        StructuredPage page = Assert.Single(PdfTextExtractor.ExtractStructuredByPage(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }));

        Assert.Contains(page.ListNodes, item => item.Marker == "3" && item.Text == "First numbered" && item.Level == 1);
        Assert.Contains(page.ListNodes, item => item.Marker.Length > 0 && item.Text == "First bullet" && item.Level == 1);

        StructuredListItemPage listPage = Assert.Single(PdfTextExtractor.ExtractListItemsByPage(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }));
        Assert.Equal(1, listPage.PageNumber);
        Assert.Contains(listPage.ListItems, item => item.Text == "Second bullet");
        Assert.Contains(listPage.ListItems, item => item.Marker == "4" && item.Text == "Second numbered");

        using var stream = new MemoryStream(bytes);
        StructuredListItemPage streamListPage = Assert.Single(PdfTextExtractor.ExtractListItemsByPage(stream, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }));
        Assert.Contains(streamListPage.ListItems, item => item.Text == "First bullet");

        var rangeListPages = PdfTextExtractor.ExtractListItemsByPageRanges(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }, PdfPageRange.ParseMany("1,1"));
        Assert.Equal(2, rangeListPages.Count);
        Assert.All(rangeListPages, item => Assert.Contains(item.ListItems, listItem => listItem.Text == "First numbered"));

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-list-ranges-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "lists.pdf");
        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, bytes);

            StructuredListItemPage pathListPage = Assert.Single(PdfTextExtractor.ExtractListItemsByPageRanges(inputPath, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }, PdfPageRange.ParseMany("1")));
            Assert.Contains(pathListPage.ListItems, item => item.Marker == "3" && item.Text == "First numbered");
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}
