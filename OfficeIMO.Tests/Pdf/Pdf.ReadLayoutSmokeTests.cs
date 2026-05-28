using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfReadLayoutSmokeTests {
    [Fact]
    public void PdfReadDocument_ColumnAndStructuredApis_DoNotThrow() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        try {
            var pdf = PdfDoc.Create()
                .Meta(title: "Smoke")
                .H1("Header")
                .Paragraph(p => p.Text("Line one for extraction."))
                .Paragraph(p => p.Text("Line two for extraction."));

            pdf.Save(path);

            var doc = PdfReadDocument.Load(path);
            Assert.NotNull(doc);
            Assert.NotEmpty(doc.Pages);

            var text = doc.ExtractTextWithColumns();
            Assert.NotNull(text);

            var structured = doc.ExtractStructured();
            Assert.NotNull(structured.Lines);
            Assert.NotNull(structured.Toc);
            Assert.NotNull(structured.Lists);
            Assert.NotNull(structured.LeaderRows);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PdfReadPage_ExtensionApis_DoNotThrow() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        try {
            var pdf = PdfDoc.Create()
                .Paragraph(p => p.Text("Page extension api smoke."));
            pdf.Save(path);

            var doc = PdfReadDocument.Load(path);
            Assert.NotEmpty(doc.Pages);

            var page = doc.Pages[0];
            var text = page.ExtractTextWithColumns(new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });
            Assert.NotNull(text);

            var structured = page.ExtractStructured(new PdfTextLayoutOptions());
            Assert.NotNull(structured);
            Assert.NotNull(structured.Lines);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PdfTextExtractor_ExtractStructuredByPage_DetectsGeneratedTableRows() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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

    [Fact]
    public void PdfTextExtractor_ExtractStructuredAndTablesByPageRanges_UsesSelectedSourcePages() {
        byte[] bytes = BuildThreePageTablePdf();

        var structuredPages = PdfTextExtractor.ExtractStructuredByPageRanges(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }, PdfPageRange.ParseMany("3,1-2,3"));

        Assert.Equal(4, structuredPages.Count);
        Assert.Contains(structuredPages[0].Lines, line => Normalize(line).Contains("Thirdpagetable", StringComparison.Ordinal));
        Assert.Contains(structuredPages[1].Lines, line => Normalize(line).Contains("Firstpagetable", StringComparison.Ordinal));
        Assert.Contains(structuredPages[2].Lines, line => Normalize(line).Contains("Secondpagemarker", StringComparison.Ordinal));
        Assert.Contains(structuredPages[3].Lines, line => Normalize(line).Contains("Thirdpagetable", StringComparison.Ordinal));

        var headingPages = PdfTextExtractor.ExtractHeadingsByPageRanges(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }, PdfPageRange.ParseMany("3,1-2,3"));

        Assert.Equal(4, headingPages.Count);
        Assert.Equal(3, headingPages[0].PageNumber);
        Assert.Equal(1, headingPages[1].PageNumber);
        Assert.Equal(2, headingPages[2].PageNumber);
        Assert.Equal(3, headingPages[3].PageNumber);
        Assert.Empty(headingPages[0].Headings);

        var paragraphPages = PdfTextExtractor.ExtractParagraphsByPageRanges(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }, PdfPageRange.ParseMany("3,1-2,3"));

        Assert.Equal(4, paragraphPages.Count);
        Assert.Equal(3, paragraphPages[0].PageNumber);
        Assert.Equal(1, paragraphPages[1].PageNumber);
        Assert.Equal(2, paragraphPages[2].PageNumber);
        Assert.Equal(3, paragraphPages[3].PageNumber);
        Assert.Contains(paragraphPages[0].Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Thirdpagetable", StringComparison.Ordinal));
        Assert.Contains(paragraphPages[1].Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Firstpagetable", StringComparison.Ordinal));
        Assert.Contains(paragraphPages[2].Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Secondpagemarker", StringComparison.Ordinal));

        var tablePages = PdfTextExtractor.ExtractTablesByPageRanges(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }, PdfPageRange.ParseMany("3,1-2,3"));

        Assert.Equal(4, tablePages.Count);
        Assert.Equal(3, tablePages[0].PageNumber);
        Assert.Equal(1, tablePages[1].PageNumber);
        Assert.Equal(2, tablePages[2].PageNumber);
        Assert.Equal(3, tablePages[3].PageNumber);
        Assert.Contains(tablePages[0].Tables.SelectMany(table => table.Rows), row => row.Length >= 3 && Normalize(row[0]) == "C-300");
        Assert.Contains(tablePages[1].Tables.SelectMany(table => table.Rows), row => row.Length >= 3 && Normalize(row[0]) == "A-100");
        Assert.DoesNotContain(tablePages[2].Tables.SelectMany(table => table.Rows), row => row.Length >= 3 && Normalize(row[0]) == "A-100");
        Assert.Contains(tablePages[3].Tables.SelectMany(table => table.Rows), row => row.Length >= 3 && Normalize(row[0]) == "C-300");

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-table-page-ranges-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputDirectory = Path.Combine(directory, "tables");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, bytes);

            var pathHeadingPages = PdfTextExtractor.ExtractHeadingsByPageRanges(inputPath, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }, PdfPageRange.ParseMany("1"));

            StructuredHeadingPage pathHeadingPage = Assert.Single(pathHeadingPages);
            Assert.Equal(1, pathHeadingPage.PageNumber);
            Assert.Empty(pathHeadingPage.Headings);

            using var headingStream = new MemoryStream(bytes);
            var streamHeadingPages = PdfTextExtractor.ExtractHeadingsByPageRanges(headingStream, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }, PdfPageRange.ParseMany("2"));

            StructuredHeadingPage streamHeadingPage = Assert.Single(streamHeadingPages);
            Assert.Equal(2, streamHeadingPage.PageNumber);
            Assert.Empty(streamHeadingPage.Headings);

            var pathParagraphPages = PdfTextExtractor.ExtractParagraphsByPageRanges(inputPath, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }, PdfPageRange.ParseMany("1"));

            StructuredParagraphPage pathParagraphPage = Assert.Single(pathParagraphPages);
            Assert.Equal(1, pathParagraphPage.PageNumber);
            Assert.Contains(pathParagraphPage.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Firstpagetable", StringComparison.Ordinal));

            using var paragraphStream = new MemoryStream(bytes);
            var streamParagraphPages = PdfTextExtractor.ExtractParagraphsByPageRanges(paragraphStream, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }, PdfPageRange.ParseMany("2"));

            StructuredParagraphPage streamParagraphPage = Assert.Single(streamParagraphPages);
            Assert.Equal(2, streamParagraphPage.PageNumber);
            Assert.Contains(streamParagraphPage.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Secondpagemarker", StringComparison.Ordinal));

            var paths = PdfTextExtractor.ExtractTablesByPageRanges(inputPath, outputDirectory, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }, PdfPageRange.ParseMany("3,1-2,3"));

            Assert.Equal(3, paths.Count);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0003-table-0001.csv"), paths[0]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0001-table-0001.csv"), paths[1]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0003-occurrence-0002-table-0001.csv"), paths[2]);
            Assert.NotEqual(paths[0], paths[2]);
            Assert.Contains("C-300,Gamma,5", File.ReadAllText(paths[0]), StringComparison.Ordinal);
            Assert.Contains("A-100,Alpha,2", File.ReadAllText(paths[1]), StringComparison.Ordinal);
            Assert.Contains("C-300,Gamma,5", File.ReadAllText(paths[2]), StringComparison.Ordinal);

            using var stream = new MemoryStream(bytes);
            string streamOutputDirectory = Path.Combine(directory, "stream-tables");
            var streamPaths = PdfTextExtractor.ExtractTablesByPageRanges(stream, streamOutputDirectory, "stream-source.pdf", new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }, PdfPageRange.ParseMany("1"));

            string streamCsvPath = Assert.Single(streamPaths);
            Assert.Equal(Path.Combine(streamOutputDirectory, "stream-source-page-0001-table-0001.csv"), streamCsvPath);
            Assert.Contains("A-100,Alpha,2", File.ReadAllText(streamCsvPath), StringComparison.Ordinal);

            string byteOutputDirectory = Path.Combine(directory, "byte-tables");
            var bytePaths = PdfTextExtractor.ExtractTablesByPageRanges(bytes, byteOutputDirectory, "byte-source.pdf", new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }, PdfPageRange.ParseMany("3"));

            string byteCsvPath = Assert.Single(bytePaths);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-source-page-0003-table-0001.csv"), byteCsvPath);
            Assert.Contains("C-300,Gamma,5", File.ReadAllText(byteCsvPath), StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void PdfTextExtractor_ExtractTablesByPageRanges_WritesStatementFixtureTablesForWrappers() {
        byte[] bytes = PdfDocRasterVisualBaselineTests.CreateLineItemsTwoPage();
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-statement-table-csv-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "statement.pdf");
        string outputDirectory = Path.Combine(directory, "tables");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, bytes);

            var tablePages = PdfTextExtractor.ExtractTablesByPageRanges(bytes, PdfPageRange.ParseMany("2,1"));

            Assert.Equal(new[] { 2, 1 }, tablePages.Select(page => page.PageNumber).ToArray());
            Assert.Contains(tablePages[0].Tables.SelectMany(table => table.Rows), row => RowContains(row, "Subtotal", "5201,32PLN"));
            Assert.Contains(tablePages[0].Tables.SelectMany(table => table.Rows), row => RowContains(row, "Total", "6397,62PLN"));
            Assert.Contains(tablePages[1].Tables.SelectMany(table => table.Rows), row => RowContains(row, "Experientiamnostrum", "31,80PLN", "2", "63,60PLN"));

            var paths = PdfTextExtractor.ExtractTablesByPageRanges(inputPath, outputDirectory, PdfPageRange.ParseMany("2,1"));

            Assert.NotEmpty(paths);
            Assert.All(paths, path => Assert.True(File.Exists(path), "Expected extracted statement CSV file to exist: " + path));
            Assert.Contains("-page-0002-table-", Path.GetFileName(paths[0]), StringComparison.Ordinal);
            Assert.Contains(paths, path => Path.GetFileName(path).Contains("-page-0001-table-", StringComparison.Ordinal));

            string combinedCsv = string.Join("\n", paths.Select(File.ReadAllText));
            string normalizedCsv = NormalizeCsvText(combinedCsv);
            Assert.Contains("Subtotal,5201,32PLN", normalizedCsv, StringComparison.Ordinal);
            Assert.Contains("Total,6397,62PLN", normalizedCsv, StringComparison.Ordinal);
            Assert.Contains("Experientiamnostrum,31,80PLN,2,63,60PLN", normalizedCsv, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void PdfTextExtractor_ExtractStructuredAndTablesByPageRanges_RejectsInvalidInputs() {
        byte[] bytes = BuildThreePageTablePdf();

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractStructuredByPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractStructuredByPageRanges(bytes, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractStructuredByPageRanges(bytes));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractStructuredByPageRanges(bytes, default(PdfPageRange)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractStructuredByPageRanges(bytes, PdfPageRange.From(4, 4)));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractListItemsByPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractListItemsByPageRanges(bytes, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractListItemsByPageRanges(bytes));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractListItemsByPageRanges(bytes, default(PdfPageRange)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractListItemsByPageRanges(bytes, PdfPageRange.From(4, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractListItemsByPageRanges((string)null!, PdfPageRange.From(1, 1)));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractHeadingsByPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractHeadingsByPageRanges(bytes, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractHeadingsByPageRanges(bytes));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractHeadingsByPageRanges(bytes, default(PdfPageRange)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractHeadingsByPageRanges(bytes, PdfPageRange.From(4, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractHeadingsByPageRanges((string)null!, PdfPageRange.From(1, 1)));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractParagraphsByPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractParagraphsByPageRanges(bytes, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractParagraphsByPageRanges(bytes));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractParagraphsByPageRanges(bytes, default(PdfPageRange)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractParagraphsByPageRanges(bytes, PdfPageRange.From(4, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractParagraphsByPageRanges((string)null!, PdfPageRange.From(1, 1)));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTablesByPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTablesByPageRanges(bytes, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractTablesByPageRanges(bytes));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractTablesByPageRanges(bytes, default(PdfPageRange)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractTablesByPageRanges(bytes, PdfPageRange.From(4, 4)));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTablesByPageRanges((string)null!, "out", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTablesByPageRanges("input.pdf", (string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractTablesByPageRanges("input.pdf", " ", PdfPageRange.From(1, 1)));
    }

    private static byte[] BuildThreePageTablePdf() {
        return PdfDoc.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 320,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("First page table."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "A-200", "Atlas", "4" }
            }, style: TableStyle())
            .PageBreak()
            .Paragraph(p => p.Text("Second page marker."))
            .PageBreak()
            .Paragraph(p => p.Text("Third page table."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "C-300", "Gamma", "5" },
                new[] { "C-400", "Comet", "7" }
            }, style: TableStyle())
            .ToBytes();
    }

    private static PdfTableStyle TableStyle() {
        return new PdfTableStyle {
            ColumnWidthPoints = new System.Collections.Generic.List<double?> { 70, 170, 60 },
            HeaderRowCount = 1,
            CellPaddingX = 6,
            CellPaddingY = 4
        };
    }

    private static string Normalize(string text) {
        return text.Replace(" ", string.Empty);
    }

    private static bool RowContains(string[] row, params string[] expectedTokens) {
        string rowText = NormalizeCsvText(string.Join(",", row));
        return expectedTokens.All(token => rowText.Contains(token, StringComparison.Ordinal));
    }

    private static string NormalizeCsvText(string text) {
        return new string(text.Where(ch => !char.IsWhiteSpace(ch) && ch != '"').ToArray());
    }
}
