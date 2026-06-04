using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class PdfReadLayoutSmokeTests {
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
        byte[] bytes = PdfDocumentRasterVisualBaselineTests.CreateLineItemsTwoPage();
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
}
