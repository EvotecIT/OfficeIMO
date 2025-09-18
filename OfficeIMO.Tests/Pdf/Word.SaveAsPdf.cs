using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfSample.docx");
        var pdfPath = Path.Combine(_directoryWithFiles, "PdfSample.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            var defaultHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            var defaultFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            defaultHeader.AddParagraph("Sample Header");
            WordTable headerTable = defaultHeader.AddTable(1, 1);
            headerTable.Rows[0].Cells[0].Paragraphs[0].Text = "H1";
            defaultFooter.AddParagraph("Sample Footer");
            WordTable footerTable = defaultFooter.AddTable(1, 1);
            footerTable.Rows[0].Cells[0].Paragraphs[0].Text = "F1";

            WordParagraph heading = document.AddParagraph("Heading One");
            heading.Style = WordParagraphStyles.Heading1;

            WordParagraph formatted = document.AddParagraph("Centered Bold Italic Underlined");
            formatted.Bold = true;
            formatted.Italic = true;
            formatted.Underline = UnderlineValues.Single;
            formatted.ParagraphAlignment = JustificationValues.Center;

            WordList list = document.AddList(WordListStyle.ArticleSections);
            list.AddItem("Numbered Item 1");
            list.AddItem("Numbered Item 2");

            WordTable table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "A2";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";
            WordTable nested = table.Rows[0].Cells[0].AddTable(1, 1);
            nested.Rows[0].Cells[0].Paragraphs[0].Text = "N1";

            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            document.AddParagraph().AddImage(imagePath, 50, 50);

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_HeaderFooterImagesScaling() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfHeaderFooterImages.docx");
        var pdfPath = Path.Combine(_directoryWithFiles, "PdfHeaderFooterImages.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            var header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
            header.AddParagraph().AddImage(imagePath, 20, 20);
            footer.AddParagraph().AddImage(imagePath, 400, 400);
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_ToMemoryStream() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfStreamSample.docx");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");

            using (MemoryStream pdfStream = new MemoryStream()) {
                document.SaveAsPdf(pdfStream);
                Assert.True(pdfStream.Length > 0);
            }
        }
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_ToByteArray() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfByteArraySample.docx");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            byte[] bytes = document.SaveAsPdf();
            Assert.NotNull(bytes);
            Assert.True(bytes.Length > 100);
            string header = Encoding.ASCII.GetString(bytes, 0, 4);
            Assert.Equal("%PDF", header);
            string trailer = Encoding.ASCII.GetString(bytes, bytes.Length - 5, 5);
            Assert.Contains("EOF", trailer);
        }
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_ToFileStream() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfFileStreamSample.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfFileStreamSample.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            using (FileStream fileStream = File.Create(pdfPath)) {
                document.SaveAsPdf(fileStream);
            }
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_CreatesDirectory() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfCreateDir.docx");
        string pdfDir = Path.Combine(_directoryWithFiles, "MissingDir");
        string pdfPath = Path.Combine(pdfDir, "PdfCreateDir.pdf");

        if (Directory.Exists(pdfDir)) {
            Directory.Delete(pdfDir, true);
        }

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_NullPath_Throws() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNullPath.docx");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();
            Assert.Throws<ArgumentNullException>(() => document.SaveAsPdf((string)null!));
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void Test_WordDocument_SaveAsPdf_EmptyOrWhitespacePath_Throws(string path) {
        string docPath = Path.Combine(_directoryWithFiles, "PdfEmptyPath.docx");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();
            var ex = Assert.Throws<ArgumentException>(() => document.SaveAsPdf(path));
            Assert.Contains("empty or whitespace", ex.Message);
        }
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_NullDocument_Throws() {
        Assert.Throws<ArgumentNullException>(() => WordPdfConverterExtensions.SaveAsPdf(null!, "file.pdf"));
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_CustomParagraphFont() {
        string font = FontResolver.Resolve("monospace")!;
        string expected = Regex.Replace(font, @"\s+", "");
        string docPath = Path.Combine(_directoryWithFiles, "PdfCustomParagraphFont.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfCustomParagraphFont.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordParagraph p = document.AddParagraph("Hello World");
            p.FontFamily = font;
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains(expected, pdfContent);
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_CustomDefaultFont() {
        string font = FontResolver.Resolve("monospace")!;
        string expected = Regex.Replace(font, @"\s+", "");
        string docPath = Path.Combine(_directoryWithFiles, "PdfCustomDefaultFont.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfCustomDefaultFont.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions { FontFamily = font });
        }

        Assert.True(File.Exists(pdfPath));
        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains(expected, pdfContent);
    }

    [Theory]
    [InlineData(PdfPageOrientation.Portrait)]
    [InlineData(PdfPageOrientation.Landscape)]
    public void Test_WordDocument_SaveAsPdf_PageOrientation(PdfPageOrientation orientation) {
        string docPath = Path.Combine(_directoryWithFiles, $"PdfOrientation{orientation}.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, $"PdfOrientation{orientation}.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                Orientation = orientation,
                PageSize = PageSizes.A4
            });
        }

        Assert.True(File.Exists(pdfPath));

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Match mediaBox = Regex.Match(pdfContent, @"/MediaBox\s*\[\s*0\s+0\s+(?<w>[0-9\.]+)\s+(?<h>[0-9\.]+)\s*\]");
        Assert.True(mediaBox.Success, "MediaBox not found");
        double width = double.Parse(mediaBox.Groups["w"].Value, CultureInfo.InvariantCulture);
        double height = double.Parse(mediaBox.Groups["h"].Value, CultureInfo.InvariantCulture);
        if (orientation == PdfPageOrientation.Landscape) {
            Assert.True(width > height);
        } else {
            Assert.True(height > width);
        }
    }

    [Theory]
    [InlineData("Portrait")]
    [InlineData("Landscape")]
    public void Test_WordDocument_SaveAsPdf_SectionOrientationWithoutPageSize(string orientationValue) {
        PageOrientationValues orientation = orientationValue == "Landscape" ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
        string docPath = Path.Combine(_directoryWithFiles, $"PdfSectionOrientation{orientationValue}.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, $"PdfSectionOrientation{orientationValue}.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.Sections[0].PageSettings.Orientation = orientation;
            document.AddParagraph("Hello World");
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Match mediaBox = Regex.Match(pdfContent, @"/MediaBox\s*\[\s*0\s+0\s+(?<w>[0-9\.]+)\s+(?<h>[0-9\.]+)\s*\]");
        Assert.True(mediaBox.Success, "MediaBox not found");
        double width = double.Parse(mediaBox.Groups["w"].Value, CultureInfo.InvariantCulture);
        double height = double.Parse(mediaBox.Groups["h"].Value, CultureInfo.InvariantCulture);
        if (orientation == PageOrientationValues.Landscape) {
            Assert.True(width > height);
        } else {
            Assert.True(height > width);
        }
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_CustomPageSize() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfCustomSize.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfCustomSize.pdf");
        QuestPDF.Helpers.PageSize size = new QuestPDF.Helpers.PageSize(300, 500);

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");

            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                PageSize = size
            });
        }

        Assert.True(File.Exists(pdfPath));

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Match mediaBox = Regex.Match(pdfContent, @"/MediaBox\s*\[\s*0\s+0\s+(?<w>[0-9\.]+)\s+(?<h>[0-9\.]+)\s*\]");
        Assert.True(mediaBox.Success, "MediaBox not found");
        double width = double.Parse(mediaBox.Groups["w"].Value, CultureInfo.InvariantCulture);
        double height = double.Parse(mediaBox.Groups["h"].Value, CultureInfo.InvariantCulture);
        Assert.True(System.Math.Abs(width - size.Width) < 1);
        Assert.True(System.Math.Abs(height - size.Height) < 1);
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_CustomMargins() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfCustomMargins.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfCustomMargins.pdf");
        const double marginCm = 2;

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");

            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                Margin = (float)marginCm,
                MarginUnit = Unit.Centimetre
            });
        }

        Assert.True(File.Exists(pdfPath));

        byte[] bytes = File.ReadAllBytes(pdfPath);
        byte[] startPattern = Encoding.ASCII.GetBytes("stream\n");
        byte[] endPattern = Encoding.ASCII.GetBytes("\nendstream");
        int start = IndexOf(bytes, startPattern, 0);
        Assert.True(start >= 0, "stream marker not found");
        start += startPattern.Length;
        int end = IndexOf(bytes, endPattern, start);
        Assert.True(end >= 0, "endstream marker not found");
        int length = end - start;
        // remove zlib header (2 bytes) and Adler32 checksum (last 4 bytes)
        int deflateLength = length - 6;
        byte[] deflateData = new byte[deflateLength];
        Array.Copy(bytes, start + 2, deflateData, 0, deflateLength);
        using MemoryStream ms = new MemoryStream(deflateData);
        using DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress);
        using StreamReader reader = new StreamReader(ds, Encoding.GetEncoding("ISO-8859-1"));
        string content = reader.ReadToEnd();

        MatchCollection matches = Regex.Matches(content, @"4 0 0 4 (?<x>[0-9\.]+) (?<y>[0-9\.]+) cm");
        Assert.True(matches.Count > 0, "Margin transform not found");
        double value = 0;
        foreach (Match m in matches) {
            value = double.Parse(m.Groups["x"].Value, CultureInfo.InvariantCulture);
            if (value > 0) {
                break;
            }
        }
        double marginPoints = value / 4.0;
        double resultMarginCm = marginPoints / 28.3464566929;
        Assert.True(System.Math.Abs(resultMarginCm - marginCm) < 0.1);
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_MixedMargins() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfMixedMargins.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfMixedMargins.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                Margin = 2,
                MarginUnit = Unit.Centimetre,
                MarginLeft = 1,
                MarginLeftUnit = Unit.Centimetre,
                MarginTop = 3,
                MarginTopUnit = Unit.Centimetre
            });
        }

        Assert.True(File.Exists(pdfPath));

        byte[] bytes = File.ReadAllBytes(pdfPath);
        byte[] startPattern = Encoding.ASCII.GetBytes("stream\n");
        byte[] endPattern = Encoding.ASCII.GetBytes("\nendstream");
        int start = IndexOf(bytes, startPattern, 0);
        Assert.True(start >= 0, "stream marker not found");
        start += startPattern.Length;
        int end = IndexOf(bytes, endPattern, start);
        Assert.True(end >= 0, "endstream marker not found");
        int length = end - start;
        int deflateLength = length - 6;
        byte[] deflateData = new byte[deflateLength];
        Array.Copy(bytes, start + 2, deflateData, 0, deflateLength);
        using MemoryStream ms = new MemoryStream(deflateData);
        using DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress);
        using StreamReader reader = new StreamReader(ds, Encoding.GetEncoding("ISO-8859-1"));
        string content = reader.ReadToEnd();

        Match transform = Regex.Match(content, @"4 0 0 4 (?<x>[0-9\.]+) (?<y>[0-9\.]+) cm");
        Assert.True(transform.Success, "Margin transform not found");
        double leftMarginCm = double.Parse(transform.Groups["x"].Value, CultureInfo.InvariantCulture) / 4.0 / 28.3464566929;

        Assert.True(System.Math.Abs(leftMarginCm - 1) < 0.1);
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_Hyperlink() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfHyperlink.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfHyperlink.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHyperLink("OfficeIMO", new Uri("https://evotec.xyz"), addStyle: true);
            document.Save();

            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/URI (https://evotec.xyz", pdfContent);
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_SectionPageSettings() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfSectionPageSettings.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfSectionPageSettings.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.Sections[0].PageSettings.PageSize = WordPageSize.A4;
            document.Sections[0].PageSettings.Orientation = PageOrientationValues.Landscape;
            document.AddParagraph("Section1");

            WordSection section2 = document.AddSection();
            section2.PageSettings.PageSize = WordPageSize.A5;
            section2.PageSettings.Orientation = PageOrientationValues.Portrait;
            section2.AddParagraph("Section2");

            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        MatchCollection boxes = Regex.Matches(pdfContent, @"/MediaBox\s*\[\s*0\s+0\s+(?<w>[0-9\.]+)\s+(?<h>[0-9\.]+)\s*\]");
        Assert.Equal(2, boxes.Count);

        double w1 = double.Parse(boxes[0].Groups["w"].Value, CultureInfo.InvariantCulture);
        double h1 = double.Parse(boxes[0].Groups["h"].Value, CultureInfo.InvariantCulture);
        QuestPDF.Helpers.PageSize a4 = PageSizes.A4.Landscape();
        Assert.True(System.Math.Abs(w1 - a4.Width) < 1);
        Assert.True(System.Math.Abs(h1 - a4.Height) < 1);

        double w2 = double.Parse(boxes[1].Groups["w"].Value, CultureInfo.InvariantCulture);
        double h2 = double.Parse(boxes[1].Groups["h"].Value, CultureInfo.InvariantCulture);
        Assert.True(System.Math.Abs(w2 - PageSizes.A5.Width) < 1);
        Assert.True(System.Math.Abs(h2 - PageSizes.A5.Height) < 1);

        Assert.True(w1 > h1);
        Assert.True(h2 > w2);
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_BookmarkLink() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfBookmarkLink.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfBookmarkLink.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            var target = document.AddParagraph("Bookmark target");
            target.AddBookmark("TargetBookmark");
            document.AddHyperLink("Go to bookmark", "TargetBookmark");
            document.Save();

            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/Dest /0#20|#20TargetBookmark", pdfContent);
    }

    private static int IndexOf(byte[] buffer, byte[] pattern, int start) {
        for (int i = start; i <= buffer.Length - pattern.Length; i++) {
            int j = 0;
            for (; j < pattern.Length; j++) {
                if (buffer[i + j] != pattern[j]) {
                    break;
                }
            }
            if (j == pattern.Length) {
                return i;
            }
        }
        return -1;
    }
}
