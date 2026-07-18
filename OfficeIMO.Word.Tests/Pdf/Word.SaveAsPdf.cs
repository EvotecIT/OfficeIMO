using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

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

            byte[] bytes = document.ToPdf();
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
    public void Test_WordDocument_SaveAsPdf_DirectoryPath_Throws() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfDirectoryPath.docx");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            var ex = Assert.Throws<ArgumentException>(() => document.SaveAsPdf(_directoryWithFiles));
            Assert.Contains("directory", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_NullDocument_Throws() {
        Assert.Throws<ArgumentNullException>(() => WordPdfConverterExtensions.SaveAsPdf(null!, "file.pdf"));
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_CustomParagraphFont() {
        string font = "Courier New";
        string docPath = Path.Combine(_directoryWithFiles, "PdfCustomParagraphFont.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfCustomParagraphFont.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordParagraph p = document.AddParagraph("Hello World");
            p.FontFamily = font;
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
        AssertPdfUsesAnyFont(pdfPath, "Courier New", "Courier");
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_CustomDefaultFont() {
        string font = "Times New Roman";
        string docPath = Path.Combine(_directoryWithFiles, "PdfCustomDefaultFont.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfCustomDefaultFont.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions { FontFamily = font });
        }

        Assert.True(File.Exists(pdfPath));
        AssertPdfUsesFont(pdfPath, "Times");
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_ExplicitMappedDefaultFontFamily_StaysOnStandardFamily() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfExplicitSerifDefaultFont.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfExplicitSerifDefaultFont.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello explicit serif default font");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions { FontFamily = "serif" });
        }

        Assert.True(File.Exists(pdfPath));
        AssertPdfUsesFont(pdfPath, "Times");
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_DocumentDefaultFont() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfDocumentDefaultFont.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfDocumentDefaultFont.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.Settings.FontFamily = "Consolas";
            document.AddParagraph("Hello document default font");
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
        AssertPdfUsesAnyFont(pdfPath, "Consolas", "Courier");
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_DocumentMappedDefaultFont_StaysOnStandardFamily() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfDocumentMappedDefaultFont.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfDocumentMappedDefaultFont.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.Settings.FontFamily = "serif";
            document.AddParagraph("Hello mapped document default font");
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
        AssertPdfUsesFont(pdfPath, "Times");
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_DocumentDefaultFont_FallsBack_To_HighAnsi_When_Primary_Font_Is_Unmapped() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfDocumentDefaultFontHighAnsiFallback.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfDocumentDefaultFontHighAnsiFallback.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.Settings.FontFamily = "OfficeIMO Missing Theme Font";
            document.Settings.FontFamilyHighAnsi = "Times New Roman";
            document.AddParagraph("Hello HighAnsi fallback font");
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
        AssertPdfUsesFont(pdfPath, "Times");
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_RequestedUnavailableFont_FallsBack_To_DocumentDefaultFont() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfUnavailableRequestedFontFallsBack.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfUnavailableRequestedFontFallsBack.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.Settings.FontFamily = "OfficeIMO Missing Theme Font";
            document.Settings.FontFamilyHighAnsi = "Times New Roman";
            document.AddParagraph("Hello unavailable requested font fallback");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions { FontFamily = "OfficeIMO Missing Requested Font" });
        }

        Assert.True(File.Exists(pdfPath));
        AssertPdfUsesFont(pdfPath, "Times");
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_Preserves_Default_And_Distinct_Run_Fonts_During_Font_Prepass() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfDefaultFontSlotPrepass.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfDefaultFontSlotPrepass.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Styled serif").SetFontFamily("Georgia");
            document.AddParagraph("Default serif");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions { FontFamily = "Times New Roman" });
        }

        Assert.True(File.Exists(pdfPath));
        AssertPdfUsesFont(pdfPath, "Times");
        if (PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Georgia", out _)) {
            AssertPdfUsesFont(pdfPath, "Georgia");
        } else {
            AssertPdfDoesNotUseFont(pdfPath, "Georgia");
        }
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
                PageSize = PdfCore.PageSizes.A4
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
        PdfCore.PageSize size = new PdfCore.PageSize(300, 500);

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
                Margins = PdfCore.PageMargins.UniformCentimeters(marginCm)
            });
        }

        Assert.True(File.Exists(pdfPath));

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var hello = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Hello");
        double expectedMarginPoints = marginCm * 72D / 2.54D;
        Assert.InRange(hello.BoundingBox.Left, expectedMarginPoints - 2D, expectedMarginPoints + 4D);
    }

    [Fact]
    public void Test_WordDocument_SaveAsPdf_MixedMargins() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfMixedMargins.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfMixedMargins.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Hello World");
            document.Save();

            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                Margins = PdfCore.PageMargins.FromCentimeters(1, 3, 2, 2)
            });
        }

        Assert.True(File.Exists(pdfPath));

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var hello = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Hello");
        double expectedLeftMarginPoints = 72D / 2.54D;
        Assert.InRange(hello.BoundingBox.Left, expectedLeftMarginPoints - 2D, expectedLeftMarginPoints + 4D);
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
        PdfCore.PageSize a4 = PdfCore.PageSizes.A4.Landscape();
        Assert.True(System.Math.Abs(w1 - a4.Width) < 1);
        Assert.True(System.Math.Abs(h1 - a4.Height) < 1);

        double w2 = double.Parse(boxes[1].Groups["w"].Value, CultureInfo.InvariantCulture);
        double h2 = double.Parse(boxes[1].Groups["h"].Value, CultureInfo.InvariantCulture);
        Assert.True(System.Math.Abs(w2 - PdfCore.PageSizes.A5.Width) < 1);
        Assert.True(System.Math.Abs(h2 - PdfCore.PageSizes.A5.Height) < 1);

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

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(File.ReadAllBytes(pdfPath), new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        Assert.Contains(logical.NamedDestinations, destination => destination.Name == "TargetBookmark");
        Assert.Contains(logical.GetLinksByDestinationName("TargetBookmark"), link => link.IsNamedDestinationLink);
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

    private static string ReadFirstPdfStreamContent(byte[] bytes) {
        int start = FindPdfStreamDataStart(bytes);
        Assert.True(start >= 0, "stream marker not found");

        byte[] endPattern = Encoding.ASCII.GetBytes("endstream");
        int end = IndexOf(bytes, endPattern, start);
        Assert.True(end >= 0, "endstream marker not found");

        int length = end - start;
        while (length > 0 && (bytes[start + length - 1] == '\r' || bytes[start + length - 1] == '\n')) {
            length--;
        }

        byte[] streamData = new byte[length];
        Array.Copy(bytes, start, streamData, 0, length);

        if (TryInflatePdfStream(streamData, 0, streamData.Length, out string inflated)) {
            return inflated;
        }

        if (streamData.Length > 6 &&
            TryInflatePdfStream(streamData, 2, streamData.Length - 6, out inflated)) {
            return inflated;
        }

        return Encoding.GetEncoding("ISO-8859-1").GetString(streamData);
    }

    private static string ReadPdfPageContent(byte[] bytes, int pageNumber = 1) {
        var document = PdfCore.PdfReadDocument.Open(bytes);
        var (objects, _) = PdfCore.PdfSyntax.ParseObjects(bytes);
        int pageObjectNumber = document.Pages[pageNumber - 1].ObjectNumber;
        if (!objects.TryGetValue(pageObjectNumber, out PdfCore.PdfIndirectObject pageObject) ||
            pageObject.Value is not PdfCore.PdfDictionary pageDictionary) {
            throw new InvalidOperationException("Page object was not found.");
        }

        if (!pageDictionary.Items.TryGetValue("Contents", out PdfCore.PdfObject contents)) {
            throw new InvalidOperationException("Page contents were not found.");
        }

        var streams = new List<string>();
        AppendPdfPageContentStreams(objects, contents, streams);
        return string.Join("\n", streams);
    }

    private static void AppendPdfPageContentStreams(
        Dictionary<int, PdfCore.PdfIndirectObject> objects,
        PdfCore.PdfObject contents,
        List<string> streams) {
        if (contents is PdfCore.PdfReference reference) {
            if (objects.TryGetValue(reference.ObjectNumber, out PdfCore.PdfIndirectObject indirect) &&
                indirect.Value is PdfCore.PdfStream stream) {
                streams.Add(Encoding.GetEncoding("ISO-8859-1").GetString(stream.Data));
            }

            return;
        }

        if (contents is PdfCore.PdfArray array) {
            foreach (PdfCore.PdfObject item in array.Items) {
                AppendPdfPageContentStreams(objects, item, streams);
            }
        }
    }

    private static int FindPdfStreamDataStart(byte[] bytes) {
        byte[] lfPattern = Encoding.ASCII.GetBytes("stream\n");
        int lfStart = IndexOf(bytes, lfPattern, 0);
        if (lfStart >= 0) {
            return lfStart + lfPattern.Length;
        }

        byte[] crlfPattern = Encoding.ASCII.GetBytes("stream\r\n");
        int crlfStart = IndexOf(bytes, crlfPattern, 0);
        return crlfStart >= 0 ? crlfStart + crlfPattern.Length : -1;
    }

    private static bool TryInflatePdfStream(byte[] bytes, int offset, int count, out string content) {
        try {
            using var source = new MemoryStream(bytes, offset, count);
            using var deflate = new System.IO.Compression.DeflateStream(source, System.IO.Compression.CompressionMode.Decompress);
            using var reader = new StreamReader(deflate, Encoding.GetEncoding("ISO-8859-1"));
            content = reader.ReadToEnd();
            return !string.IsNullOrEmpty(content);
        } catch (InvalidDataException) {
            content = string.Empty;
            return false;
        }
    }

    private static void AssertPdfUsesFont(string pdfPath, string expectedFontNamePart) {
        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.Contains(pdf.GetPage(1).Letters, letter =>
            letter.FontName != null &&
            letter.FontName.Contains(expectedFontNamePart, StringComparison.OrdinalIgnoreCase));

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/BaseFont /" + expectedFontNamePart, pdfContent, StringComparison.OrdinalIgnoreCase);
    }

    private static void AssertPdfUsesAnyFont(string pdfPath, params string[] expectedFontNameParts) {
        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.Contains(pdf.GetPage(1).Letters, letter =>
            letter.FontName != null &&
            expectedFontNameParts.Any(expected => letter.FontName.Contains(expected, StringComparison.OrdinalIgnoreCase)));

        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains(expectedFontNameParts, expected =>
            pdfContent.Contains("/BaseFont /" + expected, StringComparison.OrdinalIgnoreCase));
    }

    private static void AssertPdfDoesNotUseFont(string pdfPath, string fontNamePart) {
        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.DoesNotContain("/BaseFont /" + fontNamePart, pdfContent, StringComparison.OrdinalIgnoreCase);
    }
}
