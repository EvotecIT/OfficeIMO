using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
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
            document.Header.Default.AddParagraph("Sample Header");
            WordTable headerTable = document.Header.Default.AddTable(1, 1);
            headerTable.Rows[0].Cells[0].Paragraphs[0].Text = "H1";
            document.Footer.Default.AddParagraph("Sample Footer");
            WordTable footerTable = document.Footer.Default.AddTable(1, 1);
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
    public void Test_WordDocument_SaveAsPdf_WithMargins() {
        var docPath = Path.Combine(_directoryWithFiles, "PdfMargin.docx");
        var defaultPdf = Path.Combine(_directoryWithFiles, "PdfMarginDefault.pdf");
        var customPdf = Path.Combine(_directoryWithFiles, "PdfMarginCustom.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Margins test");
            document.Save();
            document.SaveAsPdf(defaultPdf);
            PdfSaveOptions options = new PdfSaveOptions {
                MarginLeft = 2,
                MarginTop = 2,
                MarginRight = 2,
                MarginBottom = 2
            };
            document.SaveAsPdf(customPdf, options);
        }

        double defaultLeft = ExtractMargin(defaultPdf);
        double customLeft = ExtractMargin(customPdf);

        Assert.True(customLeft > defaultLeft + 20);

        static double ExtractMargin(string pdfPath) {
            byte[] bytes = File.ReadAllBytes(pdfPath);
            byte[] streamMarker = Encoding.ASCII.GetBytes("stream");
            byte[] endMarker = Encoding.ASCII.GetBytes("endstream");
            int streamIndex = IndexOf(bytes, streamMarker, 0);
            int dataStart = streamIndex + streamMarker.Length;
            if (bytes[dataStart] == 13 && bytes[dataStart + 1] == 10) {
                dataStart += 2;
            } else if (bytes[dataStart] == 10 || bytes[dataStart] == 13) {
                dataStart += 1;
            }
            int endIndex = IndexOf(bytes, endMarker, dataStart);
            using MemoryStream ms = new MemoryStream(bytes, dataStart, endIndex - dataStart);
            using ZLibStream z = new ZLibStream(ms, CompressionMode.Decompress);
            using StreamReader sr = new StreamReader(z);
            string content = sr.ReadToEnd();
            MatchCollection matches = Regex.Matches(content, @"4 0 0 4 ([0-9\.]+) [0-9\.]+ cm");
            if (matches.Count > 0) {
                Match match = matches[matches.Count - 1];
                return double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
            }
            return double.NaN;
        }

        static int IndexOf(byte[] array, byte[] pattern, int start) {
            for (int i = start; i <= array.Length - pattern.Length; i++) {
                bool match = true;
                for (int j = 0; j < pattern.Length; j++) {
                    if (array[i + j] != pattern[j]) {
                        match = false;
                        break;
                    }
                }
                if (match) {
                    return i;
                }
            }
            return -1;
        }
    }
}