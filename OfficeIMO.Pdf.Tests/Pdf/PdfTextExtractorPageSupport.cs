using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfTextExtractorPageTests {
    private static byte[] BuildThreePagePdf() {
        var doc = PdfDocument.Create();
        doc.Compose(compose => {
            compose.Page(page => {
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("First page marker"))));
            });

            compose.Page(page => {
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Second page marker"))));
            });

            compose.Page(page => {
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Third page marker"))));
            });
        });

        return doc.ToBytes();
    }

    private static byte[] BuildTwoColumnPdf() {
        return PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row => row
                            .Gap(36)
                            .Column(50, column => column
                                .Paragraph(p => p.Text("Left Start marker"))
                                .Paragraph(p => p.Text("Left Finish marker")))
                            .Column(50, column => column
                                .Paragraph(p => p.Text("Right Start marker"))
                                .Paragraph(p => p.Text("Right Finish marker")))))))
            .ToBytes();
    }

    private static byte[] BuildMarkdownPdf() {
        return PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("Markdown Heading")
            .Paragraph(p => p.Text("Markdown readback marker."))
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
            .ToBytes();
    }

    private static byte[] BuildThreePageMarkdownPdf() {
        return PdfDocument.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 10
            })
            .H1("First Page")
            .Paragraph(p => p.Text("First markdown marker."))
            .PageBreak()
            .H1("Second Page")
            .Paragraph(p => p.Text("Second markdown marker."))
            .PageBreak()
            .H1("Third Page")
            .Paragraph(p => p.Text("Third markdown marker."))
            .ToBytes();
    }

    private static MemoryStream BuildPrefixedStream(byte[] pdf) {
        var data = new byte[pdf.Length + 5];
        data[0] = 1;
        data[1] = 2;
        data[2] = 3;
        data[3] = 4;
        data[4] = 5;
        Array.Copy(pdf, 0, data, 5, pdf.Length);
        return new MemoryStream(data);
    }

    private static MemoryStream CreateOutputStream(out int prefixLength) {
        byte[] prefix = Encoding.ASCII.GetBytes("output-prefix");
        var stream = new MemoryStream();
        stream.Write(prefix, 0, prefix.Length);
        prefixLength = prefix.Length;
        return stream;
    }

    private static string GetOutputText(MemoryStream output, int prefixLength) {
        byte[] bytes = output.ToArray();
        Assert.True(bytes.Length > prefixLength);
        Assert.Equal("output-prefix", Encoding.ASCII.GetString(bytes, 0, prefixLength));
        return Encoding.UTF8.GetString(bytes, prefixLength, bytes.Length - prefixLength);
    }

    private static string Normalize(string text) {
        return text.Replace(" ", string.Empty);
    }

    private static void AssertColumnAwareTextOrder(string text) {
        string normalized = Normalize(text);
        int leftStart = normalized.IndexOf("LeftStart", StringComparison.Ordinal);
        int leftFinish = normalized.IndexOf("LeftFinish", StringComparison.Ordinal);
        int rightStart = normalized.IndexOf("RightStart", StringComparison.Ordinal);
        int rightFinish = normalized.IndexOf("RightFinish", StringComparison.Ordinal);

        Assert.True(leftStart >= 0, "Expected left column start marker to be extracted.");
        Assert.True(leftFinish > leftStart, "Expected left column markers to preserve top-to-bottom order.");
        Assert.True(rightStart >= 0, "Expected right column start marker to be extracted.");
        Assert.True(rightFinish > rightStart, "Expected right column markers to preserve top-to-bottom order.");
        Assert.True(leftFinish < rightStart,
            $"Expected column-aware extraction to finish the left column before reading the right column. Text: {text}");
    }

    private static void AssertContainsInOrder(string text, params string[] markers) {
        int lastIndex = -1;
        for (int i = 0; i < markers.Length; i++) {
            int index = text.IndexOf(markers[i], lastIndex + 1, StringComparison.Ordinal);
            Assert.True(index > lastIndex, $"Expected marker '{markers[i]}' after index {lastIndex}. Text: {text}");
            lastIndex = index;
        }
    }

    private static int CountOccurrences(string text, string marker) {
        int count = 0;
        int index = 0;
        while (true) {
            index = text.IndexOf(marker, index, StringComparison.Ordinal);
            if (index < 0) {
                return count;
            }

            count++;
            index += marker.Length;
        }
    }

    private sealed class WriteOnlyStream : Stream {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => 0;

        public override long Position {
            get => 0;
            set => throw new NotSupportedException();
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin) {
            throw new NotSupportedException();
        }

        public override void SetLength(long value) {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count) {
        }
    }

    private sealed class ReadOnlyStream : MemoryStream {
        public override bool CanWrite => false;
    }
}
