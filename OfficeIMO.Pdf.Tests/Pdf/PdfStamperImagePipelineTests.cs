using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfStamperTests {
    [Fact]
    public void StampImage_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildTwoPagePdf());

        byte[] stamped = PdfStamper.StampImage(stream, CreateMinimalRgbPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            X = 72,
            Y = 650,
            Width = 24,
            Height = 24
        });

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);

        string text = Normalize(PdfReadDocument.Open(stamped).ExtractText());
        Assert.Contains("Firstpagebody", text);
        Assert.Contains("Secondpagebody", text);
    }

    [Fact]
    public void StampImage_ReadsImageStreamFromCurrentPosition() {
        byte[] source = BuildTwoPagePdf();
        using var imageStream = CreatePrefixedStream(CreateMinimalRgbPng());

        byte[] stamped = PdfStamper.StampImage(source, imageStream, new PdfImageStampOptions {
            PageNumbers = new[] { 2 },
            X = 72,
            Y = 650,
            Width = 24,
            Height = 24
        });

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
    }

    [Fact]
    public void StampImage_ReadsPdfAndImageStreamsFromCurrentPositions() {
        using var pdfStream = CreatePrefixedStream(BuildTwoPagePdf());
        using var imageStream = CreatePrefixedStream(CreateMinimalRgbPng());

        byte[] stamped = PdfStamper.StampImage(pdfStream, imageStream, new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            X = 72,
            Y = 650,
            Width = 24,
            Height = 24
        });

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);

        string text = Normalize(PdfReadDocument.Open(stamped).ExtractText());
        Assert.Contains("Firstpagebody", text);
        Assert.Contains("Secondpagebody", text);
    }

    [Fact]
    public void StampImage_WritesPathInputAndImageStreamToPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-stamp-image-stream-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "stamped-image.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildTwoPagePdf());
            using var imageStream = CreatePrefixedStream(CreateMinimalRgbPng());

            PdfStamper.StampImage(inputPath, outputPath, imageStream, new PdfImageStampOptions {
                PageNumbers = new[] { 1 },
                X = 72,
                Y = 650,
                Width = 24,
                Height = 24
            });

            Assert.True(File.Exists(outputPath));
            byte[] stamped = File.ReadAllBytes(outputPath);
            string pdfContent = Encoding.ASCII.GetString(stamped);
            Assert.Contains("/Subtype /Image", pdfContent);
            Assert.Contains("/OIMOStampIm", pdfContent);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void StampImage_WritesPdfAndImageStreamsToOutputStreamAtCurrentPosition() {
        using var pdfStream = CreatePrefixedStream(BuildTwoPagePdf());
        using var imageStream = CreatePrefixedStream(CreateMinimalRgbPng());
        using var output = CreateOutputStream(out int prefixLength);

        PdfStamper.StampImage(pdfStream, output, imageStream, new PdfImageStampOptions {
            PageNumbers = new[] { 2 },
            X = 72,
            Y = 650,
            Width = 24,
            Height = 24
        });

        byte[] stamped = GetOutputPayload(output, prefixLength);
        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);

        string text = Normalize(PdfReadDocument.Open(stamped).ExtractText());
        Assert.Contains("Firstpagebody", text);
        Assert.Contains("Secondpagebody", text);
    }

    [Fact]
    public void WatermarkImage_AddsCenteredImageToEveryPage() {
        byte[] source = BuildTwoPagePdf();
        byte[] image = CreateMinimalRgbPng();

        byte[] stamped = PdfStamper.WatermarkImage(source, image, new PdfImageStampOptions {
            Width = 32,
            Height = 32
        });

        using var pdf = PdfPigDocument.Open(new MemoryStream(stamped));
        Assert.Equal(2, pdf.NumberOfPages);

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/XObject", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
        Assert.Contains(" Do", pdfContent);
    }

    [Fact]
    public void WatermarkImage_WritesByteInputAndImageStreamToOutputStreamAtCurrentPosition() {
        byte[] source = BuildTwoPagePdf();
        using var imageStream = CreatePrefixedStream(CreateMinimalRgbPng());
        using var output = CreateOutputStream(out int prefixLength);

        PdfStamper.WatermarkImage(source, output, imageStream, new PdfImageStampOptions {
            Width = 32,
            Height = 32
        });

        byte[] stamped = GetOutputPayload(output, prefixLength);
        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
        Assert.Contains(" Do", pdfContent);
    }

    [Fact]
    public void WatermarkImage_WritesPdfAndImageStreamsToPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-watermark-image-stream-" + Guid.NewGuid().ToString("N"));
        string outputPath = Path.Combine(directory, "out", "watermark-image.pdf");

        try {
            using var pdfStream = CreatePrefixedStream(BuildTwoPagePdf());
            using var imageStream = CreatePrefixedStream(CreateMinimalRgbPng());

            PdfStamper.WatermarkImage(pdfStream, outputPath, imageStream, new PdfImageStampOptions {
                Width = 32,
                Height = 32
            });

            Assert.True(File.Exists(outputPath));
            byte[] stamped = File.ReadAllBytes(outputPath);
            string pdfContent = Encoding.ASCII.GetString(stamped);
            Assert.Contains("/Subtype /Image", pdfContent);
            Assert.Contains("/OIMOStampIm", pdfContent);
            Assert.Contains(" Do", pdfContent);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void WatermarkImage_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildTwoPagePdf());

        byte[] stamped = PdfStamper.WatermarkImage(stream, CreateMinimalRgbPng(), new PdfImageStampOptions {
            Width = 32,
            Height = 32
        });

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
        Assert.Contains(" Do", pdfContent);
    }

    [Fact]
    public void WatermarkImage_ReadsImageStreamFromCurrentPosition() {
        byte[] source = BuildTwoPagePdf();
        using var imageStream = CreatePrefixedStream(CreateMinimalRgbPng());

        byte[] stamped = PdfStamper.WatermarkImage(source, imageStream, new PdfImageStampOptions {
            Width = 32,
            Height = 32
        });

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
        Assert.Contains(" Do", pdfContent);
    }

    [Fact]
    public void WatermarkImage_ReadsPdfAndImageStreamsFromCurrentPositions() {
        using var pdfStream = CreatePrefixedStream(BuildTwoPagePdf());
        using var imageStream = CreatePrefixedStream(CreateMinimalRgbPng());

        byte[] stamped = PdfStamper.WatermarkImage(pdfStream, imageStream, new PdfImageStampOptions {
            Width = 32,
            Height = 32
        });

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
        Assert.Contains(" Do", pdfContent);
    }
}
