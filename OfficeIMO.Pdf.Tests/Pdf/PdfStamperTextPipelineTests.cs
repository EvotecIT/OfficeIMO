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
    public void StampText_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildTwoPagePdf());

        byte[] stamped = PdfStamper.StampText(stream, "STREAM-STAMP", new PdfTextStampOptions {
            PageNumbers = new[] { 1 },
            X = 72,
            Y = 700,
            FontSize = 14
        });

        var read = PdfReadDocument.Open(stamped);
        Assert.Contains("STREAM-STAMP", Normalize(read.Pages[0].ExtractText()));
        Assert.DoesNotContain("STREAM-STAMP", Normalize(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void WatermarkText_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildTwoPagePdf());

        byte[] stamped = PdfStamper.WatermarkText(stream, "STREAM-DRAFT");

        string text = Normalize(PdfReadDocument.Open(stamped).ExtractText());
        Assert.Contains("STREAM-DRAFT", text);
        Assert.Contains("Firstpagebody", text);
        Assert.Contains("Secondpagebody", text);
    }

    [Fact]
    public void StampText_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-stamp-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "stamped.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildTwoPagePdf());

            PdfStamper.StampText(inputPath, outputPath, "PATH-STAMP", new PdfTextStampOptions {
                PageNumbers = new[] { 1 },
                X = 80,
                Y = 650,
                FontSize = 14
            });

            Assert.True(File.Exists(outputPath));
            string text = Normalize(PdfReadDocument.Open(outputPath).ExtractText());
            Assert.Contains("PATH-STAMP", text);
            Assert.Contains("Firstpagebody", text);
            Assert.Contains("Secondpagebody", text);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void StamperPathInputs_ReturnBytesForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-stamp-path-bytes-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildTwoPagePdf());

            byte[] textStamped = PdfStamper.StampTextToBytes(inputPath, "PATH-BYTES-STAMP", new PdfTextStampOptions {
                PageNumbers = new[] { 1 },
                X = 80,
                Y = 650,
                FontSize = 14
            });

            var stampedRead = PdfReadDocument.Open(textStamped);
            Assert.Contains("PATH-BYTES-STAMP", Normalize(stampedRead.Pages[0].ExtractText()));
            Assert.DoesNotContain("PATH-BYTES-STAMP", Normalize(stampedRead.Pages[1].ExtractText()));

            byte[] textWatermarked = PdfStamper.WatermarkTextToBytes(inputPath, "PATH-BYTES-DRAFT");
            Assert.Contains("PATH-BYTES-DRAFT", Normalize(PdfReadDocument.Open(textWatermarked).ExtractText()));

            byte[] imageStamped = PdfStamper.StampImageToBytes(inputPath, CreateMinimalRgbPng(), new PdfImageStampOptions {
                PageNumbers = new[] { 2 },
                X = 72,
                Y = 650,
                Width = 24,
                Height = 24
            });

            string imageStampedContent = Encoding.ASCII.GetString(imageStamped);
            Assert.Contains("/Subtype /Image", imageStampedContent);
            Assert.Contains("/OIMOStampIm", imageStampedContent);

            using var imageStream = CreatePrefixedStream(CreateMinimalRgbPng());
            byte[] imageWatermarked = PdfStamper.WatermarkImageToBytes(inputPath, imageStream, new PdfImageStampOptions {
                Width = 32,
                Height = 32
            });

            string imageWatermarkedContent = Encoding.ASCII.GetString(imageWatermarked);
            Assert.Contains("/Subtype /Image", imageWatermarkedContent);
            Assert.Contains("/OIMOStampIm", imageWatermarkedContent);
            Assert.Contains(" Do", imageWatermarkedContent);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void StamperPathInputs_WriteToOutputStreamsForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-stamp-path-stream-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildTwoPagePdf());

            using var textOutput = CreateOutputStream(out int textPrefixLength);
            PdfStamper.StampText(inputPath, textOutput, "PATH-STREAM-STAMP", new PdfTextStampOptions {
                PageNumbers = new[] { 1 },
                X = 80,
                Y = 650,
                FontSize = 14
            });

            byte[] textStamped = GetOutputPayload(textOutput, textPrefixLength);
            var stampedRead = PdfReadDocument.Open(textStamped);
            Assert.Contains("PATH-STREAM-STAMP", Normalize(stampedRead.Pages[0].ExtractText()));
            Assert.DoesNotContain("PATH-STREAM-STAMP", Normalize(stampedRead.Pages[1].ExtractText()));

            using var watermarkOutput = CreateOutputStream(out int watermarkPrefixLength);
            PdfStamper.WatermarkText(inputPath, watermarkOutput, "PATH-STREAM-DRAFT");

            byte[] textWatermarked = GetOutputPayload(watermarkOutput, watermarkPrefixLength);
            Assert.Contains("PATH-STREAM-DRAFT", Normalize(PdfReadDocument.Open(textWatermarked).ExtractText()));

            using var imageOutput = CreateOutputStream(out int imagePrefixLength);
            PdfStamper.StampImage(inputPath, imageOutput, CreateMinimalRgbPng(), new PdfImageStampOptions {
                PageNumbers = new[] { 2 },
                X = 72,
                Y = 650,
                Width = 24,
                Height = 24
            });

            string imageStampedContent = Encoding.ASCII.GetString(GetOutputPayload(imageOutput, imagePrefixLength));
            Assert.Contains("/Subtype /Image", imageStampedContent);
            Assert.Contains("/OIMOStampIm", imageStampedContent);

            using var imageStream = CreatePrefixedStream(CreateMinimalRgbPng());
            using var imageWatermarkOutput = CreateOutputStream(out int imageWatermarkPrefixLength);
            PdfStamper.WatermarkImage(inputPath, imageWatermarkOutput, imageStream, new PdfImageStampOptions {
                Width = 32,
                Height = 32
            });

            string imageWatermarkedContent = Encoding.ASCII.GetString(GetOutputPayload(imageWatermarkOutput, imageWatermarkPrefixLength));
            Assert.Contains("/Subtype /Image", imageWatermarkedContent);
            Assert.Contains("/OIMOStampIm", imageWatermarkedContent);
            Assert.Contains(" Do", imageWatermarkedContent);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void StampText_WritesStreamInputToPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-stamp-stream-" + Guid.NewGuid().ToString("N"));
        string outputPath = Path.Combine(directory, "out", "stamped.pdf");

        try {
            using var stream = CreatePrefixedStream(BuildTwoPagePdf());

            PdfStamper.StampText(stream, outputPath, "STREAM-PATH-STAMP", new PdfTextStampOptions {
                PageNumbers = new[] { 2 },
                X = 80,
                Y = 650,
                FontSize = 14
            });

            Assert.True(File.Exists(outputPath));
            var read = PdfReadDocument.Open(outputPath);
            Assert.DoesNotContain("STREAM-PATH-STAMP", Normalize(read.Pages[0].ExtractText()));
            Assert.Contains("STREAM-PATH-STAMP", Normalize(read.Pages[1].ExtractText()));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void StampText_WritesByteInputToOutputStreamAtCurrentPosition() {
        byte[] source = BuildTwoPagePdf();
        using var output = CreateOutputStream(out int prefixLength);

        PdfStamper.StampText(source, output, "OUTPUT-STAMP", new PdfTextStampOptions {
            PageNumbers = new[] { 1 },
            X = 72,
            Y = 700,
            FontSize = 14
        });

        byte[] stamped = GetOutputPayload(output, prefixLength);
        var read = PdfReadDocument.Open(stamped);
        Assert.Contains("OUTPUT-STAMP", Normalize(read.Pages[0].ExtractText()));
        Assert.DoesNotContain("OUTPUT-STAMP", Normalize(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void WatermarkText_WritesStreamInputToOutputStreamAtCurrentPosition() {
        using var stream = CreatePrefixedStream(BuildTwoPagePdf());
        using var output = CreateOutputStream(out int prefixLength);

        PdfStamper.WatermarkText(stream, output, "OUTPUT-DRAFT");

        byte[] stamped = GetOutputPayload(output, prefixLength);
        string text = Normalize(PdfReadDocument.Open(stamped).ExtractText());
        Assert.Contains("OUTPUT-DRAFT", text);
        Assert.Contains("Firstpagebody", text);
        Assert.Contains("Secondpagebody", text);
    }

    [Fact]
    public void WatermarkText_WritesStreamInputToPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-watermark-stream-" + Guid.NewGuid().ToString("N"));
        string outputPath = Path.Combine(directory, "out", "watermark.pdf");

        try {
            using var stream = CreatePrefixedStream(BuildTwoPagePdf());

            PdfStamper.WatermarkText(stream, outputPath, "STREAM-PATH-DRAFT");

            Assert.True(File.Exists(outputPath));
            string text = Normalize(PdfReadDocument.Open(outputPath).ExtractText());
            Assert.Contains("STREAM-PATH-DRAFT", text);
            Assert.Contains("Firstpagebody", text);
            Assert.Contains("Secondpagebody", text);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}
