using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfStamperTests {
    [Fact]
    public void StampText_AddsTextToSelectedPageAndPreservesOriginalContent() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampText(source, "APPROVED", new PdfTextStampOptions {
            PageNumbers = new[] { 2 },
            X = 72,
            Y = 700,
            FontSize = 16,
            Color = PdfColor.Black
        });

        using var pdf = PdfDocument.Open(new MemoryStream(stamped));
        Assert.Equal(2, pdf.NumberOfPages);

        var read = PdfReadDocument.Load(stamped);
        string firstPage = Normalize(read.Pages[0].ExtractText());
        string secondPage = Normalize(read.Pages[1].ExtractText());

        Assert.Contains("Firstpagebody", firstPage);
        Assert.DoesNotContain("APPROVED", firstPage);
        Assert.Contains("Secondpagebody", secondPage);
        Assert.Contains("APPROVED", secondPage);
    }

    [Fact]
    public void StampText_UsesInclusivePageRangeSelection() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampText(source, "RANGE-STAMP", new PdfTextStampOptions()
            .UsePageRange(2, 2));

        var read = PdfReadDocument.Load(stamped);
        Assert.DoesNotContain("RANGE-STAMP", Normalize(read.Pages[0].ExtractText()));
        Assert.Contains("RANGE-STAMP", Normalize(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void StampText_UsesInclusivePageRangeListSelectionAndDeduplicatesOverlap() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampText(source, "RANGE-LIST-STAMP", new PdfTextStampOptions()
            .UsePageRanges(PdfPageRange.ParseMany("1-2,2")));

        var read = PdfReadDocument.Load(stamped);
        string firstPage = Normalize(read.Pages[0].ExtractText());
        string secondPage = Normalize(read.Pages[1].ExtractText());

        Assert.Contains("RANGE-LIST-STAMP", firstPage);
        Assert.Contains("RANGE-LIST-STAMP", secondPage);
        Assert.Equal(1, CountOccurrences(firstPage, "RANGE-LIST-STAMP"));
        Assert.Equal(1, CountOccurrences(secondPage, "RANGE-LIST-STAMP"));
    }

    [Fact]
    public void StampText_FlattensReferencedContentArraysBeforeAddingStampStream() {
        byte[] stamped = PdfStamper.StampText(BuildIndirectContentsArrayPdf(), "STAMP", new PdfTextStampOptions {
            X = 20,
            Y = 20,
            FontSize = 10
        });

        var document = PdfReadDocument.Load(stamped);
        var (objects, _) = PdfSyntax.ParseObjects(stamped);
        int pageObjectNumber = document.Pages[0].ObjectNumber;
        var page = Assert.IsType<PdfDictionary>(objects[pageObjectNumber].Value);
        var contents = Assert.IsType<PdfArray>(page.Items["Contents"]);

        Assert.Equal(3, contents.Items.Count);
        foreach (var item in contents.Items) {
            var reference = Assert.IsType<PdfReference>(item);
            Assert.IsType<PdfStream>(objects[reference.ObjectNumber].Value);
        }
    }

    [Fact]
    public void WatermarkText_AddsDefaultWatermarkToEveryPage() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.WatermarkText(source, "DRAFT");

        using var pdf = PdfDocument.Open(new MemoryStream(stamped));
        Assert.Equal(2, pdf.NumberOfPages);

        var read = PdfReadDocument.Load(stamped);
        Assert.Contains("DRAFT", Normalize(read.Pages[0].ExtractText()));
        Assert.Contains("DRAFT", Normalize(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void WatermarkText_CentersUsingStandardFontGlyphWidths() {
        byte[] stamped = PdfStamper.WatermarkText(BuildTwoPagePdf(), "WWWW", new PdfTextStampOptions {
            Font = PdfStandardFont.TimesRoman,
            FontSize = 10,
            RotationDegrees = 0
        });

        string stampContent = FindContentStreamContaining(stamped, "<57575757> Tj");

        Assert.Matches(@"1 0 -?0 1 278\.\d+ 421 Tm", stampContent);
    }

    [Fact]
    public void StampText_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildTwoPagePdf());

        byte[] stamped = PdfStamper.StampText(stream, "STREAM-STAMP", new PdfTextStampOptions {
            PageNumbers = new[] { 1 },
            X = 72,
            Y = 700,
            FontSize = 14
        });

        var read = PdfReadDocument.Load(stamped);
        Assert.Contains("STREAM-STAMP", Normalize(read.Pages[0].ExtractText()));
        Assert.DoesNotContain("STREAM-STAMP", Normalize(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void WatermarkText_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildTwoPagePdf());

        byte[] stamped = PdfStamper.WatermarkText(stream, "STREAM-DRAFT");

        string text = Normalize(PdfReadDocument.Load(stamped).ExtractText());
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
            string text = Normalize(PdfReadDocument.Load(outputPath).ExtractText());
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

            var stampedRead = PdfReadDocument.Load(textStamped);
            Assert.Contains("PATH-BYTES-STAMP", Normalize(stampedRead.Pages[0].ExtractText()));
            Assert.DoesNotContain("PATH-BYTES-STAMP", Normalize(stampedRead.Pages[1].ExtractText()));

            byte[] textWatermarked = PdfStamper.WatermarkTextToBytes(inputPath, "PATH-BYTES-DRAFT");
            Assert.Contains("PATH-BYTES-DRAFT", Normalize(PdfReadDocument.Load(textWatermarked).ExtractText()));

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
            var stampedRead = PdfReadDocument.Load(textStamped);
            Assert.Contains("PATH-STREAM-STAMP", Normalize(stampedRead.Pages[0].ExtractText()));
            Assert.DoesNotContain("PATH-STREAM-STAMP", Normalize(stampedRead.Pages[1].ExtractText()));

            using var watermarkOutput = CreateOutputStream(out int watermarkPrefixLength);
            PdfStamper.WatermarkText(inputPath, watermarkOutput, "PATH-STREAM-DRAFT");

            byte[] textWatermarked = GetOutputPayload(watermarkOutput, watermarkPrefixLength);
            Assert.Contains("PATH-STREAM-DRAFT", Normalize(PdfReadDocument.Load(textWatermarked).ExtractText()));

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
            var read = PdfReadDocument.Load(outputPath);
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
        var read = PdfReadDocument.Load(stamped);
        Assert.Contains("OUTPUT-STAMP", Normalize(read.Pages[0].ExtractText()));
        Assert.DoesNotContain("OUTPUT-STAMP", Normalize(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void WatermarkText_WritesStreamInputToOutputStreamAtCurrentPosition() {
        using var stream = CreatePrefixedStream(BuildTwoPagePdf());
        using var output = CreateOutputStream(out int prefixLength);

        PdfStamper.WatermarkText(stream, output, "OUTPUT-DRAFT");

        byte[] stamped = GetOutputPayload(output, prefixLength);
        string text = Normalize(PdfReadDocument.Load(stamped).ExtractText());
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
            string text = Normalize(PdfReadDocument.Load(outputPath).ExtractText());
            Assert.Contains("STREAM-PATH-DRAFT", text);
            Assert.Contains("Firstpagebody", text);
            Assert.Contains("Secondpagebody", text);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void StampImage_AddsImageXObjectToSelectedPageAndPreservesOriginalContent() {
        byte[] source = BuildTwoPagePdf();
        byte[] image = CreateMinimalRgbPng();

        byte[] stamped = PdfStamper.StampImage(source, image, new PdfImageStampOptions {
            PageNumbers = new[] { 2 },
            X = 72,
            Y = 650,
            Width = 24,
            Height = 24
        });

        using var pdf = PdfDocument.Open(new MemoryStream(stamped));
        Assert.Equal(2, pdf.NumberOfPages);

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
        Assert.Contains(" Do", pdfContent);

        string text = Normalize(PdfReadDocument.Load(stamped).ExtractText());
        Assert.Contains("Firstpagebody", text);
        Assert.Contains("Secondpagebody", text);
    }

    [Fact]
    public void StampImage_UsesInclusivePageRangeSelection() {
        byte[] source = BuildTwoPagePdf();
        byte[] image = CreateMinimalRgbPng();

        byte[] stamped = PdfStamper.StampImage(source, image, new PdfImageStampOptions {
                X = 24,
                Y = 24,
                Width = 8,
                Height = 8
            }
            .UsePageRange(2, 2));

        Assert.DoesNotContain("/OIMOStampIm1 Do", string.Join("\n", GetPageContentStreams(stamped, 1)));
        Assert.Contains("/OIMOStampIm1 Do", string.Join("\n", GetPageContentStreams(stamped, 2)));
    }

    [Fact]
    public void StampImage_UsesInclusivePageRangeListSelectionAndDeduplicatesOverlap() {
        byte[] source = BuildTwoPagePdf();
        byte[] image = CreateMinimalRgbPng();

        byte[] stamped = PdfStamper.StampImage(source, image, new PdfImageStampOptions {
            Width = 12,
            Height = 12
        }.UsePageRanges(PdfPageRange.ParseMany("1-2,2")));

        string firstPageStreams = string.Join("\n", GetPageContentStreams(stamped, 1));
        string secondPageStreams = string.Join("\n", GetPageContentStreams(stamped, 2));
        Assert.Equal(1, CountOccurrences(firstPageStreams, "/OIMOStampIm1 Do"));
        Assert.Equal(1, CountOccurrences(secondPageStreams, "/OIMOStampIm1 Do"));
    }

    [Fact]
    public void StampImage_WithRgbaPng_PreservesSoftMaskImageObject() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampImage(source, CreateMinimalRgbaPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            X = 72,
            Y = 650,
            Width = 24,
            Height = 24
        });

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
    }

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

        string text = Normalize(PdfReadDocument.Load(stamped).ExtractText());
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

        string text = Normalize(PdfReadDocument.Load(stamped).ExtractText());
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

        string text = Normalize(PdfReadDocument.Load(stamped).ExtractText());
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

        using var pdf = PdfDocument.Open(new MemoryStream(stamped));
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

    [Fact]
    public void StampText_WritesExpectedPlacementOperators() {
        byte[] stamped = PdfStamper.StampText(BuildTwoPagePdf(), "VISUAL", new PdfTextStampOptions {
            PageNumbers = new[] { 1 },
            X = 100,
            Y = 620,
            FontSize = 18,
            RotationDegrees = 30,
            Color = PdfColor.FromRgb(255, 0, 0)
        });

        string stampContent = FindContentStreamContaining(stamped, "<56495355414C> Tj");
        Assert.Contains("1 0 0 rg", stampContent);
        Assert.Contains("/OIMOStampF1 18 Tf", stampContent);
        Assert.Contains("0.866 0.5 -0.5 0.866 100 620 Tm", stampContent);
        Assert.Contains("<56495355414C> Tj", stampContent);
    }

    [Fact]
    public void StampImage_WritesExpectedPlacementOperators() {
        byte[] stamped = PdfStamper.StampImage(BuildTwoPagePdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            X = 72,
            Y = 650,
            Width = 24,
            Height = 16,
            RotationDegrees = 90
        });

        string stampContent = FindContentStreamContaining(stamped, "/OIMOStampIm1 Do");
        Assert.Contains("0 24 -16 0 72 650 cm", stampContent);
        Assert.Contains("/OIMOStampIm1 Do", stampContent);
    }

    [Fact]
    public void StampAndWatermark_RespectContentLayeringOrder() {
        byte[] stamped = PdfStamper.StampImage(BuildTwoPagePdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            Width = 24,
            Height = 24
        });

        IReadOnlyList<string> stampedStreams = GetPageContentStreams(stamped, 1);
        Assert.True(stampedStreams.Count >= 2);
        Assert.Contains("/OIMOStampIm1 Do", stampedStreams[stampedStreams.Count - 1]);

        byte[] watermarked = PdfStamper.WatermarkImage(BuildTwoPagePdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
            Width = 32,
            Height = 32
        });

        IReadOnlyList<string> watermarkedStreams = GetPageContentStreams(watermarked, 1);
        Assert.True(watermarkedStreams.Count >= 2);
        Assert.Contains("/OIMOStampIm1 Do", watermarkedStreams[0]);
    }

    [Fact]
    public void TextStampOptions_SnapshotPageNumbersAndRejectInvalidValues() {
        var pageNumbers = new[] { 1, 2 };
        var options = new PdfTextStampOptions {
            PageNumbers = pageNumbers
        };

        pageNumbers[0] = 99;
        int[] readback = options.PageNumbers!;
        readback[1] = 88;

        Assert.Equal(new[] { 1, 2 }, options.PageNumbers);

        var ranged = new PdfTextStampOptions();
        Assert.Same(ranged, ranged.UsePageRange(2, 2));
        Assert.Equal(new[] { 2 }, ranged.PageNumbers);
        int[] rangedReadback = ranged.PageNumbers!;
        rangedReadback[0] = 99;
        Assert.Equal(new[] { 2 }, ranged.PageNumbers);

        var modelRange = new PdfTextStampOptions();
        Assert.Same(modelRange, modelRange.UsePageRange(PdfPageRange.From(1, 2)));
        Assert.Equal(new[] { 1, 2 }, modelRange.PageNumbers);

        var rangeList = new PdfTextStampOptions();
        Assert.Same(rangeList, rangeList.UsePageRanges(PdfPageRange.ParseMany("1-2,2")));
        Assert.Equal(new[] { 1, 2 }, rangeList.PageNumbers);
        int[] rangeListReadback = rangeList.PageNumbers!;
        rangeListReadback[0] = 99;
        Assert.Equal(new[] { 1, 2 }, rangeList.PageNumbers);

        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { PageNumbers = new[] { 0 } });
        Assert.Throws<ArgumentException>(() => new PdfTextStampOptions { PageNumbers = new[] { 1, 1 } });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions().UsePageRange(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions().UsePageRange(2, 1));
        Assert.Throws<ArgumentNullException>(() => new PdfTextStampOptions().UsePageRanges(null!));
        Assert.Throws<ArgumentException>(() => new PdfTextStampOptions().UsePageRanges(Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { X = double.NaN });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { Y = double.PositiveInfinity });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { FontSize = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { FontSize = double.NaN });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { RotationDegrees = double.NegativeInfinity });
        var fontException = Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { Font = (PdfStandardFont)99 });
        Assert.Equal("Font", fontException.ParamName);
        Assert.Contains("Text stamp font must be one of the supported standard PDF fonts.", fontException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageStampOptions_SnapshotPageNumbersAndRejectInvalidValues() {
        var pageNumbers = new[] { 1, 2 };
        var options = new PdfImageStampOptions {
            PageNumbers = pageNumbers
        };

        pageNumbers[0] = 99;
        int[] readback = options.PageNumbers!;
        readback[1] = 88;

        Assert.Equal(new[] { 1, 2 }, options.PageNumbers);

        var ranged = new PdfImageStampOptions();
        Assert.Same(ranged, ranged.UsePageRange(2, 2));
        Assert.Equal(new[] { 2 }, ranged.PageNumbers);
        int[] rangedReadback = ranged.PageNumbers!;
        rangedReadback[0] = 99;
        Assert.Equal(new[] { 2 }, ranged.PageNumbers);

        var modelRange = new PdfImageStampOptions();
        Assert.Same(modelRange, modelRange.UsePageRange(PdfPageRange.From(1, 2)));
        Assert.Equal(new[] { 1, 2 }, modelRange.PageNumbers);

        var rangeList = new PdfImageStampOptions();
        Assert.Same(rangeList, rangeList.UsePageRanges(PdfPageRange.ParseMany("1-2,2")));
        Assert.Equal(new[] { 1, 2 }, rangeList.PageNumbers);
        int[] rangeListReadback = rangeList.PageNumbers!;
        rangeListReadback[0] = 99;
        Assert.Equal(new[] { 1, 2 }, rangeList.PageNumbers);

        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { PageNumbers = new[] { 0 } });
        Assert.Throws<ArgumentException>(() => new PdfImageStampOptions { PageNumbers = new[] { 1, 1 } });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions().UsePageRange(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions().UsePageRange(2, 1));
        Assert.Throws<ArgumentNullException>(() => new PdfImageStampOptions().UsePageRanges(null!));
        Assert.Throws<ArgumentException>(() => new PdfImageStampOptions().UsePageRanges(Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { X = double.NaN });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Y = double.PositiveInfinity });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Width = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Width = double.NaN });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Height = -1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Height = double.PositiveInfinity });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { RotationDegrees = double.NegativeInfinity });
    }

    [Fact]
    public void Stamper_RejectsInvalidInputs() {
        byte[] source = BuildTwoPagePdf();

        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText((byte[])null!, "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText((Stream)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(new WriteOnlyStream(), "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText(source, null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(source, string.Empty));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampText(source, "x", new PdfTextStampOptions { FontSize = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampText(source, "x", new PdfTextStampOptions { PageNumbers = new[] { 0 } }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampText(source, "x", new PdfTextStampOptions { PageNumbers = new[] { 3 } }));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(source, "x", new PdfTextStampOptions { PageNumbers = new[] { 1, 1 } }));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText(source, null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(source, new ReadOnlyStream(), "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText(new MemoryStream(source), (string)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(new MemoryStream(source), " ", "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampTextToBytes(" ", "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampText("input.pdf", (Stream)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText("missing.pdf", new ReadOnlyStream(), "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(" ", new MemoryStream(), "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(" ", "out.pdf", "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText("missing.pdf", " ", "x"));

        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText((byte[])null!, "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText((Stream)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(new WriteOnlyStream(), "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText(source, null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(source, string.Empty));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText(source, null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(source, new ReadOnlyStream(), "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText(new MemoryStream(source), (string)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(new MemoryStream(source), " ", "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkTextToBytes(" ", "x"));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkText("input.pdf", (Stream)null!, "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText("missing.pdf", new ReadOnlyStream(), "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(" ", new MemoryStream(), "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText(" ", "out.pdf", "x"));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkText("missing.pdf", " ", "x"));

        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage((byte[])null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage((Stream)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(new WriteOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage(source, (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(source, new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage(new MemoryStream(source), (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(new MemoryStream(source), new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage(source, (byte[])null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(source, Array.Empty<byte>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampImage(source, CreateMinimalRgbPng(), new PdfImageStampOptions { Width = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampImage(source, CreateMinimalRgbPng(), new PdfImageStampOptions { Height = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfStamper.StampImage(source, CreateMinimalRgbPng(), new PdfImageStampOptions { PageNumbers = new[] { 0 } }));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(source, CreateMinimalRgbPng(), new PdfImageStampOptions { PageNumbers = new[] { 1, 1 } }));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage(source, null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(source, new ReadOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage(new MemoryStream(source), (string)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(new MemoryStream(source), " ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(new MemoryStream(source), " ", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImageToBytes(" ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImageToBytes(" ", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage("input.pdf", (Stream)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage("missing.pdf", new ReadOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(" ", new MemoryStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.StampImage("input.pdf", (Stream)null!, new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage("missing.pdf", new ReadOnlyStream(), new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(" ", new MemoryStream(), new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(" ", "out.pdf", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage(" ", "out.pdf", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage("missing.pdf", " ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.StampImage("missing.pdf", " ", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage((byte[])null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage((Stream)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(new WriteOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage(source, (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(source, new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage(new MemoryStream(source), (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(new MemoryStream(source), new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage(source, (byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage(source, null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(source, new ReadOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage(new MemoryStream(source), (string)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(new MemoryStream(source), " ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(new MemoryStream(source), " ", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImageToBytes(" ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImageToBytes(" ", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage("input.pdf", (Stream)null!, CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage("missing.pdf", new ReadOnlyStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(" ", new MemoryStream(), CreateMinimalRgbPng()));
        Assert.Throws<ArgumentNullException>(() => PdfStamper.WatermarkImage("input.pdf", (Stream)null!, new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage("missing.pdf", new ReadOnlyStream(), new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(" ", new MemoryStream(), new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(" ", "out.pdf", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage(" ", "out.pdf", new MemoryStream(CreateMinimalRgbPng())));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage("missing.pdf", " ", CreateMinimalRgbPng()));
        Assert.Throws<ArgumentException>(() => PdfStamper.WatermarkImage("missing.pdf", " ", new MemoryStream(CreateMinimalRgbPng())));
    }

    [Fact]
    public void Stamper_PathOutputsRejectDirectoryTargets() {
        byte[] source = BuildTwoPagePdf();
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-stamp-output-path-" + Guid.NewGuid().ToString("N"));

        try {
            Directory.CreateDirectory(directory);

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfStamper.WatermarkText(new MemoryStream(source), directory, "DRAFT"));

            Assert.Equal("outputPath", exception.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", exception.Message, StringComparison.Ordinal);

            var pathInputException = Assert.Throws<ArgumentException>(() =>
                PdfStamper.StampText("missing.pdf", directory, "DRAFT"));

            Assert.Equal("outputPath", pathInputException.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", pathInputException.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static byte[] BuildTwoPagePdf() {
        var doc = PdfDoc.Create()
            .Meta(title: "Stamp sample", author: "OfficeIMO");

        doc.Compose(compose => {
            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("First page body"))));
            });

            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Second page body"))));
            });
        });

        return doc.ToBytes();
    }

    private static byte[] BuildIndirectContentsArrayPdf() {
        string first = "BT /F1 12 Tf 20 80 Td (First) Tj ET";
        string second = "BT /F1 12 Tf 20 60 Td (Second) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 8 0 R /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> >>",
            "endobj",
            "4 0 obj",
            $"<< /Length {first.Length} >>",
            "stream",
            first,
            "endstream",
            "endobj",
            "5 0 obj",
            $"<< /Length {second.Length} >>",
            "stream",
            second,
            "endstream",
            "endobj",
            "8 0 obj",
            "[4 0 R 5 0 R]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static string Normalize(string text) {
        return text.Replace(" ", string.Empty);
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static MemoryStream CreatePrefixedStream(byte[] pdf) {
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        var stream = new MemoryStream();
        stream.Write(prefix, 0, prefix.Length);
        stream.Write(pdf, 0, pdf.Length);
        stream.Position = prefix.Length;
        return stream;
    }

    private static MemoryStream CreateOutputStream(out int prefixLength) {
        byte[] prefix = Encoding.ASCII.GetBytes("output-prefix");
        var stream = new MemoryStream();
        stream.Write(prefix, 0, prefix.Length);
        prefixLength = prefix.Length;
        return stream;
    }

    private static byte[] GetOutputPayload(MemoryStream output, int prefixLength) {
        byte[] bytes = output.ToArray();
        Assert.True(bytes.Length > prefixLength);
        Assert.Equal("output-prefix", Encoding.ASCII.GetString(bytes, 0, prefixLength));

        var payload = new byte[bytes.Length - prefixLength];
        Array.Copy(bytes, prefixLength, payload, 0, payload.Length);
        return payload;
    }

    private static string FindContentStreamContaining(byte[] pdf, string marker) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        foreach (var item in objects.Values) {
            if (item.Value is PdfStream stream) {
                string content = DecodeStream(stream.Data);
                if (content.Contains(marker)) {
                    return content;
                }
            }
        }

        throw new InvalidOperationException("Content stream marker was not found: " + marker);
    }

    private static IReadOnlyList<string> GetPageContentStreams(byte[] pdf, int pageNumber) {
        var document = PdfReadDocument.Load(pdf);
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        int pageObjectNumber = document.Pages[pageNumber - 1].ObjectNumber;
        if (!objects.TryGetValue(pageObjectNumber, out var pageObject) || pageObject.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("Page object was not found.");
        }

        if (!pageDictionary.Items.TryGetValue("Contents", out var contents)) {
            throw new InvalidOperationException("Page contents were not found.");
        }

        var streams = new List<string>();
        AppendContentStreams(objects, contents, streams);
        return streams;
    }

    private static void AppendContentStreams(Dictionary<int, PdfIndirectObject> objects, PdfObject contents, List<string> streams) {
        if (contents is PdfReference reference) {
            if (objects.TryGetValue(reference.ObjectNumber, out var indirect) && indirect.Value is PdfStream stream) {
                streams.Add(DecodeStream(stream.Data));
            }

            return;
        }

        if (contents is PdfArray array) {
            foreach (var item in array.Items) {
                AppendContentStreams(objects, item, streams);
            }
        }
    }

    private static string DecodeStream(byte[] data) {
        return Encoding.GetEncoding("ISO-8859-1").GetString(data);
    }

    private sealed class WriteOnlyStream : MemoryStream {
        public override bool CanRead => false;
    }

    private sealed class ReadOnlyStream : MemoryStream {
        public override bool CanWrite => false;
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }

    private static byte[] CreateMinimalRgbaPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 6, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 16,
            73, 68, 65, 84,
            0x78, 0x01, 0x01, 0x05, 0x00, 0xFA, 0xFF, 0x00,
            0xFF, 0x00, 0x00, 0x80, 0x04, 0x81, 0x01, 0x80,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }
}
