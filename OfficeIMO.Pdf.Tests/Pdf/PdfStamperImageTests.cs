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

        using var pdf = PdfPigDocument.Open(new MemoryStream(stamped));
        Assert.Equal(2, pdf.NumberOfPages);

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
        Assert.Contains(" Do", pdfContent);

        string text = Normalize(PdfReadDocument.Open(stamped).ExtractText());
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
    public void StampImage_WithIndexedColorPngTransparency_PreservesSoftMaskImageObject() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampImage(source, CreateMinimalIndexedColorPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            X = 72,
            Y = 650,
            Width = 24,
            Height = 12
        });

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceRGB", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
    }

    [Fact]
    public void StampImage_WithRgbPngTransparency_PreservesSoftMaskImageObject() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampImage(source, CreateMinimalRgbTransparencyPng(), new PdfImageStampOptions {
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
    public void StampImage_WithPackedGrayscalePngTransparency_PreservesSoftMaskImageObject() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampImage(source, CreateMinimalPackedGrayscalePng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            X = 72,
            Y = 650,
            Width = 24,
            Height = 12
        });

        string pdfContent = Encoding.ASCII.GetString(stamped);
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/SMask", pdfContent);
        Assert.Contains("/ColorSpace /DeviceGray", pdfContent);
        Assert.Contains("/OIMOStampIm", pdfContent);
    }
}
