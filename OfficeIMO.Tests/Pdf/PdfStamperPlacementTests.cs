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
}
