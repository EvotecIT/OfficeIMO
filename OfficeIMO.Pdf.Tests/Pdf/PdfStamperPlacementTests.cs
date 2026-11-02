using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfStamperTests {
    [Fact]
    public void StampText_AlignsEmbeddedGlyphsInPhysicalPointsOnUserUnitPages() {
        if (!PdfCreateOptions.TryResolveDefaultEmbeddedFontFamily(PdfStandardFont.Helvetica, out _)) return;
        byte[] stamped = PdfStamper.StampText(BuildUserUnitStampPdf(), "Page 1 of 1", new PdfTextStampOptions {
            X = OfficeLength.Points(300),
            Y = OfficeLength.Points(100),
            FontSize = OfficeLength.Points(24),
            Align = PdfAlign.Right,
            EmbedFont = true,
            Font = PdfStandardFont.Helvetica
        });

        Assert.Contains("/FontFile", PdfTestContentText.ReadDecoded(stamped), StringComparison.Ordinal);
        PdfReadPage page = PdfReadDocument.Open(stamped).Pages[0];
        PdfTextSpan span = page.GetTextSpans().First(item => item.Text.Contains("Page"));
        double userUnit = page.GetGeometry().UserUnit ?? 1D;
        Assert.Equal(2D, userUnit);
        Assert.InRange((span.X + span.Advance) * userUnit, 299.8D, 300.2D);
        Assert.InRange(span.FontSize * userUnit, 23.99D, 24.01D);
    }

    [Fact]
    public void StampText_WritesExpectedPlacementOperators() {
        byte[] stamped = PdfStamper.StampText(BuildTwoPagePdf(), "VISUAL", new PdfTextStampOptions {
            PageNumbers = new[] { 1 },
            X = global::OfficeIMO.OfficeLength.Points(100),
            Y = global::OfficeIMO.OfficeLength.Points(620),
            FontSize = global::OfficeIMO.OfficeLength.Points(18),
            RotationDegrees = 30,
            Color = global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0)
        });

        string stampContent = FindContentStreamContaining(stamped, "<56495355414C> Tj");
        Assert.Contains("1 0 0 rg", stampContent);
        Assert.Contains("/OIMOStampF1 18 Tf", stampContent);
        Assert.Contains("0.866 0.5 -0.5 0.866 100 620 Tm", stampContent);
        Assert.Contains("<56495355414C> Tj", stampContent);

        byte[] physical = PdfStamper.StampText(BuildUserUnitStampPdf(), "VISUAL", new PdfTextStampOptions {
            PageNumbers = new[] { 1 },
            X = OfficeLength.Centimeters(100D * 2.54D / 72D),
            Y = OfficeLength.Millimeters(620D * 25.4D / 72D),
            FontSize = OfficeLength.Inches(0.25D),
            RotationDegrees = 30,
            Color = OfficeColor.FromRgb(255, 0, 0)
        });
        IReadOnlyList<string> physicalStreams = GetPageContentStreams(physical, 1);
        string physicalContent = physicalStreams[physicalStreams.Count - 1];
        Assert.Contains("1 0 0 rg", physicalContent);
        Assert.Contains("/OIMOStampF1 9 Tf", physicalContent);
        Assert.Contains("0.866 0.5 -0.5 0.866 50 310 Tm", physicalContent);
        Assert.Contains("<56495355414C> Tj", physicalContent);
        Assert.DoesNotContain("<56495355414C> Tj", string.Join("\n", GetPageContentStreams(physical, 2)));
        PdfReadPage physicalPage = PdfReadDocument.Open(physical).Pages[0];
        Assert.Equal(2D, physicalPage.GetGeometry().UserUnit);
        Assert.Equal((400D, 600D), physicalPage.GetPageSize());
        Assert.Equal(90, physicalPage.GetRotationDegrees());

        byte[] defaults = PdfStamper.StampText(BuildUserUnitStampPdf(), "DEFAULT");
        string defaultContent = FindContentStreamContaining(defaults, "<44454641554C54> Tj");
        Assert.Contains("/OIMOStampF1 12 Tf", defaultContent);
        Assert.Matches(@"1 0 -?0 1 18 18 Tm", defaultContent);

        byte[] centered = PdfStamper.WatermarkText(BuildUserUnitStampPdf(), "WWWW", new PdfTextStampOptions {
            Font = PdfStandardFont.Courier,
            FontSize = OfficeLength.Points(36)
        });
        string centeredContent = FindContentStreamContaining(centered, "<57575757> Tj");
        Assert.Contains("/OIMOStampF1 18 Tf", centeredContent);
        Assert.Matches(@"1 0 -?0 1 198\.4 330 Tm", centeredContent);

        byte[] defaultWatermark = PdfStamper.WatermarkText(BuildUserUnitStampPdf(), "DRAFT");
        string watermarkContent = FindContentStreamContaining(defaultWatermark, "<4452414654> Tj");
        Assert.Contains("/OIMOStampF1 32 Tf", watermarkContent);
        Assert.Matches(@"0\.707 -0\.707 0\.707 0\.707 [0-9.]+ 330 Tm", watermarkContent);
    }

    [Fact]
    public void StampImage_WritesExpectedPlacementOperators() {
        byte[] stamped = PdfStamper.StampImage(BuildTwoPagePdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            X = global::OfficeIMO.OfficeLength.Points(72),
            Y = global::OfficeIMO.OfficeLength.Points(650),
            Width = global::OfficeIMO.OfficeLength.Points(24),
            Height = global::OfficeIMO.OfficeLength.Points(16),
            RotationDegrees = 90
        });

        string stampContent = FindContentStreamContaining(stamped, "/OIMOStampIm1 Do");
        AssertStampImageMatrix(stampContent, 0D, 24D, -16D, 0D, 72D, 650D);
        Assert.Contains("/OIMOStampIm1 Do", stampContent);

        byte[] physical = PdfStamper.StampImage(BuildUserUnitStampPdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            X = OfficeLength.Inches(1),
            Y = OfficeLength.Millimeters(650D * 25.4D / 72D),
            Width = OfficeLength.Centimeters(24D * 2.54D / 72D),
            Height = OfficeLength.Points(16),
            RotationDegrees = 90
        });
        IReadOnlyList<string> physicalStreams = GetPageContentStreams(physical, 1);
        string physicalContent = physicalStreams[physicalStreams.Count - 1];
        AssertStampImageMatrix(physicalContent, 0D, 12D, -8D, 0D, 36D, 325D);
        Assert.Contains("/OIMOStampIm1 Do", physicalContent);
        Assert.DoesNotContain("/OIMOStampIm1 Do", string.Join("\n", GetPageContentStreams(physical, 2)));
        PdfReadPage physicalPage = PdfReadDocument.Open(physical).Pages[0];
        Assert.Equal(2D, physicalPage.GetGeometry().UserUnit);
        Assert.Equal((400D, 600D), physicalPage.GetPageSize());
        Assert.Equal(90, physicalPage.GetRotationDegrees());

        byte[] intrinsic = PdfStamper.StampImage(BuildUserUnitStampPdf(), CreateMinimalRgbPng());
        string intrinsicContent = FindContentStreamContaining(intrinsic, "/OIMOStampIm1 Do");
        Assert.Contains("0.5 0 0 0.5 18 18 cm", intrinsicContent);
    }

    [Fact]
    public void StampAndWatermark_RespectContentLayeringOrder() {
        byte[] stamped = PdfStamper.StampImage(BuildTwoPagePdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
            PageNumbers = new[] { 1 },
            Width = global::OfficeIMO.OfficeLength.Points(24),
            Height = global::OfficeIMO.OfficeLength.Points(24)
        });

        IReadOnlyList<string> stampedStreams = GetPageContentStreams(stamped, 1);
        Assert.True(stampedStreams.Count >= 2);
        Assert.Contains("/OIMOStampIm1 Do", stampedStreams[stampedStreams.Count - 1]);

        byte[] watermarked = PdfStamper.WatermarkImage(BuildTwoPagePdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
            Width = global::OfficeIMO.OfficeLength.Points(32),
            Height = global::OfficeIMO.OfficeLength.Points(32)
        });

        IReadOnlyList<string> watermarkedStreams = GetPageContentStreams(watermarked, 1);
        Assert.True(watermarkedStreams.Count >= 2);
        Assert.Contains("/OIMOStampIm1 Do", watermarkedStreams[0]);

        byte[] physicalWatermark = PdfStamper.WatermarkImage(BuildUserUnitStampPdf(), CreateMinimalRgbPng(), new PdfImageStampOptions {
            Width = OfficeLength.Points(32),
            Height = OfficeLength.Inches(32D / 72D)
        });
        IReadOnlyList<string> physicalStreams = GetPageContentStreams(physicalWatermark, 1);
        Assert.True(physicalStreams.Count >= 2);
        Assert.Contains("16 0 0 16 212 322 cm", physicalStreams[0]);
        Assert.Contains("/OIMOStampIm1 Do", physicalStreams[0]);

        byte[] intrinsicWatermark = PdfStamper.WatermarkImage(BuildUserUnitStampPdf(), CreateMinimalRgbPng());
        IReadOnlyList<string> intrinsicStreams = GetPageContentStreams(intrinsicWatermark, 1);
        Assert.Contains("0.5 0 0 0.5 219.75 329.75 cm", intrinsicStreams[0]);
        Assert.Contains("/OIMOStampIm1 Do", intrinsicStreams[0]);
    }

    [Theory]
    [InlineData(false, false, 10000D)]
    [InlineData(false, true, 60000D)]
    [InlineData(true, false, 10000D)]
    [InlineData(true, true, 60000D)]
    public void StampGeometry_PreservesPhysicalDimensionsOnLargeUserUnitPages(bool image, bool watermark, double userUnit) {
        byte[] source = BuildUserUnitStampPdf(userUnit);
        byte[] stamped;
        if (image) {
            var options = new PdfImageStampOptions {
                X = OfficeLength.Points(-36.125), Y = OfficeLength.Points(72.375),
                Width = OfficeLength.Points(24.125), Height = OfficeLength.Points(16.25)
            };
            stamped = watermark ? PdfStamper.WatermarkImage(source, CreateMinimalRgbPng(), options)
                : PdfStamper.StampImage(source, CreateMinimalRgbPng(), options);
        } else {
            var options = new PdfTextStampOptions {
                X = OfficeLength.Points(-36.125), Y = OfficeLength.Points(72.375), FontSize = OfficeLength.Points(24.125)
            };
            stamped = watermark ? PdfStamper.WatermarkText(source, "PRECISE", options)
                : PdfStamper.StampText(source, "PRECISE", options);
        }
        string content = FindContentStreamContaining(stamped, image ? "/OIMOStampIm1 Do" : "<50524543495345> Tj");
        Match matrix = Regex.Match(content, @"([-0-9.]+) ([-0-9.]+) ([-0-9.]+) ([-0-9.]+) ([-0-9.]+) ([-0-9.]+) " + (image ? "cm" : "Tm"));
        Assert.True(matrix.Success, content);
        double Physical(int group) => double.Parse(matrix.Groups[group].Value, CultureInfo.InvariantCulture) * userUnit;
        Assert.InRange(Math.Abs(Physical(5) + 36.125), 0D, 1D / 12700D);
        Assert.InRange(Math.Abs(Physical(6) - 72.375), 0D, 1D / 12700D);
        if (image) {
            Assert.InRange(Math.Abs(Physical(1) - 24.125), 0D, 1D / 12700D);
            Assert.InRange(Math.Abs(Physical(4) - 16.25), 0D, 1D / 12700D);
        } else {
            Match font = Regex.Match(content, @"/OIMOStampF1 ([-0-9.]+) Tf");
            Assert.True(font.Success, content);
            double physicalSize = double.Parse(font.Groups[1].Value, CultureInfo.InvariantCulture) * userUnit;
            Assert.InRange(Math.Abs(physicalSize - 24.125), 0D, 1D / 12700D);
        }
        Assert.Equal(userUnit, PdfReadDocument.Open(stamped).Pages[0].GetGeometry().UserUnit);
    }

    private static void AssertStampImageMatrix(string content, params double[] expected) {
        Match matrix = Regex.Match(content, @"([-0-9.]+) ([-0-9.]+) ([-0-9.]+) ([-0-9.]+) ([-0-9.]+) ([-0-9.]+) cm");
        Assert.True(matrix.Success, content);
        for (int index = 0; index < expected.Length; index++) {
            double actual = double.Parse(matrix.Groups[index + 1].Value, CultureInfo.InvariantCulture);
            Assert.InRange(Math.Abs(actual - expected[index]), 0D, 1D / 12700D);
        }
    }

    private static byte[] BuildUserUnitStampPdf(double userUnit = 2D) =>
        PdfDocumentObjectGraphRewriter.Rewrite(BuildTwoPagePdf(), sourceReadOptions: null, outputEncryption: null, mutateObjectGraph: (objects, security) => {
            foreach (PdfIndirectObject indirect in objects.Values) {
                if (indirect.Value is not PdfDictionary page ||
                    !page.Items.TryGetValue("Type", out PdfObject? type) || type is not PdfName name || name.Name != "Page") continue;
                var crop = new PdfArray();
                foreach (double coordinate in new[] { 20D, 30D, 420D, 630D }) crop.Items.Add(new PdfNumber(coordinate));
                page.Items["CropBox"] = crop;
                page.Items["UserUnit"] = new PdfNumber(userUnit);
                page.Items["Rotate"] = new PdfNumber(90D);
            }
            return security.InfoObjectNumber;
        });
}
