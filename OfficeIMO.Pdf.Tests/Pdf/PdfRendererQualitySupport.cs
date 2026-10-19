using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Shared corpus access and image metrics for managed-rasterizer fidelity tests.</summary>
internal static class PdfRendererQualitySupport {
    internal const double Dpi = 150D;

    internal static string CorpusPath(string fileName) =>
        Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), "Pdf", "RendererQuality", fileName);

    internal static OfficeImageExportResult Export(string pdfFileName, double dpi = Dpi) {
        PdfReadDocument document = PdfReadDocument.Open(File.ReadAllBytes(CorpusPath(pdfFileName)));
        return document.Pages[0].ExportImage(OfficeImageExportFormat.Png, new PdfImageExportOptions { TargetDpi = dpi });
    }

    internal static OfficeRasterImage Render(string pdfFileName, double dpi = Dpi) => Decode(Export(pdfFileName, dpi).Bytes);

    internal static OfficeRasterImage Baseline(string pngFileName) => Decode(File.ReadAllBytes(CorpusPath(pngFileName)));

    internal static OfficeRasterImage Decode(byte[] png) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(png, out OfficeRasterImage? image));
        return image!;
    }

    /// <summary>Peak signal-to-noise ratio over the RGB channels of the shared area, composited on white.</summary>
    internal static double Psnr(OfficeRasterImage actual, OfficeRasterImage expected) {
        int width = Math.Min(actual.Width, expected.Width), height = Math.Min(actual.Height, expected.Height);
        byte[] a = actual.GetPixels(), b = expected.GetPixels();
        double sum = 0D;
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                int ia = (y * actual.Width + x) * 4, ib = (y * expected.Width + x) * 4;
                for (int channel = 0; channel < 3; channel++) {
                    double difference = OnWhite(a, ia, channel) - OnWhite(b, ib, channel);
                    sum += difference * difference;
                }
            }
        }
        double mse = sum / (width * (double)height * 3D);
        return mse == 0D ? 99D : 10D * Math.Log10(255D * 255D / mse);
    }

    /// <summary>Mean structural similarity of luma over 8x8 windows with a stride of four pixels.</summary>
    internal static double Ssim(OfficeRasterImage actual, OfficeRasterImage expected) {
        int width = Math.Min(actual.Width, expected.Width), height = Math.Min(actual.Height, expected.Height);
        double[] la = Luma(actual, width, height), lb = Luma(expected, width, height);
        const double c1 = 0.01D * 255D * 0.01D * 255D, c2 = 0.03D * 255D * 0.03D * 255D;
        double total = 0D;
        int windows = 0;
        for (int y = 0; y + 8 <= height; y += 4) {
            for (int x = 0; x + 8 <= width; x += 4) {
                double ma = 0D, mb = 0D;
                for (int j = 0; j < 8; j++) for (int i = 0; i < 8; i++) { ma += la[(y + j) * width + x + i]; mb += lb[(y + j) * width + x + i]; }
                ma /= 64D; mb /= 64D;
                double va = 0D, vb = 0D, cov = 0D;
                for (int j = 0; j < 8; j++) {
                    for (int i = 0; i < 8; i++) {
                        double da = la[(y + j) * width + x + i] - ma, db = lb[(y + j) * width + x + i] - mb;
                        va += da * da; vb += db * db; cov += da * db;
                    }
                }
                va /= 64D; vb /= 64D; cov /= 64D;
                total += (2D * ma * mb + c1) * (2D * cov + c2) / ((ma * ma + mb * mb + c1) * (va + vb + c2));
                windows++;
            }
        }
        return windows == 0 ? 1D : total / windows;
    }

    /// <summary>Ink extents of every horizontal band of the expected image, measured in both images.</summary>
    internal static IReadOnlyList<(int Top, int Bottom, InkBox Expected, InkBox Actual)> CompareInkBands(
        OfficeRasterImage actual, OfficeRasterImage expected, int threshold = 128) {
        int width = Math.Min(actual.Width, expected.Width), height = Math.Min(actual.Height, expected.Height);
        var bands = new List<(int, int, InkBox, InkBox)>();
        int y = 0;
        while (y < height) {
            while (y < height && !RowHasInk(expected, y, width, threshold)) y++;
            if (y >= height) break;
            int top = y;
            int gap = 0;
            while (y < height && gap < 4) {
                gap = RowHasInk(expected, y, width, threshold) ? 0 : gap + 1;
                y++;
            }
            int bottom = y - gap;
            bands.Add((top, bottom, Ink(expected, top - 2, bottom + 2, width, threshold), Ink(actual, top - 2, bottom + 2, width, threshold)));
        }
        return bands;
    }

    internal readonly struct InkBox {
        internal InkBox(int left, int top, int right, int bottom, double darkness) {
            Left = left; Top = top; Right = right; Bottom = bottom; Darkness = darkness;
        }
        internal int Left { get; }
        internal int Top { get; }
        internal int Right { get; }
        internal int Bottom { get; }
        internal double Darkness { get; }
        public override string ToString() => $"[{Left},{Top}-{Right},{Bottom} ink={Darkness:0}]";
    }

    private static InkBox Ink(OfficeRasterImage image, int top, int bottom, int width, int threshold) {
        int left = int.MaxValue, right = -1, inkTop = int.MaxValue, inkBottom = -1;
        double darkness = 0D;
        for (int y = Math.Max(0, top); y < Math.Min(image.Height, bottom); y++) {
            for (int x = 0; x < width; x++) {
                double luma = LumaAt(image, x, y);
                darkness += 255D - luma;
                if (luma >= threshold) continue;
                left = Math.Min(left, x); right = Math.Max(right, x);
                inkTop = Math.Min(inkTop, y); inkBottom = Math.Max(inkBottom, y);
            }
        }
        return new InkBox(left, inkTop, right, inkBottom, darkness / 255D);
    }

    private static bool RowHasInk(OfficeRasterImage image, int y, int width, int threshold) {
        for (int x = 0; x < width; x++) if (LumaAt(image, x, y) < threshold) return true;
        return false;
    }

    private static double[] Luma(OfficeRasterImage image, int width, int height) {
        var result = new double[width * height];
        for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) result[y * width + x] = LumaAt(image, x, y);
        return result;
    }

    private static double LumaAt(OfficeRasterImage image, int x, int y) {
        OfficeColor color = image.GetPixel(x, y);
        double alpha = color.A / 255D;
        double r = color.R * alpha + 255D * (1D - alpha), g = color.G * alpha + 255D * (1D - alpha), b = color.B * alpha + 255D * (1D - alpha);
        return 0.299D * r + 0.587D * g + 0.114D * b;
    }

    private static double OnWhite(byte[] pixels, int offset, int channel) {
        double alpha = pixels[offset + 3] / 255D;
        return pixels[offset + channel] * alpha + 255D * (1D - alpha);
    }
}
