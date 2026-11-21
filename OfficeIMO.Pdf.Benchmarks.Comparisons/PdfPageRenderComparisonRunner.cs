using System.Diagnostics;
using System.Globalization;
using System.Text;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Renders the same PDF pages to PNG with OfficeIMO and with PDFium (through Docnet.Core) and compares the results.
/// Both lanes produce complete PNG bytes: PDFium's pixels are composited on white and encoded with the same OfficeIMO
/// PNG encoder, so the timed work matches. PDFium's render-only time is reported separately. Every page must have
/// matching dimensions and visible ink in both engines before its timing is reported. The renderers antialias text
/// differently, so pixel agreement is a similarity measure rather than an equality gate.
/// </summary>
internal static class PdfPageRenderComparisonRunner {
    private static readonly OfficeRasterEncodingOptions EncodingOptions = new() {
        WriteResolutionMetadata = false,
        Png = new OfficePngEncodeOptions { WritePhysicalResolution = false }
    };

    internal static int Run(string[] args) {
        int pages = ReadInt(args, "--pages", 4);
        int samples = ReadInt(args, "--samples", 5);
        var inputs = new List<(string Name, byte[] Bytes)>();
        PdfBenchmarkScenario scenario = PdfBenchmarkScenario.Get(PdfBenchmarkScale.Medium);
        foreach (PdfBenchmarkProducer producer in Enum.GetValues<PdfBenchmarkProducer>()) {
            inputs.Add((producer + " report", PdfPreparedReadInputs.Resolve(producer, scenario)));
        }
        for (int index = 1; index < args.Length - 1; index++) {
            if (string.Equals(args[index], "--input", StringComparison.OrdinalIgnoreCase)) {
                string path = Path.GetFullPath(args[++index]);
                inputs.Add((Path.GetFileName(path), File.ReadAllBytes(path)));
            }
        }

        var report = new StringBuilder();
        report.AppendLine($"PDFium through Docnet.Core {typeof(DocLib).Assembly.GetName().Version}; {pages} pages per input; minimum of {samples} timed samples.");
        report.AppendLine();
        report.AppendLine("| Input | Scale | Pages | OfficeIMO ms/page | OfficeIMO KiB/page | PDFium ms/page (render + PNG) | PDFium ms/page (render) | Mean luma difference | Pixels within 32 |");
        report.AppendLine("| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |");
        foreach ((string name, byte[] pdf) in inputs) {
            foreach (double scale in new[] { 1D, 2D }) {
                int count = Math.Min(pages, PdfDocument.Load(new MemoryStream(pdf, writable: false)).Pages.Count);
                byte[][] officeImo = RenderOfficeImo(pdf, count, scale);
                byte[][] pdfium = RenderPdfium(pdf, count, scale, encode: true, out _);
                (double meanDifference, double withinTolerance) = Compare(name, scale, officeImo, pdfium);

                double officeImoTime = double.MaxValue, pdfiumTime = double.MaxValue, pdfiumRenderTime = double.MaxValue;
                long officeImoAllocated = long.MaxValue;
                for (int sample = 0; sample < samples; sample++) {
                    GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
                    long before = GC.GetAllocatedBytesForCurrentThread();
                    var stopwatch = Stopwatch.StartNew();
                    RenderOfficeImo(pdf, count, scale);
                    officeImoTime = Math.Min(officeImoTime, stopwatch.Elapsed.TotalMilliseconds);
                    officeImoAllocated = Math.Min(officeImoAllocated, GC.GetAllocatedBytesForCurrentThread() - before);

                    stopwatch.Restart();
                    RenderPdfium(pdf, count, scale, encode: true, out _);
                    pdfiumTime = Math.Min(pdfiumTime, stopwatch.Elapsed.TotalMilliseconds);
                    RenderPdfium(pdf, count, scale, encode: false, out double renderOnly);
                    pdfiumRenderTime = Math.Min(pdfiumRenderTime, renderOnly);
                }
                report.AppendLine(string.Create(CultureInfo.InvariantCulture,
                    $"| {name} | {scale:0.#}x | {count} | {officeImoTime / count:F1} | {officeImoAllocated / 1024D / count:F0} | {pdfiumTime / count:F1} | {pdfiumRenderTime / count:F1} | {meanDifference:F2} | {withinTolerance:P2} |"));
            }
        }
        Console.Write(report.ToString());
        string? output = ReadOption(args, "--output");
        if (output != null) File.WriteAllText(Path.GetFullPath(output), report.ToString());
        return 0;
    }

    private static byte[][] RenderOfficeImo(byte[] pdf, int count, double scale) {
        IReadOnlyList<PdfPageRenderResult> results = PdfDocument.Load(new MemoryStream(pdf, writable: false)).Render.Pages("1-" + count,
            new PdfPageRenderOptions {
                Format = OfficeImageExportFormat.Png,
                Scale = scale,
                Background = OfficeIMO.OfficeColor.White,
                MaxPages = count,
                ContinueOnError = false
            });
        return results.Select(result => result.Bytes ?? throw new InvalidDataException($"OfficeIMO did not render page {result.PageNumber}.")).ToArray();
    }

    private static byte[][] RenderPdfium(byte[] pdf, int count, double scale, bool encode, out double renderMilliseconds) {
        var pagesOut = new byte[count][];
        renderMilliseconds = 0D;
        using IDocReader document = DocLib.Instance.GetDocReader(pdf, new PageDimensions(scale));
        for (int page = 0; page < count; page++) {
            var stopwatch = Stopwatch.StartNew();
            using IPageReader reader = document.GetPageReader(page);
            byte[] bgra = reader.GetImage(RenderFlags.RenderAnnotations);
            int width = reader.GetPageWidth(), height = reader.GetPageHeight();
            renderMilliseconds += stopwatch.Elapsed.TotalMilliseconds;
            if (!encode) continue;
            // PDFium leaves unpainted pixels transparent; composite on white like the OfficeIMO lane.
            var rgba = new byte[bgra.Length];
            for (int index = 0; index < bgra.Length; index += 4) {
                int alpha = bgra[index + 3], inverse = 255 - alpha;
                rgba[index] = (byte)((bgra[index + 2] * alpha + 255 * inverse + 127) / 255);
                rgba[index + 1] = (byte)((bgra[index + 1] * alpha + 255 * inverse + 127) / 255);
                rgba[index + 2] = (byte)((bgra[index] * alpha + 255 * inverse + 127) / 255);
                rgba[index + 3] = 255;
            }
            pagesOut[page] = OfficeRasterImageEncoder.Encode(OfficeRasterImage.FromRgba32(width, height, rgba), OfficeImageExportFormat.Png, EncodingOptions);
        }
        return pagesOut;
    }

    // Mean absolute luma difference (0-255) and the share of pixels whose luma differs by at most 32.
    private static (double MeanDifference, double WithinTolerance) Compare(string name, double scale, byte[][] officeImo, byte[][] pdfium) {
        long difference = 0L, within = 0L, pixels = 0L;
        for (int page = 0; page < officeImo.Length; page++) {
            OfficeRasterImage left = OfficeRasterImageDecoder.Decode(officeImo[page]);
            OfficeRasterImage right = OfficeRasterImageDecoder.Decode(pdfium[page]);
            // OfficeIMO rounds page sizes up; PDFium rounds to nearest. Compare the shared area.
            if (Math.Abs(left.Width - right.Width) > 1 || Math.Abs(left.Height - right.Height) > 1) {
                throw new InvalidDataException(
                    $"{name} page {page + 1} at {scale}x: OfficeIMO {left.Width}x{left.Height} and PDFium {right.Width}x{right.Height} differ.");
            }
            byte[] a = left.GetPixels(), b = right.GetPixels();
            RequireInk(a, $"OfficeIMO {name} page {page + 1}");
            RequireInk(b, $"PDFium {name} page {page + 1}");
            int width = Math.Min(left.Width, right.Width), height = Math.Min(left.Height, right.Height);
            for (int y = 0; y < height; y++) {
                for (int x = 0; x < width; x++) {
                    int delta = Math.Abs(Luma(a, (y * left.Width + x) * 4) - Luma(b, (y * right.Width + x) * 4));
                    difference += delta;
                    if (delta <= 32) within++;
                    pixels++;
                }
            }
        }
        return ((double)difference / pixels, (double)within / pixels);
    }

    private static void RequireInk(byte[] rgba, string label) {
        long ink = 0L;
        for (int index = 0; index < rgba.Length; index += 4) if (Luma(rgba, index) < 224) ink++;
        if (ink * 1000L < rgba.Length / 4) throw new InvalidDataException(label + " rendered blank.");
    }

    private static int Luma(byte[] rgba, int offset) =>
        (rgba[offset] * 299 + rgba[offset + 1] * 587 + rgba[offset + 2] * 114) / 1000;

    private static int ReadInt(string[] args, string option, int fallback) {
        string? value = ReadOption(args, option);
        return value == null ? fallback : int.Parse(value, CultureInfo.InvariantCulture);
    }

    private static string? ReadOption(string[] args, string option) {
        for (int index = 1; index < args.Length - 1; index++) {
            if (string.Equals(args[index], option, StringComparison.OrdinalIgnoreCase)) return args[index + 1];
        }
        return null;
    }
}
