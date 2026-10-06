using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Xps;
using NativeXps = System.Windows.Xps.Packaging.XpsDocument;
using XpsDocument = OfficeIMO.Xps.XpsDocument;

namespace OfficeIMO.XpsWindowsEvidence;

internal static class Program {
    [STAThread]
    private static int Main(string[] args) {
        if (args.Length != 3) {
            Console.Error.WriteLine("Usage: <repository-root> <output-directory> <source-commit>");
            return 2;
        }
        string repository = Path.GetFullPath(args[0]);
        string output = Path.GetFullPath(args[1]);
        Directory.CreateDirectory(output);
        var results = new List<object>();
        int failures = 0;
        foreach (var fixture in EvidenceFixtures.Create(repository, output)) {
            string directory = Path.Combine(output, fixture.Name);
            Directory.CreateDirectory(directory);
            string sourcePath = Path.Combine(directory, "source.xps");
            if (fixture.IndependentPackage == null) fixture.Document.Save(sourcePath);
            else File.WriteAllBytes(sourcePath, fixture.IndependentPackage);
            try {
                OfficeRasterImage native = RenderNative(sourcePath, fixture.PageIndex);
                XpsDocument reopened = XpsDocument.Load(sourcePath);
                OfficeRasterImage managed = OfficeDrawingRasterRenderer.Render(
                    reopened.Pages[fixture.PageIndex].ToDrawing(), background: OfficeColor.White);
                byte[] pdf = reopened.ToPdf();
                OfficeRasterImage pdfRaster = OfficeDrawingRasterRenderer.Render(
                    PdfReadDocument.Open(pdf).Pages[fixture.PageIndex].ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
                File.WriteAllBytes(Path.Combine(directory, "native.png"), OfficePngWriter.Encode(native));
                File.WriteAllBytes(Path.Combine(directory, "managed.png"), OfficePngWriter.Encode(managed));
                File.WriteAllBytes(Path.Combine(directory, "pdf.png"), OfficePngWriter.Encode(pdfRaster));
                File.WriteAllBytes(Path.Combine(directory, "converted.pdf"), pdf);
                var svg = reopened.Pages[fixture.PageIndex].ToSvg();
                File.WriteAllText(Path.Combine(directory, "converted.svg"), svg.Svg);
                Comparison managedComparison = Compare(native, managed);
                Comparison pdfComparison = Compare(native, pdfRaster);
                object? lifecycle = fixture.IndependentPackage == null ? null
                    : CheckIndependentLifecycle(reopened, fixture.PageIndex, directory, native);
                results.Add(new {
                    fixture.Name, fixture.Producer, fixture.PageIndex,
                    DocumentCount = reopened.Documents.Count, PageCount = reopened.Pages.Count,
                    Text = reopened.Pages[fixture.PageIndex].ExtractText(),
                    PdfText = PdfReadDocument.Open(pdf).Pages[fixture.PageIndex].ExtractText(),
                    SourceSha256 = HashFile(sourcePath),
                    NativeSha256 = HashFile(Path.Combine(directory, "native.png")),
                    ManagedSha256 = HashFile(Path.Combine(directory, "managed.png")),
                    PdfSha256 = HashFile(Path.Combine(directory, "converted.pdf")),
                    PdfRasterSha256 = HashFile(Path.Combine(directory, "pdf.png")),
                    SvgSha256 = HashFile(Path.Combine(directory, "converted.svg")),
                    SvgDiagnostics = svg.Diagnostics, Managed = managedComparison, Pdf = pdfComparison,
                    Lifecycle = lifecycle
                });
                Console.WriteLine(FormattableString.Invariant($"{fixture.Name}: managed mean={managedComparison.MeanChannelError:F3}, PDF mean={pdfComparison.MeanChannelError:F3}, stable-interior max={managedComparison.InteriorMaximumChannelError}/{pdfComparison.InteriorMaximumChannelError}"));
            } catch (Exception exception) {
                failures++;
                results.Add(new { fixture.Name, SourceSha256 = HashFile(sourcePath), Error = exception.ToString() });
                Console.Error.WriteLine($"{fixture.Name}: {exception.Message}");
            }
        }
        var report = new {
            SchemaVersion = 1, GeneratedAtUtc = DateTimeOffset.UtcNow, SourceCommit = args[2],
            Platform = Environment.OSVersion.ToString(), Runtime = Environment.Version.ToString(),
            Consumer = "Microsoft WPF XpsDocument / DocumentPaginator / RenderTargetBitmap, 96 dpi",
            NativeAssemblies = new[] { Describe(typeof(NativeXps).Assembly), Describe(typeof(RenderTargetBitmap).Assembly) },
            OfficeAssemblies = new[] { Describe(typeof(XpsDocument).Assembly), Describe(typeof(OfficeRasterImage).Assembly),
                Describe(typeof(PdfReadDocument).Assembly) },
            Scope = "Microsoft XPS authored fixtures and one independent WPF two-document, three-page producer. No OpenXPS, interleaved, StoryFragments, printing, viewer or independent PDF-reader conformance claim.",
            Comparison = "RGB channels on white at 96 dpi. Stable interior excludes a one-pixel border and native 3x3 neighborhoods spanning more than 4 channel values. Differences remain evidence; successful collection is not equivalence.",
            Failures = failures, Results = results
        };
        File.WriteAllText(Path.Combine(output, "report.json"), JsonSerializer.Serialize(report,
            new JsonSerializerOptions { WriteIndented = true }));
        return failures == 0 ? 0 : 1;
    }

    internal static OfficeRasterImage RenderNative(string path, int pageIndex) {
        using var document = new NativeXps(path, FileAccess.Read);
        var sequence = document.GetFixedDocumentSequence()
            ?? throw new InvalidDataException("WPF did not expose a fixed document sequence.");
        using var page = sequence.DocumentPaginator.GetPage(pageIndex);
        int width = checked((int)Math.Ceiling(page.Size.Width));
        int height = checked((int)Math.Ceiling(page.Size.Height));
        var bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(page.Visual);
        var bytes = new byte[checked(width * height * 4)];
        bitmap.CopyPixels(bytes, width * 4, 0);
        var raster = new OfficeRasterImage(width, height, OfficeColor.White);
        for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
            int offset = (y * width + x) * 4, white = 255 - bytes[offset + 3];
            raster.SetPixel(x, y, OfficeColor.FromRgb(
                (byte)Math.Min(255, bytes[offset + 2] + white),
                (byte)Math.Min(255, bytes[offset + 1] + white),
                (byte)Math.Min(255, bytes[offset] + white)));
        }
        return raster;
    }

    private static object CheckIndependentLifecycle(XpsDocument document, int pageIndex, string directory, OfficeRasterImage native) {
        string text = "Native WPF page " + (pageIndex + 1);
        if (document.Pages[pageIndex].ExtractText() != text ||
            PdfReadDocument.Open(document.ToPdf()).Pages[pageIndex].ExtractText() != text)
            throw new InvalidDataException("Independent WPF text did not survive extraction and searchable PDF conversion.");
        string roundTripPath = Path.Combine(directory, "roundtrip.xps");
        document.Save(roundTripPath);
        Comparison preserved = Compare(native, RenderNative(roundTripPath, pageIndex));
        if (preserved.MaximumChannelError != 0)
            throw new InvalidDataException("The unchanged WPF package did not retain identical native pixels after saving.");
        document.Pages[pageIndex].AddPath("M180,140H192V152H180Z", "#FF00FF00");
        string editPath = Path.Combine(directory, "edited.xps");
        document.Save(editPath);
        XpsDocument edited = XpsDocument.Load(editPath);
        if (edited.Documents.Count != 2 || edited.Pages.Count != 3 || edited.Pages[pageIndex].ExtractText() != text)
            throw new InvalidDataException("Editing the WPF input lost document sequence, pages or text.");
        OfficeRasterImage painted = RenderNative(editPath, pageIndex);
        OfficeRasterImage managed = OfficeDrawingRasterRenderer.Render(edited.Pages[pageIndex].ToDrawing(), background: OfficeColor.White);
        var green = OfficeColor.FromRgb(0, 255, 0);
        if (painted.GetPixel(185, 145) != green || managed.GetPixel(185, 145) != green)
            throw new InvalidDataException("The independently produced package's loaded edit was not visible in both renderers.");
        return new { RoundTripSha256 = HashFile(roundTripPath), RoundTripNative = preserved,
            EditedSha256 = HashFile(editPath), Documents = edited.Documents.Count, Pages = edited.Pages.Count,
            Text = edited.Pages[pageIndex].ExtractText(), EditPixel = "#00FF00 at (185,145), native and managed" };
    }

    private static Comparison Compare(OfficeRasterImage reference, OfficeRasterImage actual) {
        if (reference.Width != actual.Width || reference.Height != actual.Height)
            throw new InvalidDataException($"Raster dimensions differ: {reference.Width}x{reference.Height} / {actual.Width}x{actual.Height}.");
        long sum = 0;
        int maximum = 0, changed = 0, interior = 0, interiorMaximum = 0, interiorChanged = 0;
        for (int y = 0; y < reference.Height; y++) for (int x = 0; x < reference.Width; x++) {
            OfficeColor expected = reference.GetPixel(x, y), observed = actual.GetPixel(x, y);
            int red = Math.Abs(expected.R - observed.R), green = Math.Abs(expected.G - observed.G),
                blue = Math.Abs(expected.B - observed.B), error = Math.Max(red, Math.Max(green, blue));
            sum += red + green + blue;
            maximum = Math.Max(maximum, error);
            if (error > 4) changed++;
            if (x == 0 || y == 0 || x == reference.Width - 1 || y == reference.Height - 1) continue;
            bool stable = true;
            for (int dy = -1; dy <= 1; dy++) for (int dx = -1; dx <= 1; dx++) {
                OfficeColor neighbor = reference.GetPixel(x + dx, y + dy);
                if (Math.Abs(expected.R - neighbor.R) > 4 || Math.Abs(expected.G - neighbor.G) > 4 ||
                    Math.Abs(expected.B - neighbor.B) > 4) stable = false;
            }
            if (!stable) continue;
            interior++;
            interiorMaximum = Math.Max(interiorMaximum, error);
            if (error > 4) interiorChanged++;
        }
        return new(reference.Width, reference.Height, sum / (reference.Width * (double)reference.Height * 3),
            maximum, changed, interior, interiorMaximum, interiorChanged);
    }

    private static object Describe(Assembly assembly) => new {
        assembly.FullName, FileVersion = FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion,
        Sha256 = HashFile(assembly.Location)
    };
    private static string HashFile(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
    private sealed record Comparison(int Width, int Height, double MeanChannelError, int MaximumChannelError,
        int PixelsAboveFour, int StableInteriorPixels, int InteriorMaximumChannelError, int InteriorPixelsAboveFour);
}
