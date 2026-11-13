using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// Measures composition of live native text that overflows one slide: the layout engine measures the
/// paragraphs, splits them between sentences and words, and continues them on further slides. This is the
/// path PSWriteOffice content blocks take; it runs in process because the measured work is the layout itself.
/// </summary>
internal static class PowerPointNativeTextLayoutBenchmark {
    private const string Title = "Native text continuation";
    private static readonly Regex FieldId = new(@"<a:fld id=""\{[0-9A-F-]+\}""", RegexOptions.CultureInvariant);
    private const string Usage = "Usage: --native-text-layout [--iterations <n>] [--output <deck.pptx>] [--images <directory>]";

    internal static int Run(string[] args) {
        int iterations = 20;
        string? output = null, images = null;
        for (int index = 0; index < args.Length; index++) {
            string value = index + 1 < args.Length ? args[index + 1] : throw new ArgumentException(Usage);
            switch (args[index]) {
                case "--iterations": iterations = int.Parse(value, CultureInfo.InvariantCulture); break;
                case "--output": output = value; break;
                case "--images": images = value; break;
                default: throw new ArgumentException(Usage);
            }
            index++;
        }
        if (iterations < 1) throw new ArgumentException(Usage);

        byte[] deck = Compose();
        Validate(deck);
        for (int warmup = 0; warmup < 3; warmup++) Compose();

        var times = new double[iterations];
        long allocated = 0L;
        for (int sample = 0; sample < iterations; sample++) {
            long before = GC.GetAllocatedBytesForCurrentThread();
            long start = Stopwatch.GetTimestamp();
            Compose();
            times[sample] = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
            allocated += GC.GetAllocatedBytesForCurrentThread() - before;
        }
        Array.Sort(times);

        Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"NATIVE_TEXT_LAYOUT|iterations={iterations}|median_ms={Median(times):0.00}|min_ms={times[0]:0.00}" +
                          $"|allocated_mb_per_op={allocated / (double)iterations / (1024 * 1024):0.00}"));
        Console.WriteLine($"NATIVE_TEXT_LAYOUT_SLIDES|sha256={SlideHash(deck, out int slides)}|slides={slides}");
        if (output != null) File.WriteAllBytes(output, deck);
        if (images != null) WriteImages(deck, images);
        return 0;
    }

    /// <summary>Builds one materialized slide whose native text body holds three overflowing paragraphs, and composes it.</summary>
    internal static byte[] Compose() {
        using var stream = new MemoryStream();
        using (PowerPointDocument document = PowerPointDocument.Create()) {
            using var materialization = new PowerPointLayoutMaterialization(document);
            var body = new PowerPointLayoutNativeText(materialization.CreateTextBody());
            string[] paragraphs = Paragraphs().ToArray();
            foreach (string text in paragraphs) {
                PowerPointParagraph paragraph = body.Box.AddParagraph();
                paragraph.ClearInlineContent();
                paragraph.AddText(text);
            }
            body.Freeze();
            var blocks = new PowerPointLayoutBlock[] { new PowerPointLayoutText(paragraphs) { NativeBody = body } };
            var plan = new PowerPointContentPlanSlide(Title, blocks, PowerPointLayout.Auto);
            materialization.Complete(plan, PowerPointCompositionOptions.FromDesign(
                PowerPointDeckDesign.FromBrand("#008C95", "native-text-layout")));
            document.Save(stream);
        }
        return stream.ToArray();
    }

    /// <summary>Sentences, a run without sentence punctuation, and a long unbroken token.</summary>
    private static IEnumerable<string> Paragraphs() {
        var sentences = new StringBuilder();
        for (int index = 1; index <= 60; index++) {
            sentences.Append("Finding ").Append(index.ToString("D2", CultureInfo.InvariantCulture))
                .Append(" records how the review team traced each control to its owner and evidence. ");
        }
        yield return sentences.ToString().TrimEnd();

        var words = new StringBuilder();
        for (int index = 1; index <= 220; index++) {
            words.Append("item").Append(index.ToString(CultureInfo.InvariantCulture)).Append(index % 7 == 0 ? "  " : " ");
        }
        yield return words.ToString().TrimEnd();

        yield return "Identifier " + string.Concat(Enumerable.Repeat("0123456789ABCDEF", 40)) + " ends here.";
    }

    private static void Validate(byte[] deck) {
        using var input = new MemoryStream(deck, writable: false);
        using PowerPointDocument document = PowerPointDocument.Load(input);
        var errors = document.ValidateDocument().ToArray();
        if (errors.Length > 0) throw new InvalidDataException("Composed deck failed Open XML validation: " + errors[0].Description);
        // Leave out each slide's title and slide number; what remains is that slide's share of the body.
        string all = string.Concat(document.Slides.SelectMany(slide => slide.TextBoxes).Select(box => box.Text)
            .Where(text => !text.StartsWith(Title, StringComparison.Ordinal) && !text.All(char.IsDigit)));
        string compact = new(all.Where(c => !char.IsWhiteSpace(c)).ToArray());
        foreach (string paragraph in Paragraphs()) {
            string expected = new(paragraph.Where(c => !char.IsWhiteSpace(c)).ToArray());
            if (!compact.Contains(expected, StringComparison.Ordinal)) {
                throw new InvalidDataException("A native paragraph lost text across its continuation slides.");
            }
        }
        if (document.Slides.Count < 3) {
            throw new InvalidDataException("The native text did not continue onto further slides: " + string.Join(" / ",
                document.Slides.Select(slide => string.Join(",", slide.TextBoxes.Select(box => box.Text.Length)))));
        }
    }

    /// <summary>Hash of every slide part's XML in slide order; zip timestamps and field GUIDs are excluded.</summary>
    private static string SlideHash(byte[] deck, out int slides) {
        using var archive = new ZipArchive(new MemoryStream(deck, writable: false), ZipArchiveMode.Read);
        ZipArchiveEntry[] entries = archive.Entries
            .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
                            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .OrderBy(entry => int.Parse(entry.FullName["ppt/slides/slide".Length..^4], CultureInfo.InvariantCulture))
            .ToArray();
        slides = entries.Length;
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        foreach (ZipArchiveEntry entry in entries) {
            using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
            // Slide-number fields get a fresh GUID on every run; everything else must match byte for byte.
            string xml = FieldId.Replace(reader.ReadToEnd(), "<a:fld id=\"{}\"");
            hash.AppendData(Encoding.UTF8.GetBytes(xml));
        }
        return Convert.ToHexString(hash.GetHashAndReset());
    }

    private static void WriteImages(byte[] deck, string directory) {
        Directory.CreateDirectory(directory);
        using var input = new MemoryStream(deck, writable: false);
        using PowerPointDocument document = PowerPointDocument.Load(input);
        IReadOnlyList<OfficeImageExportResult> images = document.ToImages(OfficeImageExportFormat.Png);
        for (int index = 0; index < images.Count; index++) {
            File.WriteAllBytes(Path.Combine(directory, $"slide{index + 1:D2}.png"), images[index].Bytes);
        }
    }

    private static double Median(double[] sorted) =>
        sorted.Length % 2 == 1 ? sorted[sorted.Length / 2] : (sorted[sorted.Length / 2 - 1] + sorted[sorted.Length / 2]) / 2D;
}
