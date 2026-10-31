using OfficeIMO.Pdf;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.DocumentLayoutAnalysis.ReadingOrderDetector;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static partial class PdfBenchmarkValidation {
    /// <summary>Checks the complete report source, page geometry, tags and shared font before accepting HTML timings.</summary>
    internal static PdfReadObservation ValidateHtmlGenerated(byte[] bytes, PdfBenchmarkScenario scenario, string engine) {
        ValidateTaggedStructure(bytes, engine, scenario);
        ValidateHtmlLogicalSourceOrder(bytes, scenario, engine);
        PdfReadObservation observation = ValidateHtmlVisibleSource(bytes, scenario, engine);
        ValidateSharedFontUse(bytes, engine, "Carlito", requireBold: true);
        return observation;
    }

    // The fixed horizontal report has an authored visual order. Physical PDF
    // painting order is independent of both this order and the structure tree.
    private static PdfReadObservation ValidateHtmlVisibleSource(byte[] bytes, PdfBenchmarkScenario scenario, string engine) {
        if (bytes.Length < 5 || !bytes.AsSpan(0, 4).SequenceEqual("%PDF"u8)) {
            throw new InvalidDataException($"{engine} did not produce a PDF header for {scenario.Scale}.");
        }
        var pageTexts = new List<string>();
        using (var stream = new MemoryStream(bytes, writable: false)) {
            using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream);
            var wordExtractor = new NearestNeighbourWordExtractor(
                new NearestNeighbourWordExtractor.NearestNeighbourWordExtractorOptions { MaxDegreeOfParallelism = 3 });
            var pageSegmenter = new DefaultPageSegmenter(
                new DefaultPageSegmenter.DefaultPageSegmenterOptions { MaxDegreeOfParallelism = 3 });
            for (int index = 1; index <= document.NumberOfPages; index++) {
                var page = document.GetPage(index);
                // One point permits the browser's CSS-pixel rounding at an A4 edge.
                if (Math.Abs(page.Width - 210D / 25.4D * 72D) > 1D ||
                    Math.Abs(page.Height - 297D / 25.4D * 72D) > 1D) {
                    throw new InvalidDataException($"{engine} page {index} did not retain the A4 portrait geometry.");
                }
                if (page.Letters.Any(letter => !string.IsNullOrWhiteSpace(letter.Value)
                    && letter.TextOrientation != TextOrientation.Horizontal)) {
                    throw new InvalidDataException($"{engine} page {index} is outside the horizontal report text contract.");
                }
                string expected = scenario.PageTitle(index)
                    + string.Concat(Enumerable.Range(0, scenario.ParagraphsPerPage).Select(p => scenario.Narrative(index, p)))
                    + string.Concat(scenario.TableRows(index).SelectMany(row => row));
                // This fixture is one top-to-bottom report with left-to-right table rows.
                // Use the independent layout owners, not the producer's text painting order.
                var words = wordExtractor.GetWords(page.Letters);
                var blocks = pageSegmenter.GetBlocks(words);
                var lines = ReadingOrderHelper.OrderByReadingOrder(blocks.SelectMany(block => block.TextLines));
                string actual = string.Concat(lines.SelectMany(line => ReadingOrderHelper.OrderByReadingOrder(line.Words))
                    .Select(word => word.Text));
                string raw = string.Concat(page.Letters.Select(letter => letter.Value));
                string expectedCharacters = WithoutWhitespace(expected);
                string actualCharacters = WithoutWhitespace(actual);
                // Segmentation must neither discard nor invent non-whitespace glyph text,
                // including duplicated overlapping letters and punctuation.
                if (!WithoutWhitespace(raw).OrderBy(value => value).SequenceEqual(expectedCharacters.OrderBy(value => value))
                    || !string.Equals(actualCharacters, expectedCharacters, StringComparison.Ordinal)) {
                    throw new InvalidDataException($"{engine} page {index} changed, omitted, reordered or duplicated the visible report source.");
                }
                pageTexts.Add(actual);
            }
        }
        PdfReadObservation observation = Observe(pageTexts.Count, string.Join("\n", pageTexts));
        ValidateRead(observation, scenario, engine);
        ValidateScenarioPages(pageTexts, scenario, engine);
        return observation;
    }

    /// <summary>Reuses the multilingual fixture's font-use contract with structured embedded-font inspection.</summary>
    internal static void ValidateSharedFontUse(byte[] bytes, string engine, string familyName, string? requiredText = null, bool requireBold = false) {
        string expected = NormalizeBenchmarkFontName(familyName);
        using var input = new MemoryStream(bytes, writable: false);
        using PdfDocument pdfDocument = PdfDocument.Load(input);
        PdfFontInventory inventory = pdfDocument.Resources.Fonts();
        if (!inventory.Fonts.Any(font => font.IsEmbedded && NormalizeBenchmarkFontName(font.FamilyName).Contains(expected, StringComparison.Ordinal))) {
            throw new InvalidDataException($"{engine} did not embed the shared '{familyName}' fixture font.");
        }
        if (requireBold && (!inventory.Fonts.Any(font => font.IsEmbedded && font.IsBold && NormalizeBenchmarkFontName(font.FamilyName).Contains(expected, StringComparison.Ordinal))
            || !inventory.Fonts.Any(font => font.IsEmbedded && !font.IsBold && NormalizeBenchmarkFontName(font.FamilyName).Contains(expected, StringComparison.Ordinal)))) {
            throw new InvalidDataException($"{engine} did not embed both regular and bold '{familyName}' fixture faces.");
        }
        using var stream = new MemoryStream(bytes, writable: false);
        using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream);
        var selected = document.GetPages().SelectMany(page => page.Letters)
            .Where(letter => requiredText == null
                ? letter.Value.Any(char.IsLetterOrDigit)
                : letter.Value.Contains(requiredText, StringComparison.Ordinal)).ToArray();
        if (selected.Length == 0 || selected.Any(letter => !NormalizeBenchmarkFontName(letter.FontName ?? string.Empty).Contains(expected, StringComparison.Ordinal))) {
            throw new InvalidDataException($"{engine} did not use the shared '{familyName}' fixture font for every required text sample.");
        }
    }

    private static string NormalizeBenchmarkFontName(string value) =>
        new(value.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());

    private static string WithoutWhitespace(string value) => new(value.Where(c => !char.IsWhiteSpace(c)).ToArray());
}
