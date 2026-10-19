using OfficeIMO.Pdf;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static partial class PdfBenchmarkValidation {
    /// <summary>Checks the complete report source, page geometry, tags and shared font before accepting HTML timings.</summary>
    internal static PdfReadObservation ValidateHtmlGenerated(byte[] bytes, PdfBenchmarkScenario scenario, string engine) {
        PdfReadObservation observation = ValidateGenerated(bytes, scenario, engine);
        ValidateTaggedStructure(bytes, engine, scenario);
        using (var stream = new MemoryStream(bytes, writable: false)) {
            using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream);
            for (int index = 1; index <= document.NumberOfPages; index++) {
                var page = document.GetPage(index);
                // One point permits the browser's CSS-pixel rounding at an A4 edge.
                if (Math.Abs(page.Width - 210D / 25.4D * 72D) > 1D ||
                    Math.Abs(page.Height - 297D / 25.4D * 72D) > 1D) {
                    throw new InvalidDataException($"{engine} page {index} did not retain the A4 portrait geometry.");
                }
                string expected = scenario.PageTitle(index)
                    + string.Concat(Enumerable.Range(0, scenario.ParagraphsPerPage).Select(p => scenario.Narrative(index, p)))
                    + string.Concat(scenario.TableRows(index).SelectMany(row => row));
                string actual = ContentOrderTextExtractor.GetText(page);
                if (!string.Equals(WithoutWhitespace(actual), WithoutWhitespace(expected), StringComparison.Ordinal)) {
                    throw new InvalidDataException($"{engine} page {index} changed, omitted, reordered or duplicated the report source.");
                }
            }
        }
        ValidateSharedFontUse(bytes, engine, "Carlito", requireBold: true);
        return observation;
    }

    /// <summary>Reuses the multilingual fixture's font-use contract with structured embedded-font inspection.</summary>
    internal static void ValidateSharedFontUse(byte[] bytes, string engine, string familyName, string? requiredText = null, bool requireBold = false) {
        string expected = NormalizeBenchmarkFontName(familyName);
        PdfFontInventory inventory = PdfDocument.Load(bytes).Resources.Fonts();
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
