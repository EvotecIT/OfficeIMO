using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public readonly record struct PdfReadObservation(
    int PageCount,
    int TextLength,
    int ReportMarkerCount,
    long CharacterChecksum,
    string NormalizedText);

internal sealed record PdfExpectedPage(IReadOnlyList<string> RequiredFragments);

internal static class PdfBenchmarkValidation {
    internal static PdfReadObservation ValidateGenerated(byte[] bytes, PdfBenchmarkScenario scenario, string engine) {
        if (bytes.Length < 5 || !bytes.AsSpan(0, 4).SequenceEqual("%PDF"u8)) {
            throw new InvalidDataException($"{engine} did not produce a PDF header for {scenario.Scale}.");
        }

        PdfReadObservation observation = ReadWithPdfPig(bytes);
        ValidateRead(observation, scenario, engine);
        return observation;
    }

    internal static void ValidateTaggedStructure(byte[] bytes, string engine, int expectedPageCount) {
        PdfDocumentInfo info = OfficeIMO.Pdf.PdfDocument.Open(bytes).Inspect();
        PdfTaggedContentInfo? tagged = info.TaggedContent;
        if (!info.HasTaggedContent ||
            tagged == null ||
            !string.Equals(info.CatalogLanguage, "en-US", StringComparison.OrdinalIgnoreCase) ||
            tagged.Marked != true ||
            tagged.StructureElements.Count == 0 ||
            tagged.MarkedContentReferenceCount == 0 ||
            !tagged.HasDocumentStructureElement ||
            tagged.ParentTreeEntryCount < expectedPageCount ||
            !tagged.FiguresHaveAlternateText) {
            throw new InvalidDataException($"{engine} did not preserve a complete tagged-document structure tree.");
        }

        foreach (string requiredType in new[] { "H1", "P", "Table", "TR", "TH", "TD" }) {
            if (!tagged.StructureTypeCounts.TryGetValue(requiredType, out int count) || count == 0) {
                throw new InvalidDataException($"{engine} did not preserve required {requiredType} accessibility structure.");
            }
        }
    }

    internal static void ValidateRead(PdfReadObservation observation, PdfBenchmarkScenario scenario, string engine) {
        if (observation.PageCount != scenario.PageCount) {
            throw new InvalidDataException(
                $"{engine} produced or observed {observation.PageCount} pages for {scenario.Scale}; " +
                $"expected {scenario.PageCount}.");
        }

        if (observation.ReportMarkerCount != scenario.PageCount) {
            throw new InvalidDataException(
                $"{engine} observed {observation.ReportMarkerCount} report markers for {scenario.Scale}; " +
                $"expected {scenario.PageCount}.");
        }

        int minimumTextLength = scenario.PageCount * 100;
        if (observation.TextLength < minimumTextLength || observation.CharacterChecksum == 0L) {
            throw new InvalidDataException(
                $"{engine} under-read the {scenario.Scale} payload: " +
                $"length={observation.TextLength}, checksum={observation.CharacterChecksum}.");
        }

        ValidateScenarioContent(observation.NormalizedText, scenario, engine);
    }

    internal static PdfReadObservation ReadWithPdfPig(byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream);
        var text = new StringBuilder();
        foreach (var page in document.GetPages()) {
            text.Append(ContentOrderTextExtractor.GetText(page));
            text.Append('\n');
        }

        return Observe(document.NumberOfPages, text.ToString());
    }

    internal static PdfReadObservation Observe(int pageCount, string text) {
        string normalized = Normalize(text);
        int markerCount = CountOccurrences(normalized, "BENCHMARKREPORT");
        long checksum = 0L;
        foreach (char value in normalized) {
            checksum = unchecked((checksum * 31L) + value);
        }

        return new PdfReadObservation(pageCount, normalized.Length, markerCount, checksum, normalized);
    }

    internal static PdfExpectedPage ExpectedPage(PdfBenchmarkScenario scenario, int pageNumber) {
        var fragments = new List<string>(scenario.ParagraphsPerPage + scenario.RowsPerPage + 2) {
            Normalize(scenario.PageTitle(pageNumber))
        };
        for (int paragraph = 0; paragraph < scenario.ParagraphsPerPage; paragraph++) {
            fragments.Add(Normalize(scenario.Narrative(pageNumber, paragraph)));
        }
        foreach (string[] row in scenario.TableRows(pageNumber)) {
            fragments.Add(Normalize(string.Concat(row)));
        }
        return new PdfExpectedPage(fragments);
    }

    internal static void ValidatePageContent(string actualText, PdfExpectedPage expected, string context) {
        string actual = Normalize(actualText);
        foreach (string fragment in expected.RequiredFragments) {
            if (!actual.Contains(fragment, StringComparison.Ordinal)) {
                throw new InvalidDataException($"{context} did not preserve required content '{fragment}'.");
            }
        }
    }

    internal static void ValidateScenarioContent(string actual, PdfBenchmarkScenario scenario, string engine) {
        actual = Normalize(actual);
        var requiredCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int page = 1; page <= scenario.PageCount; page++) {
            foreach (string fragment in ExpectedPage(scenario, page).RequiredFragments) {
                requiredCounts.TryGetValue(fragment, out int count);
                requiredCounts[fragment] = count + 1;
            }
        }

        foreach ((string fragment, int requiredCount) in requiredCounts) {
            int actualCount = CountOccurrences(actual, fragment);
            if (actualCount < requiredCount) {
                throw new InvalidDataException(
                    $"{engine} preserved {actualCount} of {requiredCount} required occurrences for '{fragment}'.");
            }
        }
    }

    internal static void ValidateTableScenarioContent(string actual, PdfBenchmarkScenario scenario, string engine) {
        actual = Normalize(actual);
        for (int page = 1; page <= scenario.PageCount; page++) {
            foreach (string[] row in scenario.TableRows(page).Skip(1)) {
                string exactFragment = Normalize(string.Concat(row));
                string semanticFragment = string.Concat(row.Select(NormalizeTableCell));
                if (!actual.Contains(exactFragment, StringComparison.Ordinal) &&
                    !actual.Contains(semanticFragment, StringComparison.Ordinal)) {
                    throw new InvalidDataException(
                        $"{engine} did not preserve required table row '{exactFragment}'.");
                }
            }
        }
    }

    private static string NormalizeTableCell(string value) {
        if (decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal number)) {
            return Normalize(number.ToString("G29", CultureInfo.InvariantCulture));
        }

        return Normalize(value);
    }

    internal static string Normalize(string value) {
        var normalized = new StringBuilder(value.Length);
        foreach (char character in value) {
            if (char.IsLetterOrDigit(character)) {
                normalized.Append(char.ToUpperInvariant(character));
            }
        }

        return normalized.ToString();
    }

    private static int CountOccurrences(string value, string marker) {
        int count = 0;
        int offset = 0;
        while ((offset = value.IndexOf(marker, offset, StringComparison.Ordinal)) >= 0) {
            count++;
            offset += marker.Length;
        }

        return count;
    }
}
