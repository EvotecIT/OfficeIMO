using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static partial class PdfCorpusRunner {
    private static async Task<PdfCorpusManifest> ReadManifestAsync(string path) {
        PdfCorpusManifest manifest = JsonSerializer.Deserialize<PdfCorpusManifest>(
            await File.ReadAllTextAsync(path).ConfigureAwait(false),
            JsonOptions) ?? throw new InvalidDataException($"PDF corpus manifest is empty: {path}");
        if (manifest.SchemaVersion != 1) {
            throw new InvalidDataException(
                $"Unsupported PDF corpus schema {manifest.SchemaVersion} in {path}; expected 1.");
        }
        ValidateManifest(manifest, path);
        return manifest;
    }

    private static void ValidateManifest(PdfCorpusManifest manifest, string path) {
        if (manifest.Entries is null) {
            throw new InvalidDataException($"PDF corpus manifest has no entries array: {path}");
        }

        for (int index = 0; index < manifest.Entries.Count; index++) {
            PdfCorpusEntry entry = manifest.Entries[index]
                ?? throw new InvalidDataException($"PDF corpus manifest entry {index} is null: {path}");
            string label = string.IsNullOrWhiteSpace(entry.Id) ? $"entry {index}" : entry.Id;
            if (string.IsNullOrWhiteSpace(entry.SourceKind)) {
                throw new InvalidDataException($"PDF corpus {label} has no sourceKind in {path}.");
            }
            if (string.IsNullOrWhiteSpace(entry.Producer) ||
                string.IsNullOrWhiteSpace(entry.License) ||
                string.IsNullOrWhiteSpace(entry.Tier)) {
                throw new InvalidDataException(
                    $"PDF corpus {label} must declare producer, license, and tier in {path}.");
            }
            if (entry.ExpectedPages is <= 0) {
                throw new InvalidDataException($"PDF corpus {label} expectedPages must be positive in {path}.");
            }
            if (double.IsNaN(entry.MinimumTokenRecall) ||
                double.IsInfinity(entry.MinimumTokenRecall) ||
                entry.MinimumTokenRecall < 0D ||
                entry.MinimumTokenRecall > 1D) {
                throw new InvalidDataException(
                    $"PDF corpus {label} minimumTokenRecall must be finite and between 0 and 1 in {path}.");
            }
            if (entry.Features is null ||
                entry.RequiredText is null ||
                entry.ExpectedText is null ||
                entry.PageExpectations is null) {
                throw new InvalidDataException(
                    $"PDF corpus {label} collection properties cannot be null in {path}.");
            }
            ValidateNonBlankValues(entry.Features, label, "features", path);
            ValidateNonBlankValues(entry.RequiredText, label, "requiredText", path);
            ValidateNonBlankValues(entry.ExpectedText, label, "expectedText", path);

            string sourceKind = entry.SourceKind.Trim().ToLowerInvariant();
            switch (sourceKind) {
                case "repository":
                case "local":
                    RequireSourceField(entry.SourcePath, label, "sourcePath", path);
                    RejectSourceField(entry.Url, label, "url", sourceKind, path);
                    RejectSourceField(entry.Generator, label, "generator", sourceKind, path);
                    ValidateSha256(entry.Sha256, label, path);
                    break;
                case "download":
                    RequireSourceField(entry.Url, label, "url", path);
                    if (!Uri.TryCreate(entry.Url, UriKind.Absolute, out Uri? uri) ||
                        uri.Scheme is not ("http" or "https")) {
                        throw new InvalidDataException(
                            $"PDF corpus {label} url must be an absolute HTTP or HTTPS URL in {path}.");
                    }
                    RejectSourceField(entry.SourcePath, label, "sourcePath", sourceKind, path);
                    RejectSourceField(entry.Generator, label, "generator", sourceKind, path);
                    ValidateSha256(entry.Sha256, label, path);
                    break;
                case "generated":
                    RequireSourceField(entry.Generator, label, "generator", path);
                    RejectSourceField(entry.SourcePath, label, "sourcePath", sourceKind, path);
                    RejectSourceField(entry.Url, label, "url", sourceKind, path);
                    if (!string.IsNullOrWhiteSpace(entry.Sha256)) ValidateSha256(entry.Sha256, label, path);
                    break;
                default:
                    throw new InvalidDataException(
                        $"PDF corpus {label} has unsupported sourceKind '{entry.SourceKind}' in {path}.");
            }

            ValidatePageExpectations(entry, label, path);
        }
    }

    private static void ValidatePageExpectations(PdfCorpusEntry entry, string label, string path) {
        var pageNumbers = new HashSet<int>();
        for (int index = 0; index < entry.PageExpectations.Count; index++) {
            PdfCorpusPageExpectation expectation = entry.PageExpectations[index]
                ?? throw new InvalidDataException(
                    $"PDF corpus {label} pageExpectations entry {index} is null in {path}.");
            if (expectation.PageNumber <= 0 ||
                entry.ExpectedPages.HasValue && expectation.PageNumber > entry.ExpectedPages.Value) {
                throw new InvalidDataException(
                    $"PDF corpus {label} has invalid semantic page {expectation.PageNumber} in {path}.");
            }
            if (!pageNumbers.Add(expectation.PageNumber)) {
                throw new InvalidDataException(
                    $"PDF corpus {label} repeats semantic page {expectation.PageNumber} in {path}.");
            }
            ValidateExactCount(expectation.ExpectedTables, label, expectation.PageNumber, "expectedTables", path);
            ValidateExactCount(expectation.ExpectedImages, label, expectation.PageNumber, "expectedImages", path);
            ValidateExactCount(expectation.ExpectedImageRegions, label, expectation.PageNumber, "expectedImageRegions", path);
            ValidateExactCount(expectation.ExpectedFigures, label, expectation.PageNumber, "expectedFigures", path);
            ValidatePositiveMinimum(expectation.MinimumVectorPrimitives, label, expectation.PageNumber, "minimumVectorPrimitives", path);
            if (expectation.Tables is null) {
                throw new InvalidDataException(
                    $"PDF corpus {label} page {expectation.PageNumber} tables cannot be null in {path}.");
            }
            ValidateTableExpectations(expectation, label, index, path);
            if (!expectation.ExpectedTables.HasValue &&
                !expectation.ExpectedImages.HasValue &&
                !expectation.ExpectedImageRegions.HasValue &&
                !expectation.ExpectedFigures.HasValue &&
                !expectation.MinimumVectorPrimitives.HasValue &&
                expectation.Tables.Count == 0) {
                throw new InvalidDataException(
                    $"PDF corpus {label} page {expectation.PageNumber} has an empty semantic expectation in {path}.");
            }
        }
    }

    private static void ValidateTableExpectations(
        PdfCorpusPageExpectation expectation,
        string label,
        int pageExpectationIndex,
        string path) {
        if (expectation.Tables.Count > 0 &&
            (!expectation.ExpectedTables.HasValue || expectation.ExpectedTables.Value != expectation.Tables.Count)) {
            throw new InvalidDataException(
                $"PDF corpus {label} page {expectation.PageNumber} must set expectedTables to its {expectation.Tables.Count} table contracts in {path}.");
        }
        for (int tableIndex = 0; tableIndex < expectation.Tables.Count; tableIndex++) {
            PdfCorpusTableExpectation table = expectation.Tables[tableIndex]
                ?? throw new InvalidDataException(
                    $"PDF corpus {label} pageExpectations[{pageExpectationIndex}].tables[{tableIndex}] is null in {path}.");
            if (table.Rows <= 0 || table.Columns < 2) {
                throw new InvalidDataException(
                    $"PDF corpus {label} page {expectation.PageNumber} table {tableIndex} must declare positive rows and at least two columns in {path}.");
            }
            if (table.RequiredCells is null || table.RequiredCells.Count == 0) {
                throw new InvalidDataException(
                    $"PDF corpus {label} page {expectation.PageNumber} table {tableIndex} must declare requiredCells in {path}.");
            }
            ValidateNonBlankValues(
                table.RequiredCells,
                label,
                $"pageExpectations[{pageExpectationIndex}].tables[{tableIndex}].requiredCells",
                path);
            if (table.RequiredCells.Distinct(StringComparer.Ordinal).Count() != table.RequiredCells.Count) {
                throw new InvalidDataException(
                    $"PDF corpus {label} page {expectation.PageNumber} table {tableIndex} repeats requiredCells in {path}.");
            }
        }
    }

    private static void ValidateExactCount(
        int? value,
        string label,
        int pageNumber,
        string property,
        string path) {
        if (value < 0) {
            throw new InvalidDataException(
                $"PDF corpus {label} page {pageNumber} {property} cannot be negative in {path}.");
        }
    }

    private static void ValidatePositiveMinimum(
        int? value,
        string label,
        int pageNumber,
        string property,
        string path) {
        if (value is <= 0) {
            throw new InvalidDataException(
                $"PDF corpus {label} page {pageNumber} {property} must be positive when specified in {path}.");
        }
    }

    private static void ValidateNonBlankValues(
        IReadOnlyList<string> values,
        string label,
        string property,
        string path) {
        if (values.Any(string.IsNullOrWhiteSpace)) {
            throw new InvalidDataException(
                $"PDF corpus {label} {property} cannot contain null or blank values in {path}.");
        }
    }

    private static void RequireSourceField(string? value, string label, string property, string path) {
        if (string.IsNullOrWhiteSpace(value)) {
            throw new InvalidDataException($"PDF corpus {label} must declare {property} in {path}.");
        }
    }

    private static void RejectSourceField(
        string? value,
        string label,
        string property,
        string sourceKind,
        string path) {
        if (!string.IsNullOrWhiteSpace(value)) {
            throw new InvalidDataException(
                $"PDF corpus {label} cannot declare {property} for sourceKind '{sourceKind}' in {path}.");
        }
    }

    private static void ValidateSha256(string? value, string label, string path) {
        if (string.IsNullOrWhiteSpace(value) || !Sha256Regex().IsMatch(value)) {
            throw new InvalidDataException(
                $"PDF corpus {label} must declare a 64-character hexadecimal sha256 in {path}.");
        }
    }

    private static PdfCorpusSemanticObservation CreateSemanticObservation(
        OfficeIMO.Pdf.PdfDocumentReadResult document) {
        PdfCorpusPageSemanticObservation[] pages = document.Pages
            .Select(static page => new PdfCorpusPageSemanticObservation(
                page.PageNumber,
                page.Tables.Count,
                page.Images.Count,
                page.Analysis.ImageRegions.Count,
                page.Analysis.ImageRegions.Count(static region => region.IsFigure),
                page.VectorPrimitiveCount))
            .ToArray();
        return new PdfCorpusSemanticObservation(
            pages.Sum(static page => page.Tables),
            pages.Sum(static page => page.Images),
            pages.Sum(static page => page.ImageRegions),
            pages.Sum(static page => page.Figures),
            pages.Sum(static page => page.VectorPrimitives),
            pages);
    }

    private static void ValidatePageExpectations(
        PdfCorpusEntry entry,
        OfficeIMO.Pdf.PdfDocumentReadResult document) {
        foreach (PdfCorpusPageExpectation expectation in entry.PageExpectations) {
            OfficeIMO.Pdf.PdfLogicalPage page = document.Pages.SingleOrDefault(
                candidate => candidate.PageNumber == expectation.PageNumber)
                ?? throw new InvalidDataException(
                    $"{entry.Id} has no page {expectation.PageNumber} required by its semantic oracle.");
            RequireExact(entry.Id, page.PageNumber, "tables", page.Tables.Count, expectation.ExpectedTables);
            RequireExact(entry.Id, page.PageNumber, "images", page.Images.Count, expectation.ExpectedImages);
            RequireExact(
                entry.Id,
                page.PageNumber,
                "image regions",
                page.Analysis.ImageRegions.Count,
                expectation.ExpectedImageRegions);
            RequireExact(
                entry.Id,
                page.PageNumber,
                "figures",
                page.Analysis.ImageRegions.Count(static region => region.IsFigure),
                expectation.ExpectedFigures);
            RequireMinimum(
                entry.Id,
                page.PageNumber,
                "vector primitives",
                page.VectorPrimitiveCount,
                expectation.MinimumVectorPrimitives);

            ValidateTables(entry.Id, page, expectation.Tables);
        }
    }

    private static void ValidateTables(
        string entryId,
        OfficeIMO.Pdf.PdfLogicalPage page,
        IReadOnlyList<PdfCorpusTableExpectation> expectations) {
        if (expectations.Count == 0) return;
        var unmatched = page.Tables.ToList();
        for (int expectationIndex = 0; expectationIndex < expectations.Count; expectationIndex++) {
            PdfCorpusTableExpectation expectation = expectations[expectationIndex];
            OfficeIMO.Pdf.PdfLogicalTable[] matches = unmatched
                .Where(table => TableMatches(table, expectation))
                .ToArray();
            if (matches.Length != 1) {
                string observed = string.Join("; ", page.Tables.Select(static table =>
                    $"{table.Rows.Count}x{table.Columns.Count} [{string.Join(" | ", table.Rows.SelectMany(static row => row))}]"));
                if (observed.Length > 700) observed = observed.Substring(0, 700) + "...";
                throw new InvalidDataException(
                    $"OfficeIMO matched {matches.Length} tables to contract {expectationIndex} " +
                    $"({expectation.Rows}x{expectation.Columns}) on page {page.PageNumber} of {entryId}; expected exactly one. " +
                    $"Observed tables: {observed}");
            }
            unmatched.Remove(matches[0]);
        }
    }

    private static bool TableMatches(
        OfficeIMO.Pdf.PdfLogicalTable table,
        PdfCorpusTableExpectation expectation) {
        if (table.Rows.Count != expectation.Rows || table.Columns.Count != expectation.Columns) return false;
        var cells = new HashSet<string>(table.Rows.SelectMany(static row => row), StringComparer.Ordinal);
        return expectation.RequiredCells.All(cells.Contains);
    }

    private static void RequireExact(
        string entryId,
        int pageNumber,
        string feature,
        int actual,
        int? expected) {
        if (expected.HasValue && actual != expected.Value) {
            throw new InvalidDataException(
                $"OfficeIMO recovered {actual} {feature} on page {pageNumber} of {entryId}; expected exactly {expected.Value}.");
        }
    }

    private static void RequireMinimum(
        string entryId,
        int pageNumber,
        string feature,
        int actual,
        int? minimum) {
        if (minimum.HasValue && actual < minimum.Value) {
            throw new InvalidDataException(
                $"OfficeIMO recovered {actual} {feature} on page {pageNumber} of {entryId}; expected at least {minimum.Value}.");
        }
    }

    private static void ValidateExpectedText(
        PdfCorpusEntry entry,
        IReadOnlyList<string> officePages) {
        if (entry.ExpectedText.Count == 0) return;

        string officeText = string.Join('\n', officePages);
        foreach (string expectedText in entry.ExpectedText) {
            if (!officeText.Contains(expectedText, StringComparison.Ordinal)) {
                throw new InvalidDataException(
                    $"OfficeIMO did not recover labelled text '{expectedText}' from {entry.Id}.");
            }
        }
    }
}
