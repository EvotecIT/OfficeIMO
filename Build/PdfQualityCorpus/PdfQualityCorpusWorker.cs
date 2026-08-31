using System.Diagnostics;
using System.Security.Cryptography;
using OfficeIMO.Pdf;

namespace OfficeIMO.PdfQualityCorpus;

internal static class PdfQualityCorpusWorker {
    internal static QualityCaseResult Probe(QualityProbeOptions options) {
        Stopwatch total = Stopwatch.StartNew();
        QualityManifest manifest = PdfQualityCorpusManifest.Load(options.ManifestPath);
        QualityCase item = manifest.Cases.SingleOrDefault(candidate => string.Equals(candidate.Id, options.CaseId, StringComparison.Ordinal))
            ?? throw new InvalidDataException("Unknown PDF quality corpus case: " + options.CaseId + ".");
        string path = PdfQualityCorpusManifest.ResolveCasePath(options.RootDirectory, item);
        byte[] bytes = ReadBounded(path, options.MaxFileBytes);
        string sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        if (!string.Equals(sha256, item.Sha256, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Corpus file SHA-256 does not match the manifest.");
        if (bytes.LongLength != item.ByteLength) throw new InvalidDataException("Corpus file byte length does not match the manifest.");

        var result = new QualityCaseResult {
            Id = item.Id,
            SourceId = item.Source,
            SourcePath = item.SourcePath,
            Sha256 = sha256,
            ByteLength = bytes.LongLength,
            Features = item.Features
        };
        var checks = new List<QualityCheckResult>();
        var metrics = result.Metrics;
        PdfDocument? document = null;
        PdfDocumentInfo? info = null;
        PdfLogicalDocument? logical = null;

        Run(checks, "open-and-inspect", () => {
            document = PdfDocument.Open(bytes);
            info = document.Read.DocumentInfo();
            metrics.PageCount = info.Pages.Count;
            metrics.AttachmentCount = info.AttachmentCount;
            metrics.LinkCount = info.LinkAnnotationCount;
            metrics.AnnotationCount = info.AnnotationCount;
            metrics.OptionalContentGroupCount = info.OptionalContentGroupCount;
            metrics.RepairCodes = document.Analyze().Repair.Diagnostics.Select(diagnostic => diagnostic.Code).ToArray();
            metrics.AnnotationActionTypes = info.AnnotationActionTypes.ToArray();
        });
        Run(checks, "text", () => metrics.TextCharacters = Require(document).Read.Text().Length);
        Run(checks, "logical-semantics", () => {
            logical = Require(document).Read.Logical();
            IReadOnlyList<PdfLogicalParagraphContinuationGroup> paragraphs = logical.GetParagraphContinuationGroups();
            IReadOnlyList<PdfLogicalTableContinuationGroup> tables = logical.GetTableContinuationGroups(
                new PdfLogicalTableContinuationOptions { MaxRows = 100_000 });
            metrics.ParagraphCount = logical.Paragraphs.Count;
            metrics.CrossPageParagraphCount = paragraphs.Count(paragraph => paragraph.SpansPages);
            metrics.TableCount = logical.Tables.Count;
            metrics.CrossPageTableCount = tables.Count(table => table.SpansPages);
        });
        Run(checks, "font-inspection", () => {
            PdfFontInventory fonts = Require(document).Read.Fonts();
            metrics.FontCount = fonts.FontCount;
            metrics.EmbeddedFontCount = fonts.EmbeddedFontCount;
            metrics.SubsetFontCount = fonts.SubsetFontCount;
            metrics.MissingToUnicodeFontCount = fonts.MissingToUnicodeFontCount;
            metrics.UnreadableToUnicodeFontCount = fonts.UnreadableToUnicodeFontCount;
            metrics.FontResourceReferenceCount = fonts.ResourceReferenceCount;
        });
        Run(checks, "image-inspection", () => {
            PdfDocument source = Require(document);
            metrics.ImageCount = source.Read.Images().Count;
            metrics.ImagePlacementCount = source.Read.ImagePlacements().Count;
        });
        Run(checks, "managed-svg-render", () => {
            int renderPages = Math.Min(metrics.PageCount, options.MaxRenderPages);
            if (renderPages == 0) return;
            IReadOnlyList<PdfPageRenderResult> renders = Require(document).Read.RenderPages(
                PdfPageSelection.From(new PdfPageRange(1, renderPages)),
                new PdfPageRenderOptions {
                    Format = PdfPageRenderFormat.Svg,
                    ContinueOnError = true,
                    MaxPages = renderPages
                });
            metrics.RenderAttemptedPages = renders.Count;
            metrics.RenderSucceededPages = renders.Count(render => render.Succeeded);
            metrics.RenderOutputBytes = renders.Sum(render => render.Bytes?.LongLength ?? 0L);
            metrics.RenderDiagnosticCodes = renders
                .SelectMany(render => render.CapabilityDiagnostics)
                .Select(diagnostic => diagnostic.Code)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(code => code, StringComparer.Ordinal)
                .ToArray();
        });
        Run(checks, "mutation-portfolio", () => {
            PdfMutationPortfolioReport portfolio = Require(document).AssessMutations();
            PdfMutationPlan plan = portfolio.Get(PdfMutationOperation.UpdateMetadata);
            metrics.MutationMode = plan.ExecutionMode.ToString();
            metrics.MutationBlockerCodes = plan.BlockerCodes.ToArray();
            metrics.MutationPlanCount = portfolio.Plans.Count;
            metrics.FullRewriteMutationPlanCount = portfolio.Plans.Count(item => item.ExecutionMode == PdfMutationExecutionMode.FullRewrite);
            metrics.AppendOnlyMutationPlanCount = portfolio.Plans.Count(item => item.ExecutionMode == PdfMutationExecutionMode.AppendOnly);
            metrics.BlockedMutationPlanCount = portfolio.Plans.Count(item => item.ExecutionMode == PdfMutationExecutionMode.Blocked);
            metrics.MutationPlanModes = portfolio.Plans.ToDictionary(
                item => item.Operation.ToString(),
                item => item.ExecutionMode.ToString(),
                StringComparer.Ordinal);
        });
        Run(checks, "declared-compliance-claims", () => {
            PdfDeclaredComplianceClaimsReport claims = Require(document).AssessDeclaredComplianceClaims();
            metrics.DeclaredComplianceClaimCount = claims.Claims.Count;
            metrics.RecognizedComplianceClaimCount = claims.RecognizedClaims.Count;
            metrics.ClaimableComplianceClaimCount = claims.Claims.Count(claim => claim.CanClaimConformance);
            metrics.UnsupportedComplianceClaimCount = claims.UnsupportedClaims.Count;
            metrics.DeclaredComplianceClaimStatuses = claims.Claims
                .Select(claim => claim.Declaration + ":" + claim.Status)
                .ToArray();
        });

        result.Checks = checks.AsReadOnly();
        result.Expectations = EvaluateExpectations(item, metrics);
        result.Outcome = checks.All(check => check.Succeeded) && result.Expectations.All(expectation => expectation.Succeeded)
            ? "passed"
            : "failed";
        total.Stop();
        result.DurationMilliseconds = total.ElapsedMilliseconds;
        return result;
    }

    private static IReadOnlyList<QualityExpectationResult> EvaluateExpectations(QualityCase item, QualityCaseMetrics metrics) {
        var results = new List<QualityExpectationResult> {
            Exact("page-count", item.PageCount, metrics.PageCount),
            Minimum("text-characters", item.MinimumTextCharacters, metrics.TextCharacters),
            Minimum("attachments", item.MinimumAttachments, metrics.AttachmentCount),
            Minimum("links", item.MinimumLinks, metrics.LinkCount),
            Minimum("annotations", item.MinimumAnnotations, metrics.AnnotationCount),
            Sequence("annotation-action-types", item.ExpectedAnnotationActionTypes, metrics.AnnotationActionTypes),
            Sequence("repair-codes", item.ExpectedRepairCodes, metrics.RepairCodes),
            Exact("render-succeeded", item.ExpectedRenderSucceeded, metrics.RenderAttemptedPages > 0 && metrics.RenderSucceededPages == metrics.RenderAttemptedPages),
            Sequence("render-diagnostic-codes", item.ExpectedRenderDiagnosticCodes, metrics.RenderDiagnosticCodes),
            Exact("mutation-mode", item.ExpectedMutationMode, metrics.MutationMode)
        };
        AddOptionalMinimum(results, "optional-content-groups", item.MinimumOptionalContentGroups, metrics.OptionalContentGroupCount);
        AddOptionalMinimum(results, "fonts", item.MinimumFonts, metrics.FontCount);
        AddOptionalMinimum(results, "embedded-fonts", item.MinimumEmbeddedFonts, metrics.EmbeddedFontCount);
        AddOptionalMinimum(results, "subset-fonts", item.MinimumSubsetFonts, metrics.SubsetFontCount);
        if (item.MaximumMissingToUnicodeFonts.HasValue) {
            results.Add(new QualityExpectationResult {
                Name = "missing-tounicode-fonts",
                Succeeded = metrics.MissingToUnicodeFontCount <= item.MaximumMissingToUnicodeFonts.Value,
                Expected = "<= " + item.MaximumMissingToUnicodeFonts.Value,
                Actual = metrics.MissingToUnicodeFontCount.ToString(System.Globalization.CultureInfo.InvariantCulture)
            });
        }
        return results.AsReadOnly();
    }

    private static void AddOptionalMinimum(List<QualityExpectationResult> results, string name, int? expected, int actual) {
        if (expected.HasValue) results.Add(Minimum(name, expected.Value, actual));
    }

    private static QualityExpectationResult Minimum(string name, int expected, int actual) => new() {
        Name = name,
        Succeeded = actual >= expected,
        Expected = ">= " + expected,
        Actual = actual.ToString(System.Globalization.CultureInfo.InvariantCulture)
    };

    private static QualityExpectationResult Exact<T>(string name, T expected, T actual) => new() {
        Name = name,
        Succeeded = EqualityComparer<T>.Default.Equals(expected, actual),
        Expected = expected?.ToString() ?? string.Empty,
        Actual = actual?.ToString() ?? string.Empty
    };

    private static QualityExpectationResult Sequence(string name, IReadOnlyList<string> expected, IReadOnlyList<string> actual) => new() {
        Name = name,
        Succeeded = expected.SequenceEqual(actual, StringComparer.Ordinal),
        Expected = string.Join(", ", expected),
        Actual = string.Join(", ", actual)
    };

    private static void Run(List<QualityCheckResult> checks, string name, Action action) {
        Stopwatch stopwatch = Stopwatch.StartNew();
        var check = new QualityCheckResult { Name = name };
        try {
            action();
            check.Succeeded = true;
        } catch (Exception exception) {
            check.ExceptionType = exception.GetType().FullName ?? exception.GetType().Name;
            check.Message = exception.Message;
        } finally {
            stopwatch.Stop();
            check.DurationMilliseconds = stopwatch.ElapsedMilliseconds;
            checks.Add(check);
        }
    }

    private static T Require<T>(T? value) where T : class =>
        value ?? throw new InvalidOperationException("A preceding PDF quality stage did not produce the required value.");

    private static byte[] ReadBounded(string path, long maxFileBytes) {
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (input.Length > maxFileBytes) throw new IOException("Corpus file exceeds the configured byte limit.");
        using var output = new MemoryStream((int)Math.Min(input.Length, 1024L * 1024L));
        var buffer = new byte[81920];
        long total = 0;
        while (true) {
            int read = input.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            total = checked(total + read);
            if (total > maxFileBytes) throw new IOException("Corpus file exceeds the configured byte limit.");
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }
}
