using OfficeIMO.Html;
using OfficeIMO.Tests;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class HtmlCorpusAcceptanceEvaluator {
    private const string ScreenIntent = "screen-full-page-v1";
    private const string PrintIntent = "print-paged-v1";
    private const string ScreenToPageIntent = "screen-snapshot-paged-v1";

    internal static HtmlCorpusAcceptanceEvidence Evaluate(
        HtmlRenderingAdvancedHeldOutCorpus corpus,
        IReadOnlyList<HtmlCorpusCaseEvidence> evidence) {
        Dictionary<string, HtmlCorpusCaseEvidence> evidenceById = evidence.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var cases = new List<HtmlCorpusCaseAcceptance>(corpus.Acceptance.Cases.Count);
        var failures = new List<string>();
        foreach (HtmlRenderingVisualAcceptanceCase policy in corpus.Acceptance.Cases) {
            if (!evidenceById.TryGetValue(policy.Id, out HtmlCorpusCaseEvidence? item)) {
                HtmlCorpusCaseAcceptance missing = MissingCase(policy);
                cases.Add(missing);
                failures.Add(policy.Id + " has no captured evidence.");
                continue;
            }
            HtmlCorpusIntentAcceptance screen = EvaluateScreen(item, policy.Screen);
            HtmlCorpusIntentAcceptance print = EvaluatePrint(item, policy.Print);
            HtmlCorpusIntentAcceptance screenToPage = EvaluateScreenToPage(item, policy.ScreenToPage);
            var result = new HtmlCorpusCaseAcceptance(
                policy.Id, screen.Passed && print.Passed && screenToPage.Passed,
                screen, print, screenToPage);
            cases.Add(result);
            foreach (HtmlCorpusIntentAcceptance intent in new[] { screen, print, screenToPage }) {
                foreach (HtmlCorpusAcceptanceCriterion criterion in intent.Criteria.Where(value => !value.Passed)) {
                    failures.Add(policy.Id + " " + intent.Intent + " failed " + criterion.Id + ": "
                        + (criterion.Detail ?? FormatCriterion(criterion)) + ".");
                }
            }
        }

        IReadOnlyList<HtmlCorpusCapabilityAcceptance> capabilities = EvaluateCapabilities(cases);
        foreach (HtmlCorpusCapabilityAcceptance capability in capabilities.Where(item => !item.Passed)) {
            failures.Add(capability.ProfileId + "/" + capability.CapabilityId
                + " failed selected cases: " + string.Join(", ", capability.FailedCaseIds) + ".");
        }
        return new HtmlCorpusAcceptanceEvidence(
            corpus.AcceptanceSha256,
            corpus.Acceptance.ReferencePolicy,
            failures.Count == 0,
            cases.AsReadOnly(),
            capabilities,
            failures.AsReadOnly());
    }

    private static HtmlCorpusIntentAcceptance EvaluateScreen(
        HtmlCorpusCaseEvidence item,
        HtmlRenderingScreenAcceptance policy) {
        HtmlCorpusGeometryComparison? geometry = item.Comparisons.ScreenGeometry;
        HtmlCorpusPixelComparison? pixels = item.Comparisons.ScreenPixels;
        if (item.OfficeImo == null || item.Browser == null || geometry == null || pixels == null) {
            return MissingIntent(ScreenIntent, "Chromium screen screenshot and geometry", policy.Classification, policy.Rationale);
        }
        double geometryMatch = Math.Max(geometry.OfficeImoElementCount, geometry.BrowserElementCount) == 0
            ? 1D
            : geometry.MatchedElementCount / (double)Math.Max(geometry.OfficeImoElementCount, geometry.BrowserElementCount);
        var criteria = new List<HtmlCorpusAcceptanceCriterion> {
            Required("dimensions-match", pixels.DimensionsMatch, policy.RequireDimensionsMatch),
            Maximum("width-difference-pixels", Math.Abs(pixels.ExpectedWidth - pixels.ActualWidth), policy.MaximumWidthDifferencePixels),
            Maximum("height-difference-pixels", Math.Abs(pixels.ExpectedHeight - pixels.ActualHeight), policy.MaximumHeightDifferencePixels),
            Minimum("geometry-match-ratio", geometryMatch, policy.MinimumGeometryMatchRatio),
            MaximumNullable("geometry-mean-absolute-x", geometry.MeanAbsoluteX, policy.MaximumMeanAbsoluteX),
            MaximumNullable("geometry-mean-absolute-y", geometry.MeanAbsoluteY, policy.MaximumMeanAbsoluteY),
            MaximumNullable("geometry-mean-absolute-width", geometry.MeanAbsoluteWidth, policy.MaximumMeanAbsoluteWidth),
            MaximumNullable("geometry-mean-absolute-height", geometry.MeanAbsoluteHeight, policy.MaximumMeanAbsoluteHeight),
            MaximumNullable("pixel-mean-absolute-error", pixels.MeanAbsoluteError, policy.MaximumPixelMeanAbsoluteError),
            MaximumNullable("pixel-root-mean-square-error", pixels.RootMeanSquareError, policy.MaximumPixelRootMeanSquareError),
            MaximumNullable("pixel-mean-luminance-error", pixels.MeanLuminanceError, policy.MaximumPixelMeanLuminanceError)
        };
        return Intent(ScreenIntent, "Chromium screen screenshot and geometry", policy.Classification, policy.Rationale, criteria);
    }

    private static HtmlCorpusIntentAcceptance EvaluatePrint(
        HtmlCorpusCaseEvidence item,
        HtmlRenderingPrintAcceptance policy) {
        HtmlCorpusTextComparison? text = item.Comparisons.OfficeImoPrintToChromiumPrint;
        HtmlCorpusPageComparison[] pages = item.Comparisons.PrintPagesToChromium.ToArray();
        if (item.OfficeImo == null || item.Browser == null || text == null || pages.Length == 0) {
            return MissingIntent(PrintIntent, "Chromium print-to-PDF", policy.Classification, policy.Rationale);
        }
        HtmlCorpusPixelComparison[] pixels = pages.Where(page => page.Pixels != null).Select(page => page.Pixels!).ToArray();
        bool pageCountsMatch = item.OfficeImo.PrintPdf.PageCount == item.Browser.PrintPdf.PageCount
            && pages.All(page => page.PresentInBoth);
        bool alignmentMatches = pixels.Length == pages.Length
            && pixels.All(pixel => MatchesAlignment(pixel.Alignment, policy.PixelAlignment));
        var criteria = new List<HtmlCorpusAcceptanceCriterion> {
            Required("page-count-match", pageCountsMatch, policy.RequirePageCountMatch),
            Required("pixel-alignment", alignmentMatches, required: true, policy.PixelAlignment),
            Minimum("text-token-recall", text.TokenRecall, policy.MinimumTextRecall),
            Minimum("text-token-precision", text.TokenPrecision, policy.MinimumTextPrecision),
            MaximumNullable("page-width-difference-pixels", MaximumOrNull(pixels.Select(pixel => (double?)Math.Abs(pixel.ExpectedWidth - pixel.ActualWidth))), policy.MaximumWidthDifferencePixels),
            MaximumNullable("page-height-difference-pixels", MaximumOrNull(pixels.Select(pixel => (double?)Math.Abs(pixel.ExpectedHeight - pixel.ActualHeight))), policy.MaximumHeightDifferencePixels),
            MaximumNullable("pixel-mean-absolute-error", MaximumOrNull(pixels.Select(pixel => pixel.MeanAbsoluteError)), policy.MaximumPixelMeanAbsoluteError),
            MaximumNullable("pixel-root-mean-square-error", MaximumOrNull(pixels.Select(pixel => pixel.RootMeanSquareError)), policy.MaximumPixelRootMeanSquareError),
            MaximumNullable("pixel-mean-luminance-error", MaximumOrNull(pixels.Select(pixel => pixel.MeanLuminanceError)), policy.MaximumPixelMeanLuminanceError)
        };
        return Intent(PrintIntent, "Chromium print-to-PDF", policy.Classification, policy.Rationale, criteria);
    }

    private static HtmlCorpusIntentAcceptance EvaluateScreenToPage(
        HtmlCorpusCaseEvidence item,
        HtmlRenderingScreenToPageAcceptance policy) {
        HtmlCorpusScreenToPageComparison? comparison = item.Comparisons.ScreenToPage;
        if (item.OfficeImo == null || comparison == null) {
            return MissingIntent(ScreenToPageIntent, "OfficeIMO continuous screen display list", policy.Classification, policy.Rationale);
        }
        var criteria = new List<HtmlCorpusAcceptanceCriterion> {
            Required("covers-screen-height", comparison.CoversScreenHeight, policy.RequireCoversScreenHeight),
            Maximum("clipped-width-pixels", comparison.ClippedWidth, policy.MaximumClippedWidthPixels),
            Maximum("trailing-height-pixels", comparison.TrailingHeight, policy.MaximumTrailingHeightPixels),
            Maximum("pixel-mean-absolute-error", comparison.MeanAbsoluteError, policy.MaximumPixelMeanAbsoluteError),
            Maximum("pixel-root-mean-square-error", comparison.RootMeanSquareError, policy.MaximumPixelRootMeanSquareError),
            Maximum("pixel-mean-luminance-error", comparison.MeanLuminanceError, policy.MaximumPixelMeanLuminanceError)
        };
        return Intent(ScreenToPageIntent, "OfficeIMO continuous screen display list", policy.Classification, policy.Rationale, criteria);
    }

    private static IReadOnlyList<HtmlCorpusCapabilityAcceptance> EvaluateCapabilities(
        IReadOnlyList<HtmlCorpusCaseAcceptance> cases) {
        var byId = cases.ToDictionary(item => item.Id, StringComparer.OrdinalIgnoreCase);
        var results = new List<HtmlCorpusCapabilityAcceptance>();
        AddCapabilities(
            HtmlCapabilityProfileIds.StaticScreenV1,
            HtmlCapabilityEvidenceIds.H4ScreenAdvancedHeldOut,
            new[] { ScreenIntent },
            result => result.Screen.Passed,
            byId,
            results);
        AddCapabilities(
            HtmlCapabilityProfileIds.PagedPrintV1,
            HtmlCapabilityEvidenceIds.H4PagedAdvancedHeldOut,
            new[] { PrintIntent, ScreenToPageIntent },
            result => result.Print.Passed && result.ScreenToPage.Passed,
            byId,
            results);
        return results.OrderBy(item => item.ProfileId, StringComparer.Ordinal)
            .ThenBy(item => item.CapabilityId, StringComparer.Ordinal).ToArray();
    }

    private static void AddCapabilities(
        string profileId,
        string evidenceId,
        IReadOnlyList<string> intents,
        Func<HtmlCorpusCaseAcceptance, bool> passed,
        IReadOnlyDictionary<string, HtmlCorpusCaseAcceptance> cases,
        ICollection<HtmlCorpusCapabilityAcceptance> results) {
        HtmlCapabilityEvidencePin evidence = HtmlRenderCapabilityCatalog.GetProfile(profileId).Evidence
            .Single(item => string.Equals(item.Id, evidenceId, StringComparison.OrdinalIgnoreCase));
        foreach (HtmlCapabilityEvidenceSelection selection in evidence.Selections) {
            string[] failed = selection.RequiredCaseIds
                .Where(id => !cases.TryGetValue(id, out HtmlCorpusCaseAcceptance? item) || !passed(item))
                .OrderBy(id => id, StringComparer.Ordinal).ToArray();
            results.Add(new HtmlCorpusCapabilityAcceptance(
                profileId,
                selection.CapabilityId,
                intents,
                selection.RequiredCaseIds,
                failed.Length == 0,
                failed));
        }
    }

    private static HtmlCorpusCaseAcceptance MissingCase(HtmlRenderingVisualAcceptanceCase policy) {
        HtmlCorpusIntentAcceptance screen = MissingIntent(ScreenIntent, "Chromium screen screenshot and geometry", policy.Screen.Classification, policy.Screen.Rationale);
        HtmlCorpusIntentAcceptance print = MissingIntent(PrintIntent, "Chromium print-to-PDF", policy.Print.Classification, policy.Print.Rationale);
        HtmlCorpusIntentAcceptance screenToPage = MissingIntent(ScreenToPageIntent, "OfficeIMO continuous screen display list", policy.ScreenToPage.Classification, policy.ScreenToPage.Rationale);
        return new HtmlCorpusCaseAcceptance(policy.Id, false, screen, print, screenToPage);
    }

    private static HtmlCorpusIntentAcceptance MissingIntent(
        string intent,
        string reference,
        string classification,
        string rationale) => new(
            intent, reference, classification, rationale, false,
            new[] { new HtmlCorpusAcceptanceCriterion("evidence-present", 0D, "required", 1D, false, "required evidence was not captured") });

    private static HtmlCorpusIntentAcceptance Intent(
        string intent,
        string reference,
        string classification,
        string rationale,
        IReadOnlyList<HtmlCorpusAcceptanceCriterion> criteria) =>
        new(intent, reference, classification, rationale, criteria.All(item => item.Passed), criteria);

    private static HtmlCorpusAcceptanceCriterion Maximum(string id, double actual, double threshold) =>
        new(id, actual, "less-than-or-equal", threshold, actual <= threshold);

    private static HtmlCorpusAcceptanceCriterion MaximumNullable(string id, double? actual, double threshold) =>
        actual.HasValue
            ? Maximum(id, actual.Value, threshold)
            : new HtmlCorpusAcceptanceCriterion(id, null, "less-than-or-equal", threshold, false, "metric was unavailable");

    private static HtmlCorpusAcceptanceCriterion Minimum(string id, double actual, double threshold) =>
        new(id, actual, "greater-than-or-equal", threshold, actual >= threshold);

    private static HtmlCorpusAcceptanceCriterion Required(string id, bool actual, bool required, string? detail = null) =>
        new(id, actual ? 1D : 0D, required ? "required-true" : "observed", required ? 1D : null, !required || actual,
            detail == null ? null : "expected " + detail);

    private static double? MaximumOrNull(IEnumerable<double?> values) {
        double[] present = values.Where(value => value.HasValue).Select(value => value!.Value).ToArray();
        return present.Length == 0 ? null : present.Max();
    }

    private static bool MatchesAlignment(string actual, string policy) => policy switch {
        "exact-or-nearest-resize" => actual is "exact" or "nearest-resize",
        _ => string.Equals(actual, policy, StringComparison.Ordinal)
    };

    private static string FormatCriterion(HtmlCorpusAcceptanceCriterion criterion) =>
        "observed " + (criterion.Actual?.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) ?? "unavailable")
        + " " + criterion.Comparison + " "
        + (criterion.Threshold?.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) ?? "n/a");
}
