using System.Security.Cryptography;

namespace OfficeIMO.Html;

/// <summary>
/// Scores structural HTML round-trip fidelity for gallery manifests and regression tests.
/// </summary>
public static class HtmlRoundTripScorer {
    private static readonly char[] WhitespaceSeparators = { ' ', '\t', '\r', '\n', '\f' };

    /// <summary>
    /// Compares source HTML with target HTML and returns a structural score.
    /// </summary>
    public static HtmlRoundTripScore Compare(string sourceHtml, string targetHtml) {
        HtmlLogicalDocument source = HtmlLogicalDocumentBuilder.FromHtml(sourceHtml);
        HtmlLogicalDocument target = HtmlLogicalDocumentBuilder.FromHtml(targetHtml);
        return Compare(source, target, sourceHtml, targetHtml);
    }

    /// <summary>
    /// Compares logical documents and returns a structural score.
    /// </summary>
    public static HtmlRoundTripScore Compare(HtmlLogicalDocument source, HtmlLogicalDocument target) {
        return Compare(source, target, string.Empty, string.Empty);
    }

    private static HtmlRoundTripScore Compare(HtmlLogicalDocument source, HtmlLogicalDocument target, string sourceHtml, string targetHtml) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }

        if (target == null) {
            throw new ArgumentNullException(nameof(target));
        }

        var metrics = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        AddMetric(metrics, "nodes", Ratio(SumCounts(target), SumCounts(source)));
        AddMetric(metrics, "headings", Ratio(target.Count(HtmlLogicalNodeKind.Heading), source.Count(HtmlLogicalNodeKind.Heading)));
        AddMetric(metrics, "tables", Ratio(target.Count(HtmlLogicalNodeKind.Table), source.Count(HtmlLogicalNodeKind.Table)));
        AddMetric(metrics, "images", Ratio(target.Count(HtmlLogicalNodeKind.Image), source.Count(HtmlLogicalNodeKind.Image)));
        AddMetric(metrics, "forms", Ratio(target.Count(HtmlLogicalNodeKind.FormControl) + target.Count(HtmlLogicalNodeKind.Form), source.Count(HtmlLogicalNodeKind.FormControl) + source.Count(HtmlLogicalNodeKind.Form)));
        AddMetric(metrics, "links", Ratio(target.Count(HtmlLogicalNodeKind.Link), source.Count(HtmlLogicalNodeKind.Link)));
        AddMetric(metrics, "text", TextSimilarity(sourceHtml, targetHtml));

        int compared = metrics.Count;
        int matched = metrics.Values.Count(value => value >= 0.95D);
        double score = compared == 0 ? 1D : metrics.Values.Average();
        return new HtmlRoundTripScore(score, SumCounts(source), SumCounts(target), matched, compared, metrics);
    }

    private static void AddMetric(IDictionary<string, double> metrics, string name, double value) {
        metrics[name] = Math.Max(0D, Math.Min(1D, value));
    }

    private static int SumCounts(HtmlLogicalDocument document) {
        return document.GetCounts().Values.Sum();
    }

    private static double Ratio(int actual, int expected) {
        if (expected == 0) {
            return actual == 0 ? 1D : 0D;
        }

        return Math.Min(actual, expected) / (double)Math.Max(actual, expected);
    }

    private static double TextSimilarity(string sourceHtml, string targetHtml) {
        if (string.IsNullOrWhiteSpace(sourceHtml) && string.IsNullOrWhiteSpace(targetHtml)) {
            return 1D;
        }

        string sourceText = NormalizeText(HtmlDocumentParser.ParseDocument(sourceHtml).Body?.TextContent ?? sourceHtml);
        string targetText = NormalizeText(HtmlDocumentParser.ParseDocument(targetHtml).Body?.TextContent ?? targetHtml);
        if (sourceText.Length == 0 && targetText.Length == 0) {
            return 1D;
        }

        if (string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
            return 1D;
        }

        return Ratio(HashWindows(sourceText).Intersect(HashWindows(targetText)).Count(), HashWindows(sourceText).Count);
    }

    private static HashSet<string> HashWindows(string text) {
        var windows = new HashSet<string>(StringComparer.Ordinal);
        if (text.Length <= 32) {
            windows.Add(Hash(text));
            return windows;
        }

        for (int i = 0; i <= text.Length - 32; i += 16) {
            windows.Add(Hash(text.Substring(i, 32)));
        }

        return windows;
    }

    private static string NormalizeText(string text) {
        return string.IsNullOrWhiteSpace(text)
            ? string.Empty
            : string.Join(" ", text.Split(WhitespaceSeparators, StringSplitOptions.RemoveEmptyEntries));
    }

    private static string Hash(string value) {
        using (SHA256 sha = SHA256.Create()) {
            byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value));
            return Convert.ToBase64String(bytes);
        }
    }
}
