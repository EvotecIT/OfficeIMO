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
        return Compare(source, target, TextSimilarityFromText(ExtractVisibleTextFromHtml(sourceHtml), ExtractVisibleTextFromHtml(targetHtml)));
    }

    /// <summary>
    /// Compares logical documents and returns a structural score.
    /// </summary>
    public static HtmlRoundTripScore Compare(HtmlLogicalDocument source, HtmlLogicalDocument target) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }

        if (target == null) {
            throw new ArgumentNullException(nameof(target));
        }

        return Compare(source, target, TextSimilarityFromText(ExtractLogicalText(source), ExtractLogicalText(target)));
    }

    private static HtmlRoundTripScore Compare(HtmlLogicalDocument source, HtmlLogicalDocument target, double textSimilarity) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }

        if (target == null) {
            throw new ArgumentNullException(nameof(target));
        }

        var metrics = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        AddMetric(metrics, "nodes", Ratio(SumCounts(target), SumCounts(source)));
        AddMetric(metrics, "headings", Ratio(target.Count(HtmlLogicalNodeKind.Heading), source.Count(HtmlLogicalNodeKind.Heading)));
        AddMetric(metrics, "paragraphs", Ratio(target.Count(HtmlLogicalNodeKind.Paragraph), source.Count(HtmlLogicalNodeKind.Paragraph)));
        AddMetric(metrics, "tables", Ratio(target.Count(HtmlLogicalNodeKind.Table), source.Count(HtmlLogicalNodeKind.Table)));
        AddMetric(metrics, "images", Ratio(target.Count(HtmlLogicalNodeKind.Image), source.Count(HtmlLogicalNodeKind.Image)));
        AddMetric(metrics, "lists", Ratio(target.Count(HtmlLogicalNodeKind.List), source.Count(HtmlLogicalNodeKind.List)));
        AddMetric(metrics, "list-items", Ratio(target.Count(HtmlLogicalNodeKind.ListItem), source.Count(HtmlLogicalNodeKind.ListItem)));
        AddMetric(metrics, "forms", Ratio(target.Count(HtmlLogicalNodeKind.FormControl) + target.Count(HtmlLogicalNodeKind.Form), source.Count(HtmlLogicalNodeKind.FormControl) + source.Count(HtmlLogicalNodeKind.Form)));
        AddMetric(metrics, "form-state", SignatureSimilarity(ExtractFormSignatures(target), ExtractFormSignatures(source)));
        AddMetric(metrics, "links", Ratio(target.Count(HtmlLogicalNodeKind.Link), source.Count(HtmlLogicalNodeKind.Link)));
        AddMetric(metrics, "text", textSimilarity);

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

    private static double SignatureSimilarity(IReadOnlyList<string> actual, IReadOnlyList<string> expected) {
        if (expected.Count == 0) {
            return actual.Count == 0 ? 1D : 0D;
        }

        var remaining = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string signature in expected) {
            if (!remaining.ContainsKey(signature)) {
                remaining[signature] = 0;
            }

            remaining[signature]++;
        }

        int matched = 0;
        foreach (string signature in actual) {
            int count;
            if (remaining.TryGetValue(signature, out count) && count > 0) {
                remaining[signature] = count - 1;
                matched++;
            }
        }

        return matched / (double)Math.Max(actual.Count, expected.Count);
    }

    private static IReadOnlyList<string> ExtractFormSignatures(HtmlLogicalDocument document) {
        var signatures = new List<string>();
        AppendFormSignatures(document.Root, signatures);
        return signatures;
    }

    private static void AppendFormSignatures(HtmlLogicalNode node, ICollection<string> signatures) {
        if (node.Kind == HtmlLogicalNodeKind.FormControl) {
            signatures.Add(CreateFormSignature(node));
        }

        foreach (HtmlLogicalNode child in node.Children) {
            AppendFormSignatures(child, signatures);
        }
    }

    private static string CreateFormSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };

        foreach (string attributeName in new[] { "type", "name", "value", "checked", "selected", "disabled", "multiple", "placeholder" }) {
            string? value;
            if (node.Attributes.TryGetValue(attributeName, out value)) {
                parts.Add(attributeName + "=" + value);
            }
        }

        string text = ExtractLogicalNodeText(node);
        if (!string.IsNullOrWhiteSpace(text)) {
            parts.Add("text=" + NormalizeText(text));
        }

        return string.Join("|", parts);
    }

    private static string ExtractLogicalNodeText(HtmlLogicalNode node) {
        var parts = new List<string>();
        AppendLogicalText(node, parts);
        return string.Join(" ", parts);
    }

    private static double TextSimilarityFromText(string sourceText, string targetText) {
        sourceText = NormalizeText(sourceText);
        targetText = NormalizeText(targetText);
        if (sourceText.Length == 0 && targetText.Length == 0) {
            return 1D;
        }

        if (string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
            return 1D;
        }

        HashSet<string> sourceWindows = HashWindows(sourceText);
        HashSet<string> targetWindows = HashWindows(targetText);
        int unionCount = sourceWindows.Union(targetWindows).Count();
        if (unionCount == 0) {
            return 1D;
        }

        return sourceWindows.Intersect(targetWindows).Count() / (double)unionCount;
    }

    private static string ExtractVisibleTextFromHtml(string html) {
        if (string.IsNullOrWhiteSpace(html)) {
            return string.Empty;
        }

        var parts = new List<string>();
        var document = HtmlDocumentParser.ParseDocument(html);
        INode? root = document.Body ?? (INode?)document.DocumentElement;
        if (root != null) {
            AppendVisibleText(root, parts);
        }

        return string.Join(" ", parts);
    }

    private static void AppendVisibleText(INode node, ICollection<string> parts) {
        if (node is IElement element && IsNonVisibleTextElement(element.TagName)) {
            return;
        }

        if (node.NodeType == NodeType.Text && !string.IsNullOrWhiteSpace(node.TextContent)) {
            parts.Add(node.TextContent);
            return;
        }

        foreach (INode child in node.ChildNodes) {
            AppendVisibleText(child, parts);
        }
    }

    private static string ExtractLogicalText(HtmlLogicalDocument document) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        var parts = new List<string>();
        AppendLogicalText(document.Root, parts);
        return string.Join(" ", parts);
    }

    private static void AppendLogicalText(HtmlLogicalNode node, ICollection<string> parts) {
        if (IsNonVisibleTextElement(node.Name)) {
            return;
        }

        if (!string.IsNullOrWhiteSpace(node.Text) && (node.Kind == HtmlLogicalNodeKind.Text || !HasTextChild(node))) {
            parts.Add(node.Text);
        }

        foreach (HtmlLogicalNode child in node.Children) {
            AppendLogicalText(child, parts);
        }
    }

    private static bool HasTextChild(HtmlLogicalNode node) {
        foreach (HtmlLogicalNode child in node.Children) {
            if (!string.IsNullOrWhiteSpace(child.Text) || HasTextChild(child)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsNonVisibleTextElement(string name) {
        return string.Equals(name, "script", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "style", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "template", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "noscript", StringComparison.OrdinalIgnoreCase);
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

        windows.Add(Hash(text.Substring(text.Length - 32, 32)));
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
