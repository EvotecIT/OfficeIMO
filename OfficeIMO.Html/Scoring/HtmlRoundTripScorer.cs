using System.Security.Cryptography;
using System.Globalization;

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
        AddCountMetric(metrics, "headings", target.Count(HtmlLogicalNodeKind.Heading), source.Count(HtmlLogicalNodeKind.Heading));
        AddSignatureMetric(metrics, "heading-levels", ExtractSignatures(target, HtmlLogicalNodeKind.Heading, CreateTextualNodeSignature), ExtractSignatures(source, HtmlLogicalNodeKind.Heading, CreateTextualNodeSignature));
        AddCountMetric(metrics, "paragraphs", target.Count(HtmlLogicalNodeKind.Paragraph), source.Count(HtmlLogicalNodeKind.Paragraph));
        AddCountMetric(metrics, "tables", target.Count(HtmlLogicalNodeKind.Table), source.Count(HtmlLogicalNodeKind.Table));
        AddCountMetric(metrics, "table-rows", target.Count(HtmlLogicalNodeKind.TableRow), source.Count(HtmlLogicalNodeKind.TableRow));
        AddCountMetric(metrics, "table-cells", target.Count(HtmlLogicalNodeKind.TableCell), source.Count(HtmlLogicalNodeKind.TableCell));
        AddSignatureMetric(metrics, "table-grid", ExtractTableGridSignatures(target), ExtractTableGridSignatures(source));
        AddCountMetric(metrics, "figures", target.Count(HtmlLogicalNodeKind.Figure), source.Count(HtmlLogicalNodeKind.Figure));
        AddSignatureMetric(metrics, "figure-signatures", ExtractSignatures(target, HtmlLogicalNodeKind.Figure, CreateFigureSignature), ExtractSignatures(source, HtmlLogicalNodeKind.Figure, CreateFigureSignature));
        AddCountMetric(metrics, "images", target.Count(HtmlLogicalNodeKind.Image), source.Count(HtmlLogicalNodeKind.Image));
        AddSignatureMetric(metrics, "image-sources", ExtractSignatures(target, HtmlLogicalNodeKind.Image, CreateImageSignature), ExtractSignatures(source, HtmlLogicalNodeKind.Image, CreateImageSignature));
        AddCountMetric(metrics, "lists", target.Count(HtmlLogicalNodeKind.List), source.Count(HtmlLogicalNodeKind.List));
        AddCountMetric(metrics, "list-items", target.Count(HtmlLogicalNodeKind.ListItem), source.Count(HtmlLogicalNodeKind.ListItem));
        AddCountMetric(metrics, "forms", target.Count(HtmlLogicalNodeKind.FormControl) + target.Count(HtmlLogicalNodeKind.Form), source.Count(HtmlLogicalNodeKind.FormControl) + source.Count(HtmlLogicalNodeKind.Form));
        AddSignatureMetric(metrics, "form-state", ExtractFormSignatures(target), ExtractFormSignatures(source));
        AddCountMetric(metrics, "links", target.Count(HtmlLogicalNodeKind.Link), source.Count(HtmlLogicalNodeKind.Link));
        AddSignatureMetric(metrics, "link-targets", ExtractSignatures(target, HtmlLogicalNodeKind.Link, CreateLinkSignature), ExtractSignatures(source, HtmlLogicalNodeKind.Link, CreateLinkSignature));
        AddMetric(metrics, "text", textSimilarity);

        int compared = metrics.Count;
        int matched = metrics.Values.Count(value => value >= 0.95D);
        double score = compared == 0 ? 1D : metrics.Values.Average();
        return new HtmlRoundTripScore(score, SumCounts(source), SumCounts(target), matched, compared, metrics);
    }

    private static void AddMetric(IDictionary<string, double> metrics, string name, double value) {
        metrics[name] = Math.Max(0D, Math.Min(1D, value));
    }

    private static void AddCountMetric(IDictionary<string, double> metrics, string name, int actual, int expected) {
        if (actual == 0 && expected == 0) {
            return;
        }

        AddMetric(metrics, name, Ratio(actual, expected));
    }

    private static void AddSignatureMetric(IDictionary<string, double> metrics, string name, IReadOnlyList<string> actual, IReadOnlyList<string> expected) {
        if (actual.Count == 0 && expected.Count == 0) {
            return;
        }

        AddMetric(metrics, name, SignatureSimilarity(actual, expected));
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

    private static IReadOnlyList<string> ExtractSignatures(HtmlLogicalDocument document, HtmlLogicalNodeKind kind, Func<HtmlLogicalNode, string> createSignature) {
        var signatures = new List<string>();
        AppendSignatures(document.Root, kind, createSignature, signatures);
        return signatures;
    }

    private static void AppendSignatures(HtmlLogicalNode node, HtmlLogicalNodeKind kind, Func<HtmlLogicalNode, string> createSignature, ICollection<string> signatures) {
        if (node.Kind == kind) {
            signatures.Add(createSignature(node));
        }

        foreach (HtmlLogicalNode child in node.Children) {
            AppendSignatures(child, kind, createSignature, signatures);
        }
    }

    private static IReadOnlyList<string> ExtractTableGridSignatures(HtmlLogicalDocument document) {
        var signatures = new List<string>();
        AppendTableGridSignatures(document.Root, signatures);
        return signatures;
    }

    private static void AppendTableGridSignatures(HtmlLogicalNode node, ICollection<string> signatures) {
        if (node.Kind == HtmlLogicalNodeKind.Table) {
            var rowSignatures = new List<string>();
            foreach (HtmlLogicalNode row in Descendants(node, HtmlLogicalNodeKind.TableRow)) {
                var cellSignatures = new List<string>();
                foreach (HtmlLogicalNode cell in Descendants(row, HtmlLogicalNodeKind.TableCell)) {
                    cellSignatures.Add(CreateTableCellGridSignature(cell));
                }

                rowSignatures.Add(string.Join("+", cellSignatures));
            }

            signatures.Add("table|" + string.Join(",", rowSignatures));
        }

        foreach (HtmlLogicalNode child in node.Children) {
            AppendTableGridSignatures(child, signatures);
        }
    }

    private static IEnumerable<HtmlLogicalNode> Descendants(HtmlLogicalNode node, HtmlLogicalNodeKind kind) {
        foreach (HtmlLogicalNode child in node.Children) {
            if (child.Kind == kind) {
                yield return child;
            }

            foreach (HtmlLogicalNode descendant in Descendants(child, kind)) {
                yield return descendant;
            }
        }
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

    private static string CreateTextualNodeSignature(HtmlLogicalNode node) {
        string text = ExtractLogicalNodeText(node);
        return string.IsNullOrWhiteSpace(text)
            ? node.Name
            : node.Name + "|text=" + NormalizeText(text);
    }

    private static string CreateLinkSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };
        AddAttributePart(parts, node, "href");
        string text = ExtractLogicalNodeText(node);
        if (!string.IsNullOrWhiteSpace(text)) {
            parts.Add("text=" + NormalizeText(text));
        }

        return string.Join("|", parts);
    }

    private static string CreateFigureSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };
        string text = ExtractLogicalNodeText(node);
        if (!string.IsNullOrWhiteSpace(text)) {
            parts.Add("text=" + NormalizeText(text));
        }

        foreach (HtmlLogicalNode image in Descendants(node, HtmlLogicalNodeKind.Image)) {
            parts.Add("image=" + CreateImageSignature(image));
        }

        return string.Join("|", parts);
    }

    private static string CreateImageSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };
        AddAttributePart(parts, node, "src");
        AddAttributePart(parts, node, "srcset");
        AddAttributePart(parts, node, "data-src");
        AddAttributePart(parts, node, "data-srcset");
        AddAttributePart(parts, node, "alt");
        return string.Join("|", parts);
    }

    private static string CreateTableCellGridSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            "cell"
        };
        AddAttributePart(parts, node, "colspan");
        AddAttributePart(parts, node, "rowspan");
        return string.Join("|", parts);
    }

    private static void AddAttributePart(ICollection<string> parts, HtmlLogicalNode node, string attributeName) {
        string? value;
        if (node.Attributes.TryGetValue(attributeName, out value)) {
            parts.Add(attributeName + "=" + value);
        }
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

        Dictionary<string, int> sourceWindows = HashWindows(sourceText);
        Dictionary<string, int> targetWindows = HashWindows(targetText);
        int unionCount = CountWindowUnion(sourceWindows, targetWindows);
        if (unionCount == 0) {
            return 1D;
        }

        return CountWindowIntersection(sourceWindows, targetWindows) / (double)unionCount;
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
        if (node is IElement element) {
            if (IsNonVisibleTextElement(element.TagName) || IsHiddenElement(element)) {
                return;
            }
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
        if (IsNonVisibleTextElement(node.Name) || IsHiddenLogicalNode(node)) {
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

    private static bool IsHiddenElement(IElement element) {
        if (element.HasAttribute("hidden")) {
            return true;
        }

        string? ariaHidden = element.GetAttribute("aria-hidden");
        if (string.Equals(ariaHidden, "true", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        return ContainsHiddenStyle(element.GetAttribute("style"));
    }

    private static bool IsHiddenLogicalNode(HtmlLogicalNode node) {
        if (node.Attributes.ContainsKey("hidden")) {
            return true;
        }

        string? ariaHidden;
        if (node.Attributes.TryGetValue("aria-hidden", out ariaHidden) && string.Equals(ariaHidden, "true", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        string? style;
        return node.Attributes.TryGetValue("style", out style) && ContainsHiddenStyle(style);
    }

    private static bool ContainsHiddenStyle(string? style) {
        if (string.IsNullOrWhiteSpace(style)) {
            return false;
        }

        string styleText = style!;
        return styleText.IndexOf("display:none", StringComparison.OrdinalIgnoreCase) >= 0
            || styleText.IndexOf("display: none", StringComparison.OrdinalIgnoreCase) >= 0
            || styleText.IndexOf("visibility:hidden", StringComparison.OrdinalIgnoreCase) >= 0
            || styleText.IndexOf("visibility: hidden", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static Dictionary<string, int> HashWindows(string text) {
        var windows = new Dictionary<string, int>(StringComparer.Ordinal);
        if (text.Length <= 32) {
            AddWindow(windows, Hash(text));
            return windows;
        }

        for (int i = 0; i <= text.Length - 32; i += 16) {
            AddWindow(windows, Hash(text.Substring(i, 32)));
        }

        AddWindow(windows, Hash(text.Substring(text.Length - 32, 32)));
        return windows;
    }

    private static void AddWindow(IDictionary<string, int> windows, string hash) {
        int count;
        windows.TryGetValue(hash, out count);
        windows[hash] = count + 1;
    }

    private static int CountWindowIntersection(IReadOnlyDictionary<string, int> source, IReadOnlyDictionary<string, int> target) {
        int count = 0;
        foreach (KeyValuePair<string, int> pair in source) {
            int targetCount;
            if (target.TryGetValue(pair.Key, out targetCount)) {
                count += Math.Min(pair.Value, targetCount);
            }
        }

        return count;
    }

    private static int CountWindowUnion(IReadOnlyDictionary<string, int> source, IReadOnlyDictionary<string, int> target) {
        int count = 0;
        var keys = new HashSet<string>(source.Keys, StringComparer.Ordinal);
        keys.UnionWith(target.Keys);
        foreach (string key in keys) {
            int sourceCount;
            int targetCount;
            source.TryGetValue(key, out sourceCount);
            target.TryGetValue(key, out targetCount);
            count += Math.Max(sourceCount, targetCount);
        }

        return count;
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
