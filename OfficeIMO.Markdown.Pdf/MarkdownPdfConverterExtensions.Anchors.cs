using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

public static partial class MarkdownPdfConverterExtensions {
    /// <summary>Indexes rendered destinations once per export and diagnoses missing fragment links.</summary>
    internal sealed class AnchorContext {
        private readonly Dictionary<IMarkdownBlock, List<string>> _byBlock = new Dictionary<IMarkdownBlock, List<string>>();
        private readonly HashSet<string> _known = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _registered = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _missing = new HashSet<string>(StringComparer.Ordinal);
        private readonly MarkdownToPdfOptions _options;

        internal AnchorContext(MarkdownDoc document, IReadOnlyList<IMarkdownBlock> blocks, MarkdownToPdfOptions options) {
            _options = options;
            foreach (IMarkdownBlock block in blocks) {
                if (!(block is MarkdownObject root)) continue;
                var stack = new Stack<(MarkdownObject Node, IMarkdownBlock Owner, bool Flattened)>();
                stack.Push((root, block, false));
                while (stack.Count > 0) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    var entry = stack.Pop();
                    MarkdownObject node = entry.Node;
                    if (node is HtmlCommentBlock || (node is FrontMatterBlock && options.FrontMatterRenderMode == MarkdownPdfFrontMatterRenderMode.Hidden)) continue;
                    IMarkdownBlock owner = entry.Owner;
                    bool renderedWithOwner = node is SummaryBlock && owner is DetailsBlock ||
                        owner is FootnoteDefinitionBlock footnote && footnote.ChildBlocks.Count > 0 && ReferenceEquals(node, footnote.ChildBlocks[0]);
                    if (!entry.Flattened && node is IMarkdownBlock childBlock &&
                        !renderedWithOwner && !(node is ParagraphBlock && node.Parent is ListItem)) owner = childBlock;
                    string? identifier = node.Attributes.ElementId;
                    if (node is HeadingBlock heading) Add(owner, document.GetHeadingAnchor(heading));
                    Add(owner, identifier);
                    if (node is HtmlRawInline html && TryGetInlineAnchor(html.Html, out string? inlineAnchor)) Add(owner, inlineAnchor);
                    bool flattened = entry.Flattened || node is TableBlock || node is DefinitionListBlock;
                    IReadOnlyList<MarkdownObject> children = node.ChildObjects;
                    for (int index = children.Count - 1; index >= 0; index--) stack.Push((children[index], owner, flattened));
                }
            }
        }

        private void Add(IMarkdownBlock owner, string? name) {
            if (string.IsNullOrWhiteSpace(name)) return;
            if (!_byBlock.TryGetValue(owner, out List<string>? names)) {
                names = new List<string>();
                _byBlock.Add(owner, names);
            }
            names.Add(name!);
            _known.Add(name!);
        }

        internal void Register(PdfCore.PdfDocument pdf, IMarkdownBlock block) {
            if (!_byBlock.TryGetValue(block, out List<string>? names)) return;
            foreach (string name in names) RegisterNamed(pdf, name);
        }

        internal void RegisterNamed(PdfCore.PdfDocument pdf, string name) {
            if (!string.IsNullOrWhiteSpace(name) && _registered.Add(name)) pdf.Bookmark(name);
        }

        internal bool CanLink(string name) {
            if (_known.Contains(name)) return true;
            if (_missing.Add(name)) AddWarning(_options, "UnresolvedInternalLink", name,
                "The internal link target is absent from rendered content; its visible text was retained without a link.");
            return false;
        }

        // Recognize opening or empty anchor carriers used by supported semantic producers. Other raw HTML
        // remains under the renderer's existing plain-text fallback policy.
        private static bool TryGetInlineAnchor(string html, out string? name) {
            name = null;
            string value = html.Trim();
            foreach (string opening in new[] { "<a id=\"", "<a id='" }) {
                string quote = opening.EndsWith("\"", StringComparison.Ordinal) ? "\"" : "'";
                string ending = value.EndsWith("</a>", StringComparison.OrdinalIgnoreCase) ? quote + "></a>" : quote + ">";
                if (value.Length >= opening.Length + ending.Length && value.StartsWith(opening, StringComparison.OrdinalIgnoreCase) && value.EndsWith(ending, StringComparison.OrdinalIgnoreCase)) {
                    name = System.Net.WebUtility.HtmlDecode(value.Substring(opening.Length, value.Length - opening.Length - ending.Length));
                    return !string.IsNullOrWhiteSpace(name);
                }
            }
            return false;
        }
    }
}
