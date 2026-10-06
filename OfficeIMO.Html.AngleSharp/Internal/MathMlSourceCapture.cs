using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser.Tokens;
using AngleSharp.Text;
using OfficeIMO.Html.Dom;
using System.Threading;

namespace OfficeIMO.Html;

// Token callbacks belong to the retained parser: comments, raw text, quoted attributes and
// foreign-content CDATA are handled there. Only disjoint outer math ranges retain payloads,
// bounding total retained source characters by the caller's input length.
internal sealed class MathMlSourceCapture {
    private readonly string _source;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<int, int> _ends = new();
    private readonly Dictionary<IElement, int> _created = new();
    private readonly Dictionary<int, int> _parents = new();
    private readonly Stack<int> _starts = new();

    private MathMlSourceCapture(string source, CancellationToken cancellationToken) {
        _source = source;
        _cancellationToken = cancellationToken;
    }

    internal static MathMlSourceCapture? Configure(ref HtmlParserOptions options, string source, CancellationToken cancellationToken) {
        if (source.IndexOf("<math", StringComparison.OrdinalIgnoreCase) < 0) return null;
        var capture = new MathMlSourceCapture(source, cancellationToken);
        options.OnToken = capture.OnToken;
        options.OnCreated = capture.OnCreated;
        return capture;
    }

    private void OnCreated(IElement element, TextPosition position) {
        // The retained provider does not invoke OnCreated for an HTML-to-MathML root.
        // Its foreign descendants do carry creation positions. Associate those with the
        // token scope, then resolve their actual root after the parser's recovery completes.
        if (element.NamespaceUri == "http://www.w3.org/1998/Math/MathML") {
            int start = element.LocalName == "math" ? position.Index : _starts.Count == 0 ? -1 : _starts.Peek();
            if (start >= 0) _created[element] = start;
        }
    }

    private void OnToken(HtmlToken token, TextRange range) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (!string.Equals(token.Name, "math", StringComparison.OrdinalIgnoreCase)) return;
        if (token.Type == HtmlTokenType.StartTag) {
            _parents[range.Start.Index] = _starts.Count == 0 ? -1 : _starts.Peek();
            _starts.Push(range.Start.Index);
            if (token is HtmlTagToken tag && tag.IsSelfClosing) Close(range.End.Index);
        } else if (token.Type == HtmlTokenType.EndTag && _starts.Count != 0) Close(range.End.Index);
    }

    private void Close(int end) {
        int start = _starts.Pop();
        if (start >= 0 && end > start && end <= _source.Length) _ends[start] = end;
    }

    internal void Attach(IEnumerable<INode> nodes) {
        if (_created.Count == 0) return;
        var candidates = new Dictionary<IElement, int>();
        var pending = new Stack<(INode Node, IElement? Root, int Depth)>();
        foreach (INode node in nodes) pending.Push((node, null, 0));
        while (pending.Count != 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            var item = pending.Pop();
            IElement? root = item.Root;
            int depth = item.Depth;
            if (item.Node is IElement element) {
                if (element.LocalName == "math" && element.NamespaceUri == "http://www.w3.org/1998/Math/MathML") {
                    root ??= element;
                    depth++;
                }
                // The first foreign descendant identifies this root's token scope. Resolve
                // its lexical ancestors only once; deeply nested content must not cause a
                // repeated ancestry walk for every descendant before DOM limits are checked.
                if (root != null && !candidates.ContainsKey(root) && _created.TryGetValue(element, out int start)) {
                    for (int level = 1; level < depth && start >= 0; level++)
                        start = _parents.TryGetValue(start, out int parent) ? parent : -1;
                    if (start >= 0) candidates[root] = start;
                }
                if (element is IHtmlTemplateElement template) pending.Push((template.Content, null, 0));
            }
            for (int index = item.Node.ChildNodes.Length - 1; index >= 0; index--)
                pending.Push((item.Node.ChildNodes[index], root, depth));
        }
        var roots = candidates.Select(pair => (Element: pair.Key, Start: pair.Value)).ToList();
        roots.Sort((left, right) => left.Start.CompareTo(right.Start));
        for (int index = 0; index < roots.Count; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            var item = roots[index];
            // A recovered root's lexical range may include a later independent MathML root.
            // Such a range cannot describe its current subtree. Reject it before copying so
            // recovered markup cannot multiply retained payload size through overlapping ranges.
            int nextRoot = index + 1 < roots.Count ? roots[index + 1].Start : _source.Length;
            if (_ends.TryGetValue(item.Start, out int end) && end <= nextRoot)
                NativeSourceMarkup.Attach(item.Element, new HtmlSourceMarkupSnapshot(_source.Substring(item.Start, end - item.Start)));
        }
    }
}
