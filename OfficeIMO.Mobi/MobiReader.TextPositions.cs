using System.Text;
using System.Xml.Linq;
using AngleSharp.Dom;

namespace OfficeIMO.Mobi;

public static partial class MobiReader {
    private readonly record struct MarkupTextRun(int Start, int End, string Text);

    private sealed partial class MarkupPart {
        private readonly Dictionary<IText, List<(int Offset, string Id)>> _textAnchors = new();
        private int _textAnchorCount;

        internal bool BindPositions(HashSet<int> targets, byte[] bytes, Encoding encoding, string markup,
            IReadOnlyList<MarkupTextRun> runs, CancellationToken token) {
            if (targets.Count == 0) return false;
            if (targets.Count == 1 && targets.Contains(0)) {
                _positions.Add(0, TopAnchor);
                return false;
            }

            var elements = _elements.Where(element => element.SourceReference?.Position.Index >= 0)
                .OrderBy(element => element.SourceReference!.Position.Index).ToArray();
            var mapped = MatchTextRuns(markup, runs, token);
            int elementIndex = 0, runIndex = 0, previous = 0, characters = 0;
            int closerIndex = 0, inserted = 0;
            bool approximated = false;
            foreach (int target in targets.OrderBy(item => item)) {
                token.ThrowIfCancellationRequested();
                if (target == 0) {
                    _positions.Add(target, TopAnchor);
                    continue;
                }
                int end = target;
                if (encoding.CodePage == 65001)
                    while (end > previous && end < bytes.Length && (bytes[end] & 0xc0) == 0x80) end--;
                characters += encoding.GetCharCount(bytes, previous, end - previous);
                previous = end;
                while (_insertedClosers != null && closerIndex < _insertedClosers.Count &&
                    _insertedClosers[closerIndex].Offset <= characters)
                    inserted += _insertedClosers[closerIndex++].Length;
                int position = characters + inserted;
                while (runIndex < mapped.Count && mapped[runIndex].Range.End <= position) runIndex++;
                if (runIndex < mapped.Count && mapped[runIndex].Range.Start <= position) {
                    var run = mapped[runIndex];
                    int offset = position - run.Range.Start;
                    if (!run.CharacterOffsetsMatch && offset != 0) {
                        offset = 0;
                        approximated = true;
                    }
                    _positions.Add(target, AnchorText(run.Node, offset));
                    continue;
                }
                while (elementIndex + 1 < elements.Length && elements[elementIndex + 1].SourceReference!.Position.Index <= position + 1)
                    elementIndex++;
                IElement element = elements.Length == 0 ? Body : elements[elementIndex];
                if (element.SourceReference?.Position.Index != position && element.SourceReference?.Position.Index != position + 1)
                    approximated = true;
                _positions.Add(target, EnsureAnchor(element));
            }
            return approximated;
        }

        private List<(MarkupTextRun Range, IText Node, bool CharacterOffsetsMatch)> MatchTextRuns(
            string markup, IReadOnlyList<MarkupTextRun> ranges, CancellationToken token) {
            var matched = new List<(MarkupTextRun, IText, bool)>();
            var pending = new Stack<INode>();
            pending.Push(Body);
            int cursor = 0;
            while (cursor < ranges.Count && pending.Count > 0) {
                token.ThrowIfCancellationRequested();
                INode node = pending.Pop();
                if (node is IElement element) {
                    if (IsOmittedMarkup(element.LocalName)) continue;
                    if (element.SourceReference?.Position.Index is { } position)
                        while (cursor < ranges.Count && ranges[cursor].Start < position) cursor++;
                } else if (node is IText text && text.Data == ranges[cursor].Text) {
                    // DOM text can be coalesced or repaired. Bind when the decoded token and the reading-order node agree.
                    MarkupTextRun range = ranges[cursor++];
                    bool exact = range.Start >= 0 && range.End <= markup.Length && range.End >= range.Start &&
                        markup.AsSpan(range.Start, range.End - range.Start).SequenceEqual(text.Data.AsSpan());
                    matched.Add((range, text, exact));
                }
                for (INode? child = node.LastChild; child != null; child = child.PreviousSibling) pending.Push(child);
            }
            return matched;
        }

        private string AnchorText(IText node, int offset) {
            if (!_textAnchors.TryGetValue(node, out var anchors)) _textAnchors.Add(node, anchors = new());
            if (anchors.Count > 0 && anchors[anchors.Count - 1].Offset == offset) return anchors[anchors.Count - 1].Id;
            string id = "mobi-text-" + Number(++_textAnchorCount);
            anchors.Add((offset, id));
            return id;
        }

        internal IEnumerable<XElement> OmittedAnchors(IElement root, CancellationToken token, bool includeRoot = true) {
            var pending = new Stack<INode>();
            if (includeRoot) pending.Push(root);
            else PushChildren(root);
            while (pending.Count > 0) {
                token.ThrowIfCancellationRequested();
                INode node = pending.Pop();
                if (node is IElement element && Anchor(element) is { } id)
                    yield return new XElement("span", new XAttribute("id", id));
                if (node is IText text && _textAnchors.TryGetValue(text, out var anchors)) {
                    foreach (var anchor in anchors)
                        yield return new XElement("span", new XAttribute("id", anchor.Id));
                }
                PushChildren(node);
            }

            void PushChildren(INode node) {
                for (INode? child = node.LastChild; child != null; child = child.PreviousSibling)
                    pending.Push(child);
            }
        }

        internal IEnumerable<XNode> RenderText(IText text) {
            string value = text.Data;
            int start = 0;
            if (_textAnchors.TryGetValue(text, out var anchors)) {
                foreach (var anchor in anchors) {
                    if (anchor.Offset > start) yield return new XText(XmlText(value.Substring(start, anchor.Offset - start)));
                    yield return new XElement("span", new XAttribute("id", anchor.Id));
                    start = anchor.Offset;
                }
            }
            if (start < value.Length) yield return new XText(XmlText(value.Substring(start)));
        }
    }
}
