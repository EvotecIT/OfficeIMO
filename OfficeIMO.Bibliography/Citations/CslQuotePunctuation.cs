using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

/// <summary>Reorders quote punctuation while preserving each text segment's original formatting ancestry.</summary>
internal sealed class CslQuotePunctuation {
    private sealed class Run {
        internal int Start;
        internal string Text = string.Empty;
        internal XElement[] Path = Array.Empty<XElement>();
        internal bool Closing;
        internal bool WrittenEmpty;
        internal int End => Start + Text.Length;
    }

    private readonly List<Run> _runs = new List<Run>();
    private readonly List<int> _textRuns = new List<int>();
    private readonly HashSet<int> _emptyBlocks = new HashSet<int>();
    private readonly Dictionary<XElement, string> _openingTags = new Dictionary<XElement, string>();
    private readonly StringBuilder _plain = new StringBuilder();
    private readonly StringBuilder _html = new StringBuilder();
    private readonly int _maximumCharacters;
    private readonly CancellationToken _token;
    private XElement[] _path = Array.Empty<XElement>();

    private CslQuotePunctuation(int maximumCharacters, CancellationToken token) {
        _maximumCharacters = maximumCharacters;
        _token = token;
    }

    internal static string Apply(string html, int maximumCharacters, CancellationToken token) {
        XElement root = CslStyle.ReadXml("<root>" + html + "</root>", (int)Math.Min(int.MaxValue, (long)maximumCharacters + 13), 1024, token);
        var formatter = new CslQuotePunctuation(maximumCharacters, token);
        formatter.Read(root, Array.Empty<XElement>());
        return formatter.Write();
    }

    private void Read(XElement element, XElement[] path) {
        _token.ThrowIfCancellationRequested();
        foreach (XNode node in element.Nodes()) {
            _token.ThrowIfCancellationRequested();
            if (node is XText text) Add(text.Value, path, (string?)element.Attribute("data-csl-quote-end") == "true");
            else if (node is XElement child) {
                XElement[] next = path.Concat(new[] { child }).ToArray();
                if (child.Nodes().Any()) Read(child, next);
                else {
                    Add(string.Empty, next, false);
                    if (child.Name.LocalName == "div") _emptyBlocks.Add(_plain.Length);
                }
            }
        }
    }

    private void Add(string text, XElement[] path, bool closing) {
        if ((long)_plain.Length + text.Length > _maximumCharacters) throw Limit();
        if (text.Length > 0) _textRuns.Add(_runs.Count);
        _runs.Add(new Run { Start = _plain.Length, Text = text, Path = path, Closing = closing });
        _plain.Append(text);
    }

    private string Write() {
        string text = _plain.ToString();
        int cursor = 0;
        for (int index = 0; index < _textRuns.Count; index++) {
            _token.ThrowIfCancellationRequested();
            Run closing = _runs[_textRuns[index]];
            if (!closing.Closing || closing.Start < cursor) continue;
            int end = closing.End, last = index;
            while (last + 1 < _textRuns.Count) {
                Run next = _runs[_textRuns[last + 1]];
                if (!next.Closing || next.Start != end || _emptyBlocks.Contains(end) || !SameBlock(closing, next)) break;
                end = next.End; last++;
            }
            int punctuation = end;
            while (punctuation < text.Length && (text[punctuation] == ',' || text[punctuation] == '.')) {
                if ((punctuation & 1023) == 0) _token.ThrowIfCancellationRequested();
                if (_emptyBlocks.Contains(punctuation) || !SameBlock(closing, _runs[_textRuns[FindTextRun(punctuation)]])) break;
                punctuation++;
            }
            if (punctuation == end) { index = last; continue; }
            Range(cursor, closing.Start);
            char previous = closing.Start > 0 ? text[closing.Start - 1] : '\0';
            int start = end;
            for (int position = end; position < punctuation; position++) {
                if ((position & 1023) == 0) _token.ThrowIfCancellationRequested();
                char current = text[position];
                if (current == '.' && (previous == '.' || previous == '?' || previous == '!')) {
                    Range(start, position); start = position + 1;
                } else previous = current;
            }
            Range(start, punctuation);
            Range(closing.Start, end);
            cursor = punctuation;
            index = last;
        }
        Range(cursor, text.Length);
        foreach (Run run in _runs) if (run.Text.Length == 0 && !run.WrittenEmpty) Empty(run);
        Path(Array.Empty<XElement>());
        return _html.ToString();
    }

    // Only inline formatting may be crossed. Moving between display containers
    // would change structural layout, even when their visible text is adjacent.
    private static bool SameBlock(Run left, Run right) => left.Path.Where(element => element.Name.LocalName == "div")
        .SequenceEqual(right.Path.Where(element => element.Name.LocalName == "div"));

    private int FindTextRun(int position) {
        int low = 0, high = _textRuns.Count - 1;
        while (low < high) {
            int middle = low + (high - low + 1) / 2;
            if (_runs[_textRuns[middle]].Start <= position) low = middle;
            else high = middle - 1;
        }
        return low;
    }

    private void Range(int start, int end) {
        if (start >= end) return;
        int index = _textRuns[FindTextRun(start)];
        while (index > 0 && _runs[index - 1].Text.Length == 0 && _runs[index - 1].Start == start) index--;
        for (; index < _runs.Count && _runs[index].Start < end; index++) {
            _token.ThrowIfCancellationRequested();
            Run run = _runs[index];
            if (run.Text.Length == 0) { if (!run.WrittenEmpty) Empty(run); continue; }
            int offset = Math.Max(start, run.Start), stop = Math.Min(end, run.End);
            if (offset >= stop) continue;
            Path(run.Path);
            Append(CslText.Escape(run.Text.Substring(offset - run.Start, stop - offset)));
        }
    }

    private void Empty(Run run) {
        _token.ThrowIfCancellationRequested();
        Path(run.Path);
        // Closing immediately retains distinct empty display elements.
        Path(run.Path.Take(Math.Max(0, run.Path.Length - 1)).ToArray());
        run.WrittenEmpty = true;
    }

    private void Path(XElement[] next) {
        int common = 0;
        while (common < _path.Length && common < next.Length && ReferenceEquals(_path[common], next[common])) common++;
        for (int index = _path.Length - 1; index >= common; index--) Append("</" + _path[index].Name.LocalName + ">");
        for (int index = common; index < next.Length; index++) {
            _token.ThrowIfCancellationRequested();
            XElement element = next[index];
            if (!_openingTags.TryGetValue(element, out string? opening)) {
                opening = "<" + element.Name.LocalName + string.Concat(element.Attributes().Select(attribute => " " + attribute.Name.LocalName + "=\"" + CslText.Escape(attribute.Value) + "\"")) + ">";
                _openingTags.Add(element, opening);
            }
            Append(opening);
        }
        _path = next;
    }

    private void Append(string value) {
        if ((long)_html.Length + value.Length > _maximumCharacters) throw Limit();
        _html.Append(value);
    }

    private static InvalidDataException Limit() => new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
}
