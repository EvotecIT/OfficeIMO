using System.Globalization;
using System.Text;

namespace OfficeIMO.Diff;

/// <summary>Kind of an <see cref="OfficeChange"/>.</summary>
public enum OfficeChangeKind {
    /// <summary>Content present only in the new document.</summary>
    Inserted,
    /// <summary>Content present only in the old document.</summary>
    Deleted,
    /// <summary>Old content replaced by different new content at the same place.</summary>
    Replaced,
    /// <summary>The same content deleted in one place and inserted in another.</summary>
    Moved
}

/// <summary>A rectangle on a page that shows where a change sits in its source document.</summary>
/// <remarks>Coordinates are in the source format's own page units; for PDF this is default user space (points, origin bottom-left).</remarks>
public sealed class OfficeChangeRegion {
    /// <summary>Creates a region.</summary>
    public OfficeChangeRegion(int pageNumber, double left, double bottom, double right, double top) {
        if (pageNumber <= 0) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        PageNumber = pageNumber;
        Left = Math.Min(left, right);
        Right = Math.Max(left, right);
        Bottom = Math.Min(bottom, top);
        Top = Math.Max(bottom, top);
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Left edge.</summary>
    public double Left { get; }

    /// <summary>Bottom edge.</summary>
    public double Bottom { get; }

    /// <summary>Right edge.</summary>
    public double Right { get; }

    /// <summary>Top edge.</summary>
    public double Top { get; }
}

/// <summary>Where one side of a change sits in its source: a token range, plus a page, a location label and page regions when the source has them.</summary>
public sealed class OfficeChangeAnchor {
    /// <summary>Creates an anchor.</summary>
    /// <param name="start">Zero-based index of the first token of the change in the source token sequence.</param>
    /// <param name="length">Number of tokens; zero marks an insertion point.</param>
    /// <param name="pageNumber">One-based page of the first token, when the source is paged.</param>
    /// <param name="location">Free-form location, for example a paragraph or cell reference.</param>
    /// <param name="regions">Page rectangles covered by the change.</param>
    public OfficeChangeAnchor(int start, int length, int? pageNumber = null, string? location = null, IEnumerable<OfficeChangeRegion>? regions = null) {
        if (start < 0) throw new ArgumentOutOfRangeException(nameof(start));
        if (length < 0) throw new ArgumentOutOfRangeException(nameof(length));
        if (pageNumber.HasValue && pageNumber.Value <= 0) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        Start = start;
        Length = length;
        PageNumber = pageNumber;
        Location = location;
        Regions = regions?.ToArray() ?? Array.Empty<OfficeChangeRegion>();
    }

    /// <summary>Zero-based index of the first token.</summary>
    public int Start { get; }

    /// <summary>Number of tokens; zero marks an insertion point.</summary>
    public int Length { get; }

    /// <summary>One-based page, when the source is paged.</summary>
    public int? PageNumber { get; }

    /// <summary>Free-form location label.</summary>
    public string? Location { get; }

    /// <summary>Page rectangles covered by the change.</summary>
    public IReadOnlyList<OfficeChangeRegion> Regions { get; }
}

/// <summary>One change between two documents.</summary>
public sealed class OfficeChange {
    /// <summary>Creates a change.</summary>
    /// <param name="id">One-based identifier, unique within its <see cref="OfficeChangeSet"/>.</param>
    /// <param name="kind">Change kind.</param>
    /// <param name="oldText">Removed text; empty for an insertion.</param>
    /// <param name="newText">Added text; empty for a deletion.</param>
    /// <param name="oldAnchor">Position in the old document; for an insertion, the insertion point.</param>
    /// <param name="newAnchor">Position in the new document; for a deletion, the point where the text used to be.</param>
    public OfficeChange(int id, OfficeChangeKind kind, string? oldText, string? newText, OfficeChangeAnchor? oldAnchor, OfficeChangeAnchor? newAnchor) {
        if (id <= 0) throw new ArgumentOutOfRangeException(nameof(id));
        Id = id;
        Kind = kind;
        OldText = oldText ?? string.Empty;
        NewText = newText ?? string.Empty;
        Old = oldAnchor;
        New = newAnchor;
    }

    /// <summary>One-based identifier.</summary>
    public int Id { get; }

    /// <summary>Change kind.</summary>
    public OfficeChangeKind Kind { get; }

    /// <summary>Removed text; empty for an insertion. For a move, the moved text.</summary>
    public string OldText { get; }

    /// <summary>Added text; empty for a deletion. For a move, the moved text.</summary>
    public string NewText { get; }

    /// <summary>Position in the old document.</summary>
    public OfficeChangeAnchor? Old { get; }

    /// <summary>Position in the new document.</summary>
    public OfficeChangeAnchor? New { get; }
}

/// <summary>Settings for <see cref="OfficeChangeSet.FromTextDiff"/>.</summary>
public sealed class OfficeChangeSetOptions {
    /// <summary>Pairs a deletion with an insertion of the same text elsewhere and reports them as one move. The default is <c>true</c>.</summary>
    public bool DetectMoves { get; set; } = true;

    /// <summary>Shortest text, in characters after collapsing whitespace, that is reported as a move. The default is 12.</summary>
    public int MinimumMoveLength { get; set; } = 12;

    /// <summary>Keeps changes that only add or remove whitespace. The default is <c>false</c>.</summary>
    public bool IncludeWhitespaceOnlyChanges { get; set; }

    /// <summary>Name of the old document, written to reports.</summary>
    public string? OldSource { get; set; }

    /// <summary>Name of the new document, written to reports.</summary>
    public string? NewSource { get; set; }

    /// <summary>Maps an old token index to its page number, when the source is paged.</summary>
    public Func<int, int?>? OldPageNumber { get; set; }

    /// <summary>Maps a new token index to its page number, when the source is paged.</summary>
    public Func<int, int?>? NewPageNumber { get; set; }
}

/// <summary>
/// The changes between two documents in a format-neutral form, ready to serialise to JSON or render as a report.
/// </summary>
public sealed class OfficeChangeSet {
    /// <summary>Creates a change set.</summary>
    /// <param name="changes">Changes in document order.</param>
    /// <param name="oldSource">Name of the old document.</param>
    /// <param name="newSource">Name of the new document.</param>
    /// <param name="unit">Token unit the changes were computed in, for example <c>word</c> or <c>line</c>.</param>
    /// <param name="isComplete">False when a work or time limit stopped the comparison early (see <see cref="IsComplete"/>).</param>
    public OfficeChangeSet(IEnumerable<OfficeChange> changes, string? oldSource = null, string? newSource = null, string unit = "word", bool isComplete = true) {
        if (changes is null) throw new ArgumentNullException(nameof(changes));
        Changes = changes.ToArray();
        OldSource = oldSource;
        NewSource = newSource;
        Unit = unit ?? "word";
        IsComplete = isComplete;
    }

    /// <summary>Changes in document order.</summary>
    public IReadOnlyList<OfficeChange> Changes { get; }

    /// <summary>Name of the old document.</summary>
    public string? OldSource { get; }

    /// <summary>Name of the new document.</summary>
    public string? NewSource { get; }

    /// <summary>Token unit the changes were computed in.</summary>
    public string Unit { get; }

    /// <summary>
    /// True when the alignment ran to the end. Either way every difference is reported, because nothing is skipped when a limit is reached;
    /// when false, some unchanged text may appear as deleted and inserted again.
    /// </summary>
    public bool IsComplete { get; }

    /// <summary>True when there are no changes.</summary>
    public bool IsEmpty => Changes.Count == 0;

    /// <summary>Number of changes of the given kind.</summary>
    public int Count(OfficeChangeKind kind) => Changes.Count(change => change.Kind == kind);

    /// <summary>Builds a change set from a word, line or character diff.</summary>
    public static OfficeChangeSet FromTextDiff(OfficeDiffResult<OfficeTextToken> diff, OfficeChangeSetOptions? options = null) {
        if (diff is null) throw new ArgumentNullException(nameof(diff));
        options ??= new OfficeChangeSetOptions();
        var drafts = new List<Draft>();
        foreach (OfficeDiffHunk region in diff.GetChangedRegions()) {
            int oldStart = region.OldStart, oldEnd = region.OldStart + region.OldLength;
            int newStart = region.NewStart, newEnd = region.NewStart + region.NewLength;
            if (!options.IncludeWhitespaceOnlyChanges) {
                TrimWhitespace(diff.Old, ref oldStart, ref oldEnd);
                TrimWhitespace(diff.New, ref newStart, ref newEnd);
                if (oldStart == oldEnd && newStart == newEnd) continue;
            }

            drafts.Add(new Draft(oldStart, oldEnd - oldStart, Join(diff.Old, oldStart, oldEnd), newStart, newEnd - newStart, Join(diff.New, newStart, newEnd)));
        }

        if (options.DetectMoves) PairMoves(drafts, options.MinimumMoveLength);

        var changes = new List<OfficeChange>(drafts.Count);
        foreach (Draft draft in drafts.Where(static draft => !draft.Absorbed).OrderBy(static draft => draft.NewStart).ThenBy(static draft => draft.OldStart)) {
            OfficeChangeKind kind = draft.MovedTo != null ? OfficeChangeKind.Moved
                : draft.OldLength == 0 ? OfficeChangeKind.Inserted
                : draft.NewLength == 0 ? OfficeChangeKind.Deleted
                : OfficeChangeKind.Replaced;
            Draft target = draft.MovedTo ?? draft;
            var oldAnchor = new OfficeChangeAnchor(draft.OldStart, draft.OldLength, options.OldPageNumber?.Invoke(draft.OldStart));
            var newAnchor = new OfficeChangeAnchor(target.NewStart, target.NewLength, options.NewPageNumber?.Invoke(target.NewStart));
            changes.Add(new OfficeChange(changes.Count + 1, kind, draft.OldText, target.NewText, oldAnchor, newAnchor));
        }

        string unit = diff.Old.Concat(diff.New).Select(static token => token.Kind).FirstOrDefault() switch {
            OfficeTextTokenKind.Line => "line",
            OfficeTextTokenKind.Character => "character",
            _ => "word"
        };
        return new OfficeChangeSet(changes, options.OldSource, options.NewSource, unit, !diff.LimitReached);
    }

    /// <summary>Serialises the change set to JSON with camel-case property names and string enum values.</summary>
    public string ToJson(bool indented = true) => OfficeChangeSetJson.Write(this, indented);

    private static void TrimWhitespace(IReadOnlyList<OfficeTextToken> tokens, ref int start, ref int end) {
        while (start < end && string.IsNullOrWhiteSpace(tokens[start].Text)) start++;
        while (end > start && string.IsNullOrWhiteSpace(tokens[end - 1].Text)) end--;
    }

    private static string Join(IReadOnlyList<OfficeTextToken> tokens, int start, int end) {
        var builder = new StringBuilder();
        for (int index = start; index < end; index++) builder.Append(tokens[index].Text);
        return builder.ToString();
    }

    internal static string NormalizeForMove(string text) =>
        string.Join(" ", (text ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static void PairMoves(List<Draft> drafts, int minimumLength) {
        var insertions = new Dictionary<string, Queue<Draft>>(StringComparer.Ordinal);
        foreach (Draft draft in drafts) {
            if (draft.OldLength != 0 || draft.NewLength == 0) continue;
            string key = NormalizeForMove(draft.NewText);
            if (key.Length < minimumLength) continue;
            if (!insertions.TryGetValue(key, out Queue<Draft>? queue)) {
                queue = new Queue<Draft>();
                insertions.Add(key, queue);
            }

            queue.Enqueue(draft);
        }

        foreach (Draft draft in drafts) {
            if (draft.NewLength != 0 || draft.OldLength == 0) continue;
            if (!insertions.TryGetValue(NormalizeForMove(draft.OldText), out Queue<Draft>? queue) || queue.Count == 0) continue;
            Draft insertion = queue.Dequeue();
            draft.MovedTo = insertion;
            insertion.Absorbed = true;
        }
    }

    private sealed class Draft {
        internal Draft(int oldStart, int oldLength, string oldText, int newStart, int newLength, string newText) {
            OldStart = oldStart;
            OldLength = oldLength;
            OldText = oldText;
            NewStart = newStart;
            NewLength = newLength;
            NewText = newText;
        }

        internal int OldStart { get; }
        internal int OldLength { get; }
        internal string OldText { get; }
        internal int NewStart { get; }
        internal int NewLength { get; }
        internal string NewText { get; }
        internal Draft? MovedTo { get; set; }
        internal bool Absorbed { get; set; }
    }
}

internal static class OfficeChangeSetJson {
    internal static string Write(OfficeChangeSet set, bool indented) {
        var writer = new JsonWriter(indented);
        writer.BeginObject();
        writer.Property("oldSource", set.OldSource);
        writer.Property("newSource", set.NewSource);
        writer.Property("unit", set.Unit);
        writer.Property("isComplete", set.IsComplete);
        writer.Name("summary");
        writer.BeginObject();
        writer.Property("total", set.Changes.Count);
        writer.Property("inserted", set.Count(OfficeChangeKind.Inserted));
        writer.Property("deleted", set.Count(OfficeChangeKind.Deleted));
        writer.Property("replaced", set.Count(OfficeChangeKind.Replaced));
        writer.Property("moved", set.Count(OfficeChangeKind.Moved));
        writer.EndObject();
        writer.Name("changes");
        writer.BeginArray();
        foreach (OfficeChange change in set.Changes) {
            writer.BeginObject();
            writer.Property("id", change.Id);
            writer.Property("kind", ToCamel(change.Kind.ToString()));
            writer.Property("oldText", change.OldText);
            writer.Property("newText", change.NewText);
            WriteAnchor(writer, "old", change.Old);
            WriteAnchor(writer, "new", change.New);
            writer.EndObject();
        }

        writer.EndArray();
        writer.EndObject();
        return writer.ToString();
    }

    private static void WriteAnchor(JsonWriter writer, string name, OfficeChangeAnchor? anchor) {
        writer.Name(name);
        if (anchor is null) {
            writer.Null();
            return;
        }

        writer.BeginObject();
        writer.Property("start", anchor.Start);
        writer.Property("length", anchor.Length);
        if (anchor.PageNumber.HasValue) writer.Property("pageNumber", anchor.PageNumber.Value);
        if (anchor.Location != null) writer.Property("location", anchor.Location);
        if (anchor.Regions.Count > 0) {
            writer.Name("regions");
            writer.BeginArray();
            foreach (OfficeChangeRegion region in anchor.Regions) {
                writer.BeginObject(inline: true);
                writer.Property("pageNumber", region.PageNumber);
                writer.Property("left", region.Left);
                writer.Property("bottom", region.Bottom);
                writer.Property("right", region.Right);
                writer.Property("top", region.Top);
                writer.EndObject();
            }

            writer.EndArray();
        }

        writer.EndObject();
    }

    private static string ToCamel(string value) => value.Length == 0 ? value : char.ToLowerInvariant(value[0]) + value.Substring(1);

    private sealed class JsonWriter {
        private readonly StringBuilder _output = new StringBuilder();
        private readonly bool _indented;
        private readonly Stack<(bool First, bool Inline)> _scopes = new Stack<(bool, bool)>();
        private bool _afterName;

        internal JsonWriter(bool indented) {
            _indented = indented;
        }

        internal void BeginObject(bool inline = false) => Begin('{', inline);

        internal void EndObject() => End('}');

        internal void BeginArray() => Begin('[', false);

        internal void EndArray() => End(']');

        internal void Name(string name) {
            Separator();
            WriteString(name);
            _output.Append(_indented ? ": " : ":");
            _afterName = true;
        }

        internal void Property(string name, string? value) {
            Name(name);
            if (value is null) Null();
            else Value(() => WriteString(value));
        }

        internal void Property(string name, int value) {
            Name(name);
            Value(() => _output.Append(value.ToString(CultureInfo.InvariantCulture)));
        }

        internal void Property(string name, double value) {
            Name(name);
            Value(() => _output.Append(double.IsNaN(value) || double.IsInfinity(value) ? "null" : Math.Round(value, 3).ToString("0.###", CultureInfo.InvariantCulture)));
        }

        internal void Property(string name, bool value) {
            Name(name);
            Value(() => _output.Append(value ? "true" : "false"));
        }

        internal void Null() => Value(() => _output.Append("null"));

        public override string ToString() => _output.ToString();

        private void Begin(char token, bool inline) {
            Separator();
            _afterName = false;
            _output.Append(token);
            _scopes.Push((true, inline || (_scopes.Count > 0 && _scopes.Peek().Inline)));
        }

        private void End(char token) {
            (bool first, bool inline) = _scopes.Pop();
            if (!first) NewLine(inline);
            _output.Append(token);
        }

        private void Value(Action write) {
            Separator();
            _afterName = false;
            write();
        }

        private void Separator() {
            if (_afterName || _scopes.Count == 0) return;
            (bool first, bool inline) = _scopes.Pop();
            if (!first) _output.Append(',');
            if (inline) {
                if (!first && _indented) _output.Append(' ');
            } else {
                _scopes.Push((false, inline));
                NewLine(false);
                return;
            }

            _scopes.Push((false, inline));
        }

        private void NewLine(bool inline) {
            if (!_indented || inline) return;
            _output.Append('\n').Append(' ', _scopes.Count * 2);
        }

        private void WriteString(string value) {
            _output.Append('"');
            foreach (char character in value) {
                switch (character) {
                    case '"': _output.Append("\\\""); break;
                    case '\\': _output.Append("\\\\"); break;
                    case '\n': _output.Append("\\n"); break;
                    case '\r': _output.Append("\\r"); break;
                    case '\t': _output.Append("\\t"); break;
                    case '\b': _output.Append("\\b"); break;
                    case '\f': _output.Append("\\f"); break;
                    default:
                        if (character < 0x20 || character == '\u2028' || character == '\u2029') {
                            _output.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                        } else {
                            _output.Append(character);
                        }

                        break;
                }
            }

            _output.Append('"');
        }
    }
}
