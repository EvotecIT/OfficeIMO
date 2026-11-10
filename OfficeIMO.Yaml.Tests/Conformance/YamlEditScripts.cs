using System.Text;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>A kind of scripted edit.</summary>
public enum YamlEditKind {
    /// <summary>Give a scalar value a new string value.</summary>
    ChangeScalar,
    /// <summary>Give a scalar value an integer value.</summary>
    ChangeScalarTyped,
    /// <summary>Append a new key to a mapping.</summary>
    AddKey,
    /// <summary>Remove a mapping entry.</summary>
    RemoveEntry,
    /// <summary>Insert a sequence item at the start.</summary>
    InsertItem,
    /// <summary>Append a sequence item.</summary>
    AddItem,
    /// <summary>Remove a sequence item.</summary>
    RemoveItem,
    /// <summary>Set the trailing comment of a scalar value.</summary>
    EditComment,
    /// <summary>Set the comment lines above a key.</summary>
    EditCommentBefore,
    /// <summary>Rename a key.</summary>
    RenameKey,
    /// <summary>Replace a scalar value with a new block mapping, changing its layout.</summary>
    ReplaceWithMapping
}

/// <summary>One scripted edit on one input, with the outcome of every check.</summary>
public sealed class YamlEditResult {
    internal YamlEditResult(string subject, YamlEditKind kind, string target, string? failure, string? output) {
        Subject = subject;
        Kind = kind;
        Target = target;
        Failure = failure;
        Output = output;
    }

    /// <summary>Corpus path or suite case.</summary>
    public string Subject { get; }

    /// <summary>The edit.</summary>
    public YamlEditKind Kind { get; }

    /// <summary>The edited node's path.</summary>
    public string Target { get; }

    /// <summary>Why the edit failed a check, or null.</summary>
    public string? Failure { get; }

    /// <summary>The saved text.</summary>
    public string? Output { get; }

    /// <summary>Returns a one-line summary.</summary>
    public override string ToString() => Subject + " " + Kind + " " + Target + (Failure is null ? ": ok" : ": " + Failure);
}

/// <summary>
/// Applies scripted edits to YAML text and checks each saved result:
/// <list type="bullet">
/// <item><description>it parses again and loads to the same data, anchors and tags as the edited model;</description></item>
/// <item><description>every node path keeps the comments the edited model reports (no comment moves to another node);</description></item>
/// <item><description>only the expected region of the input changed, as one contiguous range (the minimal diff);</description></item>
/// <item><description>no <c>!!merge</c> tag appears that the input did not have;</description></item>
/// <item><description>the same edit on a fresh load gives the same bytes.</description></item>
/// </list>
/// </summary>
public static class YamlEditScripts {
    private static readonly YamlLoadOptions Lenient = new() { DuplicateKeys = YamlDuplicateKeyPolicy.LastWins };

    /// <summary>Runs up to <paramref name="perKind"/> edits of each kind, spread over the input's nodes.</summary>
    public static IReadOnlyList<YamlEditResult> Run(string subject, string yaml, int perKind) {
        var results = new List<YamlEditResult>();
        YamlDocument probe;
        try {
            probe = YamlDocument.Parse(yaml, Lenient);
        } catch (YamlException) {
            return results;
        }
        if (probe.Documents.Count == 0) return results;
        foreach (YamlEditKind kind in (YamlEditKind[])Enum.GetValues(typeof(YamlEditKind))) {
            List<string> targets = Targets(probe, kind);
            int step = Math.Max(1, targets.Count / Math.Max(1, perKind));
            for (int i = 0, n = 0; i < targets.Count && n < perKind; i += step, n++) results.Add(RunOne(subject, yaml, kind, targets[i]));
        }
        return results;
    }

    /// <summary>Applies one edit to a fresh load and checks the result.</summary>
    public static YamlEditResult RunOne(string subject, string yaml, YamlEditKind kind, string target) {
        YamlDocument document = YamlDocument.Parse(yaml, Lenient);
        YamlNode node = Locate(document, target)!;
        int lo, hi;
        bool pureInsertion = false, pureDeletion = false;
        try {
            Apply(document, kind, node, yaml, out lo, out hi, out pureInsertion, out pureDeletion);
        } catch (InvalidOperationException) {
            // The model refused the edit (an alias depends on the node, or it has no line of its own): nothing to check.
            return new YamlEditResult(subject, kind, target, null, null);
        }
        string output;
        try {
            output = document.ToString();
        } catch (Exception exception) {
            return Fail("save threw " + exception.GetType().Name + ": " + exception.Message, null);
        }
        if (YamlDocument.Parse(yaml, Lenient) is var again && Apply(again, kind, Locate(again, target)!, yaml) is var second && second != output) {
            return Fail("the same edit on a fresh load gave different bytes", output);
        }
        YamlDocument reloaded;
        try {
            reloaded = YamlDocument.Parse(output, Lenient);
        } catch (YamlException exception) {
            return Fail("output does not parse: " + exception.Message, output);
        }
        string? data = CompareData(document, reloaded);
        if (data is not null) return Fail(data, output);
        string? comments = CompareComments(document, reloaded);
        if (comments is not null) return Fail(comments, output);
        if (output.IndexOf("!!merge", StringComparison.Ordinal) >= 0 && yaml.IndexOf("!!merge", StringComparison.Ordinal) < 0) return Fail("a !!merge tag appeared", output);
        string? diff = CheckDiff(yaml, output, lo, hi, pureInsertion, pureDeletion);
        return diff is null ? new YamlEditResult(subject, kind, target, null, output) : Fail(diff, output);

        YamlEditResult Fail(string reason, string? text) => new YamlEditResult(subject, kind, target, reason, text);
    }

    /// <summary>Paths of nodes an edit kind applies to, in document order.</summary>
    private static List<string> Targets(YamlDocument document, YamlEditKind kind) {
        var targets = new List<string>();
        for (int d = 0; d < document.Documents.Count; d++) {
            foreach (YamlNode node in document.Documents[d].Root.Descendants()) {
                if (node.Parent is YamlMapping parent && parent.IndexOfNodeIsKey(node)) {
                    if (kind is YamlEditKind.EditCommentBefore or YamlEditKind.RenameKey && node is YamlScalar { Kind: YamlScalarKind.String }) targets.Add(Id(document, d, node));
                    continue;
                }
                bool applies = kind switch {
                    YamlEditKind.ChangeScalar or YamlEditKind.ChangeScalarTyped or YamlEditKind.EditComment or YamlEditKind.ReplaceWithMapping => node is YamlScalar && node.Parent is not null,
                    YamlEditKind.AddKey or YamlEditKind.RemoveEntry => node is YamlMapping { Count: > 0 },
                    YamlEditKind.InsertItem or YamlEditKind.AddItem or YamlEditKind.RemoveItem => node is YamlSequence { Count: > 0 },
                    _ => false
                };
                if (applies) targets.Add(Id(document, d, node));
            }
        }
        return targets;
    }

    /// <summary>An identifier made of child positions, so keys and values are both addressable whatever their text.</summary>
    private static string Id(YamlDocument document, int part, YamlNode node) {
        var steps = new List<string>();
        for (YamlNode current = node; current.Parent is { } parent; current = parent) {
            if (parent is YamlSequence sequence) {
                steps.Add(IndexOf(sequence, current).ToString(System.Globalization.CultureInfo.InvariantCulture));
            } else {
                var mapping = (YamlMapping)parent;
                for (int i = 0; i < mapping.Count; i++) {
                    if (ReferenceEquals(mapping[i].Key, current)) { steps.Add("k" + i); break; }
                    if (ReferenceEquals(mapping[i].Value, current)) { steps.Add("v" + i); break; }
                }
            }
        }
        steps.Add("d" + part);
        steps.Reverse();
        return string.Join("/", steps);
    }

    private static int IndexOf(YamlSequence sequence, YamlNode item) {
        for (int i = 0; i < sequence.Count; i++) {
            if (ReferenceEquals(sequence[i], item)) return i;
        }
        return -1;
    }

    private static YamlNode? Locate(YamlDocument document, string id) {
        string[] steps = id.Split('/');
        YamlNode node = document.Documents[int.Parse(steps[0].Substring(1), System.Globalization.CultureInfo.InvariantCulture)].Root;
        for (int i = 1; i < steps.Length; i++) {
            string step = steps[i];
            if (step[0] == 'k') node = ((YamlMapping)node)[int.Parse(step.Substring(1), System.Globalization.CultureInfo.InvariantCulture)].Key;
            else if (step[0] == 'v') node = ((YamlMapping)node)[int.Parse(step.Substring(1), System.Globalization.CultureInfo.InvariantCulture)].Value;
            else node = ((YamlSequence)node)[int.Parse(step, System.Globalization.CultureInfo.InvariantCulture)];
        }
        return node;
    }

    private static string Apply(YamlDocument document, YamlEditKind kind, YamlNode node, string yaml) {
        try {
            Apply(document, kind, node, yaml, out _, out _, out _, out _);
            return document.ToString();
        } catch (InvalidOperationException) {
            return string.Empty;
        }
    }

    /// <summary>Applies an edit and returns the source range [lo, hi] the change may touch.</summary>
    private static void Apply(YamlDocument document, YamlEditKind kind, YamlNode node, string yaml, out int lo, out int hi, out bool pureInsertion, out bool pureDeletion) {
        pureInsertion = pureDeletion = false;
        int start = (int)node.Start.Offset, end = (int)node.End.Offset;
        switch (kind) {
            case YamlEditKind.ChangeScalar: {
                lo = LineStart(yaml, node.Parent is YamlMapping ? Math.Min(start, (int)KeyOf(node).Start.Offset) : start);
                hi = Math.Max(LineEnd(yaml, end), ValueLineEnd(yaml, node));
                string text = node.AsString();
                ((YamlScalar)node).SetValue(text.Length == 0 ? "edited" : text + " edited");
                return;
            }
            case YamlEditKind.ChangeScalarTyped:
                lo = LineStart(yaml, node.Parent is YamlMapping ? Math.Min(start, (int)KeyOf(node).Start.Offset) : start);
                hi = Math.Max(LineEnd(yaml, end), ValueLineEnd(yaml, node));
                ((YamlScalar)node).SetValue(4242);
                return;
            case YamlEditKind.EditComment:
                lo = LineStart(yaml, node.Parent is YamlMapping ? Math.Min(start, (int)KeyOf(node).Start.Offset) : start);
                hi = Math.Max(LineEnd(yaml, end), ValueLineEnd(yaml, node));
                node.SetComment("edited comment");
                return;
            case YamlEditKind.ReplaceWithMapping: {
                lo = node.Parent is YamlMapping ? LineStart(yaml, KeyOf(node).Start.Offset) : LineStart(yaml, start);
                hi = Math.Max(LineEnd(yaml, end), ValueLineEnd(yaml, node));
                YamlNode? parent = node.Parent;
                var mapping = new YamlMapping { { "edited", "yes" }, { "count", 2 } };
                if (parent is YamlMapping map) map.Set(KeyText(map, node), mapping);
                else if (parent is YamlSequence sequence) sequence[IndexOf(sequence, node)] = mapping;
                return;
            }
            case YamlEditKind.AddKey: {
                var mapping = (YamlMapping)node;
                lo = start;
                hi = NextLine(yaml, end);
                // A single pair in a flow sequence (`[a: b]`) is rewritten with braces to hold a second entry.
                int content = SkipProperties(yaml, start);
                pureInsertion = mapping.Style != YamlStyle.Flow || (yaml[content] == '{' && content < mapping[0].Key.Start.Offset);
                mapping.Add(UniqueKey(mapping), "added value");
                return;
            }
            case YamlEditKind.RemoveEntry: {
                var mapping = (YamlMapping)node;
                int index = mapping.Count / 2;
                lo = PreviousContent(document, yaml, (int)mapping[index].Key.Start.Offset);
                hi = NextLine(yaml, (int)mapping[index].Value.End.Offset);
                pureDeletion = mapping.Count > 1;
                if (!pureDeletion) {
                    lo = Math.Min(lo, LineStart(yaml, KeyLine(node)));
                    hi = Math.Max(hi, NextLine(yaml, end));
                }
                mapping.RemoveAt(index);
                return;
            }
            case YamlEditKind.InsertItem:
                lo = PreviousContent(document, yaml, start);
                hi = NextLine(yaml, end);
                pureInsertion = true;
                ((YamlSequence)node).Insert(0, "inserted item");
                return;
            case YamlEditKind.AddItem:
                lo = start;
                hi = NextLine(yaml, end);
                pureInsertion = true;
                ((YamlSequence)node).Add("added item");
                return;
            case YamlEditKind.RemoveItem: {
                var sequence = (YamlSequence)node;
                int index = sequence.Count / 2;
                lo = PreviousContent(document, yaml, (int)sequence[index].Start.Offset);
                hi = NextLine(yaml, (int)sequence[index].End.Offset);
                pureDeletion = sequence.Count > 1;
                if (!pureDeletion) {
                    lo = Math.Min(lo, LineStart(yaml, KeyLine(node)));
                    hi = Math.Max(hi, NextLine(yaml, end));
                }
                sequence.RemoveAt(index);
                return;
            }
            case YamlEditKind.EditCommentBefore: {
                var mapping = (YamlMapping)node.Parent!;
                int index = KeyIndex(mapping, node);
                lo = PreviousContent(document, yaml, start);
                hi = LineEnd(yaml, start);
                node.SetCommentBefore("edited note");
                return;
            }
            case YamlEditKind.RenameKey: {
                var mapping = (YamlMapping)node.Parent!;
                lo = LineStart(yaml, start);
                hi = LineEnd(yaml, end);
                string text = node.AsString();
                mapping.RenameKey(text, UniqueKey(mapping, text + "_renamed"));
                return;
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(kind));
        }
    }

    private static int SkipProperties(string yaml, int i) {
        while (i < yaml.Length && (yaml[i] == '&' || yaml[i] == '!')) {
            while (i < yaml.Length && !char.IsWhiteSpace(yaml[i])) i++;
            while (i < yaml.Length && char.IsWhiteSpace(yaml[i])) i++;
        }
        return Math.Min(i, yaml.Length - 1);
    }

    private static YamlNode KeyOf(YamlNode value) {
        var mapping = (YamlMapping)value.Parent!;
        for (int i = 0; i < mapping.Count; i++) {
            if (ReferenceEquals(mapping[i].Value, value)) return mapping[i].Key;
        }
        throw new InvalidOperationException("Not a mapping value.");
    }

    private static string KeyText(YamlMapping mapping, YamlNode value) {
        YamlNode key = KeyOf(value);
        if (key is not YamlScalar || mapping.IndexOf(key.AsString()) < 0 || !ReferenceEquals(mapping[mapping.IndexOf(key.AsString())].Value, value)) {
            throw new InvalidOperationException("The key is not addressable by text.");
        }
        return key.AsString();
    }

    private static int KeyIndex(YamlMapping mapping, YamlNode key) {
        for (int i = 0; i < mapping.Count; i++) {
            if (ReferenceEquals(mapping[i].Key, key)) return i;
        }
        return -1;
    }

    /// <summary>The line holding the key of a collection that is a mapping value, or the collection's own start.</summary>
    private static int KeyLine(YamlNode collection) => collection.Parent is YamlMapping ? (int)KeyOf(collection).Start.Offset : (int)collection.Start.Offset;

    private static string UniqueKey(YamlMapping mapping, string name = "zz_added") {
        string key = name;
        for (int i = 2; mapping.ContainsKey(key); i++) key = name + i;
        return key;
    }

    /// <summary>The start of the line where the last content before a position ends: everything after it may belong to the next entry.</summary>
    private static int PreviousContent(YamlDocument document, string yaml, int position) {
        long best = 0;
        foreach (YamlDocumentPart part in document.Documents) {
            foreach (YamlNode node in part.Root.Descendants()) {
                long end = node.End.Offset;
                if (end <= position && end > best && node.Start.Offset < position) best = end;
            }
        }
        return LineStart(yaml, best);
    }

    private static int CollectionStart(YamlNode collection) => collection.Parent is YamlMapping ? (int)KeyOf(collection).Start.Offset : (int)collection.Start.Offset;

    /// <summary>For a mapping value, the end of the line holding its key's colon, so a moved comment stays in range.</summary>
    private static int ValueLineEnd(string yaml, YamlNode node) => node.Parent is YamlMapping ? LineEnd(yaml, (int)KeyOf(node).End.Offset) : 0;

    private static int LineStart(string text, long position) {
        int i = (int)Math.Min(position, text.Length);
        while (i > 0 && text[i - 1] != '\n' && text[i - 1] != '\r') i--;
        return i;
    }

    private static int NextLine(string text, int position) {
        int i = LineEnd(text, position);
        if (i < text.Length && text[i] == '\r') i++;
        if (i < text.Length && text[i] == '\n') i++;
        return i;
    }

    private static int LineEnd(string text, int position) {
        int i = Math.Min(position, text.Length);
        while (i < text.Length && text[i] != '\n' && text[i] != '\r') i++;
        return i;
    }

    /// <summary>
    /// Checks the change is one contiguous range of the input inside [lo, hi]. Equal text next to an insertion or a
    /// removal lets the range slide, so any placement within the edited region counts; a replacement is measured both
    /// prefix-first and suffix-first.
    /// </summary>
    private static string? CheckDiff(string input, string output, int lo, int hi, bool pureInsertion, bool pureDeletion) {
        int max = Math.Min(input.Length, output.Length);
        int prefix = 0;
        while (prefix < max && input[prefix] == output[prefix]) prefix++;
        int suffix = 0;
        while (suffix < max && input[input.Length - 1 - suffix] == output[output.Length - 1 - suffix]) suffix++;
        int delta = output.Length - input.Length;
        if (pureInsertion || pureDeletion) {
            if (pureInsertion && delta <= 0) return "an insertion did not add text";
            if (pureDeletion && delta >= 0) return "a removal did not remove text";
            int size = Math.Abs(delta);
            // Valid placements p: the input is unchanged before p and after p (+ size for a removal).
            int first = input.Length - (pureDeletion ? size : 0) - suffix;
            int last = prefix;
            if (first < 0) first = 0;
            if (first > last) return (pureInsertion ? "an insertion" : "a removal") + " also changed other text near " + last;
            int from = Math.Max(first, lo), to = Math.Min(last, hi - (pureDeletion ? size : 0));
            if (from > to) return (pureInsertion ? "inserted" : "removed") + " text at " + first + ".." + last + " lies outside the edited region " + lo + ".." + hi;
            return null;
        }
        int cappedSuffix = Math.Min(suffix, max - prefix);
        int start = prefix, end = input.Length - cappedSuffix;
        int cappedPrefix = Math.Min(prefix, max - suffix);
        int start2 = cappedPrefix, end2 = input.Length - suffix;
        bool fits = (start >= lo && end <= hi) || (start2 >= lo && end2 <= hi);
        if (fits) return null;
        return "the input changed at " + start + ".." + end + ", outside the edited region " + lo + ".." + hi;
    }

    private static string? CompareData(YamlDocument expected, YamlDocument actual) {
        if (expected.Documents.Count != actual.Documents.Count) return "the output has " + actual.Documents.Count + " documents instead of " + expected.Documents.Count;
        for (int i = 0; i < expected.Documents.Count; i++) {
            string? difference = Compare(expected.Documents[i].Root, actual.Documents[i].Root, "$[" + i + "]");
            if (difference is not null) return difference;
        }
        return null;
    }

    private static string? Compare(YamlNode expected, YamlNode actual, string path) {
        if (expected is YamlAlias ea) return actual is YamlAlias aa && aa.Name == ea.Name ? null : path + ": expected alias *" + ea.Name;
        if (actual is YamlAlias) return path + ": unexpected alias";
        if (expected.Anchor != actual.Anchor) return path + ": anchor " + (actual.Anchor ?? "none") + " instead of " + (expected.Anchor ?? "none");
        if (expected.Tag != actual.Tag) return path + ": tag " + (actual.Tag ?? "none") + " instead of " + (expected.Tag ?? "none");
        switch (expected) {
            case YamlScalar es:
                if (actual is not YamlScalar acs) return path + ": expected a scalar";
                if (es.Kind != acs.Kind) return path + ": kind " + acs.Kind + " instead of " + es.Kind;
                if (es.AsString() != acs.AsString() && !(es.Kind == YamlScalarKind.Null && acs.Kind == YamlScalarKind.Null)) {
                    bool typedEqual = es.Kind switch {
                        YamlScalarKind.Boolean => es.GetValue<bool>() == acs.GetValue<bool>(),
                        YamlScalarKind.Integer or YamlScalarKind.Float => es.TryGetValue(out double x) && acs.TryGetValue(out double y) && x.Equals(y),
                        _ => false
                    };
                    if (!typedEqual) return path + ": value '" + acs.AsString() + "' instead of '" + es.AsString() + "'";
                }
                return null;
            case YamlMapping em:
                if (actual is not YamlMapping am || am.Count != em.Count) return path + ": expected a mapping of " + em.Count;
                for (int i = 0; i < em.Count; i++) {
                    string? key = Compare(em[i].Key, am[i].Key, path + ".key" + i);
                    if (key is not null) return key;
                    string? value = Compare(em[i].Value, am[i].Value, path + "." + i);
                    if (value is not null) return value;
                }
                return null;
            case YamlSequence eq:
                if (actual is not YamlSequence aq || aq.Count != eq.Count) return path + ": expected a sequence of " + eq.Count;
                for (int i = 0; i < eq.Count; i++) {
                    string? item = Compare(eq[i], aq[i], path + "[" + i + "]");
                    if (item is not null) return item;
                }
                return null;
        }
        return null;
    }

    /// <summary>Checks that every node reports the same comments and blank line in the edited model and in the reloaded output.</summary>
    private static string? CompareComments(YamlDocument expected, YamlDocument actual) {
        for (int d = 0; d < expected.Documents.Count; d++) {
            if (expected.Documents[d].CommentBefore != actual.Documents[d].CommentBefore) return "document header '" + actual.Documents[d].CommentBefore + "' instead of '" + expected.Documents[d].CommentBefore + "'";
            if (expected.Documents[d].CommentAfter != actual.Documents[d].CommentAfter) return "document footer '" + actual.Documents[d].CommentAfter + "' instead of '" + expected.Documents[d].CommentAfter + "'";
            using var left = expected.Documents[d].Root.Descendants().GetEnumerator();
            using var right = actual.Documents[d].Root.Descendants().GetEnumerator();
            while (left.MoveNext() && right.MoveNext()) {
                YamlNode e = left.Current, a = right.Current;
                if (InFlow(e)) continue;
                if (e.CommentBefore != a.CommentBefore) return Id(expected, d, e) + ": comment above '" + a.CommentBefore + "' instead of '" + e.CommentBefore + "'";
                if (e.Comment != a.Comment) return Id(expected, d, e) + ": comment '" + a.Comment + "' instead of '" + e.Comment + "'";
                if (e.SpaceBefore != a.SpaceBefore) return Id(expected, d, e) + ": blank line above " + a.SpaceBefore + " instead of " + e.SpaceBefore;
            }
        }
        return null;
    }

    private static bool InFlow(YamlNode node) {
        for (YamlNode? current = node.Parent; current is not null; current = current.Parent) {
            if (current.Style == YamlStyle.Flow) return true;
        }
        return node.Style == YamlStyle.Flow;
    }

    /// <summary>Formats failures for an assertion message, with the output of the first few.</summary>
    public static string Describe(IEnumerable<YamlEditResult> failures, int limit = 3) {
        var text = new StringBuilder();
        int shown = 0;
        foreach (YamlEditResult failure in failures) {
            text.AppendLine(failure.ToString());
            if (shown++ < limit && failure.Output is not null) text.AppendLine("---- output ----").AppendLine(failure.Output).AppendLine("----");
        }
        return text.ToString();
    }
}

internal static class YamlEditScriptExtensions {
    public static bool IndexOfNodeIsKey(this YamlMapping mapping, YamlNode node) {
        for (int i = 0; i < mapping.Count; i++) {
            if (ReferenceEquals(mapping[i].Key, node)) return true;
        }
        return false;
    }
}
