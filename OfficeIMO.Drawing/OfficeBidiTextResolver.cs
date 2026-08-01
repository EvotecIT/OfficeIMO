using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>One logical text run with an explicit resolved paint direction.</summary>
public sealed class OfficeBidiTextRun {
    internal OfficeBidiTextRun(string text, OfficeTextDirection direction, int embeddingLevel, int logicalOrder) {
        Text = text;
        Direction = direction;
        EmbeddingLevel = embeddingLevel;
        LogicalOrder = logicalOrder;
    }

    /// <summary>Text with Unicode bidi formatting controls removed.</summary>
    public string Text { get; }

    /// <summary>Resolved direction for this run.</summary>
    public OfficeTextDirection Direction { get; }

    /// <summary>Zero-based run order in the logical source text.</summary>
    public int LogicalOrder { get; }

    internal int EmbeddingLevel { get; }
}

/// <summary>
/// Resolves bounded directional runs, including Unicode embeddings, overrides, and isolates, without external shaping dependencies.
/// </summary>
public static class OfficeBidiTextResolver {
    private const int MaximumEmbeddingDepth = 125;

    /// <summary>Resolves logical directional runs and removes non-painting bidi controls.</summary>
    public static IReadOnlyList<OfficeBidiTextRun> ResolveRuns(
        string? text,
        OfficeTextDirection baseDirection = OfficeTextDirection.Auto) {
        return ResolveRuns(text, baseDirection, CancellationToken.None);
    }

    /// <summary>Resolves logical directional runs and removes non-painting bidi controls.</summary>
    public static IReadOnlyList<OfficeBidiTextRun> ResolveRuns(
        string? text,
        OfficeTextDirection baseDirection,
        CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(text)) return Array.Empty<OfficeBidiTextRun>();
        OfficeTextDirection resolvedBase = ResolveBaseDirection(text!, baseDirection);

        var runs = new List<OfficeBidiTextRun>();
        var value = new StringBuilder();
        var states = new Stack<DirectionalState>();
        states.Push(new DirectionalState(
            resolvedBase,
            false,
            false,
            resolvedBase == OfficeTextDirection.RightToLeft ? 1 : 0,
            resolvedBase));
        var overflow = new BidiOverflowState();
        OfficeTextDirection lastDirection = resolvedBase;
        int lastLevel = states.Peek().Level;
        IReadOnlyList<string> elements = OfficeTextElements.Split(text);
        OfficeTextDirection?[] firstStrongIsolateDirections = ResolveFirstStrongIsolateDirections(elements, cancellationToken);
        OfficeTextDirection?[] followingStrongDirections = ResolveFollowingStrongDirections(elements, cancellationToken);
        for (int index = 0; index < elements.Count; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            string element = elements[index];
            if (element.Length == 1 && OfficeTextElements.ContainsBidiControl(element)) {
                Flush(runs, value, lastDirection, lastLevel);
                TryApplyControl(
                    element[0],
                    firstStrongIsolateDirections[index],
                    states,
                    ref overflow,
                    ref lastDirection);
                lastLevel = ResolveLevel(states.Peek().Level, lastDirection);
                continue;
            }

            DirectionalState state = states.Peek();
            OfficeTextDirection direction = state.Override
                ? state.Direction
                : ResolveElementDirection(element, state.Direction, lastDirection, followingStrongDirections[index]);
            int level = ResolveLevel(state.Level, direction);
            if (value.Length > 0 && (direction != lastDirection || level != lastLevel)) {
                Flush(runs, value, lastDirection, lastLevel);
            }
            value.Append(element);
            lastDirection = direction;
            lastLevel = level;
        }
        Flush(runs, value, lastDirection, lastLevel);
        return runs.AsReadOnly();
    }

    /// <summary>
    /// Resolves directional runs in visual placement order while retaining logical text order inside each run.
    /// </summary>
    public static IReadOnlyList<OfficeBidiTextRun> ResolveVisualRuns(
        string? text,
        OfficeTextDirection baseDirection = OfficeTextDirection.Auto) {
        return ResolveVisualRuns(text, baseDirection, CancellationToken.None);
    }

    /// <summary>
    /// Resolves directional runs in visual placement order while retaining logical text order inside each run.
    /// </summary>
    public static IReadOnlyList<OfficeBidiTextRun> ResolveVisualRuns(
        string? text,
        OfficeTextDirection baseDirection,
        CancellationToken cancellationToken) {
        IReadOnlyList<OfficeBidiTextRun> runs = ResolveRuns(text, baseDirection, cancellationToken);
        var visualRuns = runs
            .Select(static run => new VisualElement<OfficeBidiTextRun>(run, run.EmbeddingLevel))
            .ToList();
        ReorderByEmbeddingLevel(visualRuns, cancellationToken);
        return visualRuns.Select(static run => run.Value).ToArray();
    }

    /// <summary>Returns paint-order text after applying bounded embeddings, overrides, and isolates.</summary>
    public static string ToVisualOrder(
        string? text,
        OfficeTextDirection baseDirection = OfficeTextDirection.Auto) {
        return ToVisualOrder(text, baseDirection, CancellationToken.None);
    }

    /// <summary>Returns paint-order text after applying bounded embeddings, overrides, and isolates.</summary>
    public static string ToVisualOrder(
        string? text,
        OfficeTextDirection baseDirection,
        CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(text)) return text ?? string.Empty;
        IReadOnlyList<OfficeBidiTextRun> runs = ResolveRuns(text, baseDirection, cancellationToken);
        var visualElements = new List<VisualElement<string>>();
        foreach (OfficeBidiTextRun run in runs) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (string element in OfficeTextElements.Split(run.Text)) {
                visualElements.Add(new VisualElement<string>(element, run.EmbeddingLevel));
            }
        }
        ReorderByEmbeddingLevel(visualElements, cancellationToken);
        var visual = new StringBuilder(text!.Length);
        foreach (VisualElement<string> element in visualElements) {
            visual.Append((element.Level & 1) == 1 ? MirrorText(element.Value) : element.Value);
        }
        return visual.ToString();
    }

    internal static IReadOnlyList<T> ToVisualOrder<T>(
        string directionalText,
        IReadOnlyList<T> visibleElements,
        OfficeTextDirection baseDirection,
        CancellationToken cancellationToken,
        Func<T, T>? mirrorOddLevel = null) {
        IReadOnlyList<OfficeBidiTextRun> runs = ResolveRuns(directionalText, baseDirection, cancellationToken);
        var visualElements = new List<VisualElement<T>>(visibleElements.Count);
        int elementIndex = 0;
        foreach (OfficeBidiTextRun run in runs) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = OfficeTextElements.Split(run.Text).Count;
            if (count > visibleElements.Count - elementIndex) return Array.Empty<T>();
            for (int index = 0; index < count; index++) {
                T value = visibleElements[elementIndex++];
                if ((run.EmbeddingLevel & 1) == 1 && mirrorOddLevel != null) value = mirrorOddLevel(value);
                visualElements.Add(new VisualElement<T>(value, run.EmbeddingLevel));
            }
        }
        if (elementIndex != visibleElements.Count) return Array.Empty<T>();
        ReorderByEmbeddingLevel(visualElements, cancellationToken);
        return visualElements.Select(static element => element.Value).ToArray();
    }

    internal static IReadOnlyList<IReadOnlyList<T>> ToVisualLineOrder<T>(
        string directionalText,
        IReadOnlyList<IReadOnlyList<T>> visibleLines,
        OfficeTextDirection baseDirection,
        CancellationToken cancellationToken,
        Func<T, T>? mirrorOddLevel = null) {
        IReadOnlyList<OfficeBidiTextRun> runs = ResolveRuns(directionalText, baseDirection, cancellationToken);
        int visibleElementCount = visibleLines.Sum(static line => line.Count);
        var resolvedElements = new List<VisualElement<T>>(visibleElementCount);
        int lineIndex = 0;
        int elementIndex = 0;
        foreach (OfficeBidiTextRun run in runs) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = OfficeTextElements.Split(run.Text).Count;
            for (int index = 0; index < count; index++) {
                while (lineIndex < visibleLines.Count && elementIndex >= visibleLines[lineIndex].Count) {
                    lineIndex++;
                    elementIndex = 0;
                }
                if (lineIndex >= visibleLines.Count) return Array.Empty<IReadOnlyList<T>>();
                T value = visibleLines[lineIndex][elementIndex++];
                if ((run.EmbeddingLevel & 1) == 1 && mirrorOddLevel != null) value = mirrorOddLevel(value);
                resolvedElements.Add(new VisualElement<T>(value, run.EmbeddingLevel));
            }
        }
        if (resolvedElements.Count != visibleElementCount) return Array.Empty<IReadOnlyList<T>>();

        var result = new List<IReadOnlyList<T>>(visibleLines.Count);
        int offset = 0;
        foreach (IReadOnlyList<T> line in visibleLines) {
            cancellationToken.ThrowIfCancellationRequested();
            List<VisualElement<T>> visualLine = resolvedElements.GetRange(offset, line.Count);
            ReorderByEmbeddingLevel(visualLine, cancellationToken);
            result.Add(visualLine.Select(static element => element.Value).ToArray());
            offset += line.Count;
        }
        return result.AsReadOnly();
    }

    internal static string MirrorText(string value) {
        if (value.Length == 0) return value;
        var mirrored = new StringBuilder(value.Length);
        foreach (char codeUnit in value) {
            mirrored.Append(OfficeBidiMirroring.Mirror(codeUnit));
        }
        return mirrored.ToString();
    }

    private static bool TryApplyControl(
        char control,
        OfficeTextDirection? firstStrongIsolateDirection,
        Stack<DirectionalState> states,
        ref BidiOverflowState overflow,
        ref OfficeTextDirection lastDirection) {
        switch (control) {
            case '\u061C':
            case '\u200F':
                lastDirection = OfficeTextDirection.RightToLeft;
                return true;
            case '\u200E':
                lastDirection = OfficeTextDirection.LeftToRight;
                return true;
            case '\u202A': Push(states, OfficeTextDirection.LeftToRight, false, false, lastDirection, ref overflow); return true;
            case '\u202B': Push(states, OfficeTextDirection.RightToLeft, false, false, lastDirection, ref overflow); return true;
            case '\u202D': Push(states, OfficeTextDirection.LeftToRight, true, false, lastDirection, ref overflow); return true;
            case '\u202E': Push(states, OfficeTextDirection.RightToLeft, true, false, lastDirection, ref overflow); return true;
            case '\u202C':
                if (overflow.OverflowEmbeddingCount > 0) {
                    overflow.OverflowEmbeddingCount--;
                } else if (overflow.OverflowIsolateCount == 0 && states.Count > 1 && !states.Peek().Isolate) {
                    lastDirection = states.Pop().OuterStrongDirection;
                }
                return true;
            case '\u2066': Push(states, OfficeTextDirection.LeftToRight, false, true, lastDirection, ref overflow); return true;
            case '\u2067': Push(states, OfficeTextDirection.RightToLeft, false, true, lastDirection, ref overflow); return true;
            case '\u2068':
                Push(states, firstStrongIsolateDirection ?? states.Peek().Direction, false, true, lastDirection, ref overflow);
                return true;
            case '\u2069':
                if (overflow.OverflowIsolateCount > 0) {
                    overflow.OverflowIsolateCount--;
                } else if (overflow.ValidIsolateCount > 0) {
                    overflow.OverflowEmbeddingCount = 0;
                    OfficeTextDirection outerStrongDirection = lastDirection;
                    while (states.Count > 1) {
                        DirectionalState state = states.Pop();
                        if (state.Isolate) {
                            outerStrongDirection = state.OuterStrongDirection;
                            break;
                        }
                    }
                    overflow.ValidIsolateCount--;
                    lastDirection = outerStrongDirection;
                }
                return true;
            default:
                return false;
        }
    }

    private static void Push(
        Stack<DirectionalState> states,
        OfficeTextDirection direction,
        bool @override,
        bool isolate,
        OfficeTextDirection outerStrongDirection,
        ref BidiOverflowState overflow) {
        int level = NextEmbeddingLevel(states.Peek().Level, direction);
        bool overflowed = level > MaximumEmbeddingDepth
            || overflow.OverflowIsolateCount > 0
            || overflow.OverflowEmbeddingCount > 0;
        if (overflowed) {
            if (isolate) overflow.OverflowIsolateCount++;
            else if (overflow.OverflowIsolateCount == 0) overflow.OverflowEmbeddingCount++;
            return;
        }

        states.Push(new DirectionalState(direction, @override, isolate, level, outerStrongDirection));
        if (isolate) overflow.ValidIsolateCount++;
    }

    private static int NextEmbeddingLevel(int currentLevel, OfficeTextDirection direction) {
        int next = currentLevel + 1;
        bool requiresOdd = direction == OfficeTextDirection.RightToLeft;
        if (((next & 1) == 1) != requiresOdd) next++;
        return next;
    }

    private static int ResolveLevel(int embeddingLevel, OfficeTextDirection direction) {
        bool requiresOdd = direction == OfficeTextDirection.RightToLeft;
        return ((embeddingLevel & 1) == 1) == requiresOdd ? embeddingLevel : embeddingLevel + 1;
    }

    private static void ReorderByEmbeddingLevel<T>(List<VisualElement<T>> elements, CancellationToken cancellationToken) {
        if (elements.Count == 0) return;
        int maximumLevel = elements.Max(static element => element.Level);
        int minimumOddLevel = elements
            .Where(static element => (element.Level & 1) == 1)
            .Select(static element => element.Level)
            .DefaultIfEmpty(int.MaxValue)
            .Min();
        if (minimumOddLevel == int.MaxValue) return;

        for (int level = maximumLevel; level >= minimumOddLevel; level--) {
            cancellationToken.ThrowIfCancellationRequested();
            int start = 0;
            while (start < elements.Count) {
                while (start < elements.Count && elements[start].Level < level) start++;
                int end = start;
                while (end < elements.Count && elements[end].Level >= level) end++;
                if (end > start) elements.Reverse(start, end - start);
                start = end;
            }
        }
    }

    private static OfficeTextDirection?[] ResolveFirstStrongIsolateDirections(
        IReadOnlyList<string> elements,
        CancellationToken cancellationToken) {
        var directions = new OfficeTextDirection?[elements.Count];
        var isolates = new Stack<FirstStrongIsolateFrame>();
        int overflowDepth = 0;
        for (int index = 0; index < elements.Count; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            string element = elements[index];
            if (element.Length == 1) {
                char control = element[0];
                if (control is '\u2066' or '\u2067' or '\u2068') {
                    if (overflowDepth > 0 || isolates.Count >= MaximumEmbeddingDepth) {
                        overflowDepth++;
                    } else {
                        isolates.Push(new FirstStrongIsolateFrame(control == '\u2068' ? index : -1));
                    }
                    continue;
                }
                if (control == '\u2069') {
                    if (overflowDepth > 0) {
                        overflowDepth--;
                    } else if (isolates.Count > 0) {
                        StoreFirstStrongDirection(isolates.Pop(), directions);
                    }
                    continue;
                }
                if (control is '\u061C' or '\u200E' or '\u200F') {
                    if (overflowDepth == 0 && isolates.Count > 0 && isolates.Peek().SourceIndex >= 0 && !isolates.Peek().Direction.HasValue) {
                        isolates.Peek().Direction = control == '\u200E'
                            ? OfficeTextDirection.LeftToRight
                            : OfficeTextDirection.RightToLeft;
                    }
                    continue;
                }
                if (OfficeTextElements.ContainsBidiControl(element)) continue;
            }

            if (overflowDepth > 0 || isolates.Count == 0 || isolates.Peek().SourceIndex < 0 || isolates.Peek().Direction.HasValue) {
                continue;
            }

            OfficeTextDirection direction = OfficeTextElements.ResolveBaseDirection(element);
            if (direction != OfficeTextDirection.Auto) isolates.Peek().Direction = direction;
        }

        while (isolates.Count > 0) StoreFirstStrongDirection(isolates.Pop(), directions);
        return directions;
    }

    private static void StoreFirstStrongDirection(
        FirstStrongIsolateFrame isolate,
        OfficeTextDirection?[] directions) {
        if (isolate.SourceIndex >= 0) directions[isolate.SourceIndex] = isolate.Direction;
    }

    private static OfficeTextDirection ResolveElementDirection(
        string element,
        OfficeTextDirection embeddingDirection,
        OfficeTextDirection lastDirection,
        OfficeTextDirection? followingStrongDirection) {
        OfficeTextDirection direction = OfficeTextElements.ResolveBaseDirection(element);
        if (direction != OfficeTextDirection.Auto) return direction;
        for (int index = 0; index < element.Length; index++) {
            if (char.IsDigit(element[index])) return OfficeTextDirection.LeftToRight;
        }
        return followingStrongDirection.HasValue && followingStrongDirection.Value == lastDirection
            ? lastDirection
            : embeddingDirection;
    }

    private static OfficeTextDirection?[] ResolveFollowingStrongDirections(
        IReadOnlyList<string> elements,
        CancellationToken cancellationToken) {
        var directions = new OfficeTextDirection?[elements.Count];
        OfficeTextDirection? following = null;
        for (int index = elements.Count - 1; index >= 0; index--) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            string element = elements[index];
            if (element is "\u200E") {
                following = OfficeTextDirection.LeftToRight;
                continue;
            }
            if (element is "\u061C" or "\u200F") {
                following = OfficeTextDirection.RightToLeft;
                continue;
            }
            if (element.Length == 1 && OfficeTextElements.ContainsBidiControl(element)) {
                following = null;
                continue;
            }

            directions[index] = following;
            OfficeTextDirection direction = OfficeTextElements.ResolveBaseDirection(element);
            if (direction != OfficeTextDirection.Auto) {
                following = direction;
                continue;
            }
            for (int characterIndex = 0; characterIndex < element.Length; characterIndex++) {
                if (char.IsDigit(element[characterIndex])) {
                    following = OfficeTextDirection.LeftToRight;
                    break;
                }
            }
        }
        return directions;
    }

    private static OfficeTextDirection ResolveBaseDirectionWithoutFormattingControls(string text) {
        var value = new StringBuilder(text.Length);
        foreach (string element in OfficeTextElements.Enumerate(text)) {
            if (!OfficeTextElements.ContainsBidiControl(element) || element is "\u061C" or "\u200E" or "\u200F") value.Append(element);
        }
        return OfficeTextElements.ResolveBaseDirection(value.ToString());
    }

    private static OfficeTextDirection ResolveBaseDirection(string text, OfficeTextDirection requestedDirection) {
        OfficeTextDirection direction = requestedDirection == OfficeTextDirection.Auto
            ? ResolveBaseDirectionWithoutFormattingControls(text)
            : requestedDirection;
        return direction == OfficeTextDirection.Auto ? OfficeTextDirection.LeftToRight : direction;
    }

    private static void Flush(List<OfficeBidiTextRun> runs, StringBuilder value, OfficeTextDirection direction, int level) {
        if (value.Length == 0) return;
        runs.Add(new OfficeBidiTextRun(value.ToString(), direction, level, runs.Count));
        value.Clear();
    }

    private readonly struct DirectionalState {
        internal DirectionalState(
            OfficeTextDirection direction,
            bool @override,
            bool isolate,
            int level,
            OfficeTextDirection outerStrongDirection) {
            Direction = direction;
            Override = @override;
            Isolate = isolate;
            Level = level;
            OuterStrongDirection = outerStrongDirection;
        }
        internal OfficeTextDirection Direction { get; }
        internal bool Override { get; }
        internal bool Isolate { get; }
        internal int Level { get; }
        internal OfficeTextDirection OuterStrongDirection { get; }
    }

    private struct BidiOverflowState {
        internal int OverflowEmbeddingCount;
        internal int OverflowIsolateCount;
        internal int ValidIsolateCount;
    }

    private sealed class FirstStrongIsolateFrame {
        internal FirstStrongIsolateFrame(int sourceIndex) {
            SourceIndex = sourceIndex;
        }
        internal int SourceIndex { get; }
        internal OfficeTextDirection? Direction { get; set; }
    }

    private readonly struct VisualElement<T> {
        internal VisualElement(T value, int level) {
            Value = value;
            Level = level;
        }

        internal T Value { get; }
        internal int Level { get; }
    }
}
