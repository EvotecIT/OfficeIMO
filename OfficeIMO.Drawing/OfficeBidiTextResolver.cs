using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>One logical text run with an explicit resolved paint direction.</summary>
public sealed class OfficeBidiTextRun {
    internal OfficeBidiTextRun(string text, OfficeTextDirection direction) {
        Text = text;
        Direction = direction;
    }

    /// <summary>Text with Unicode bidi formatting controls removed.</summary>
    public string Text { get; }

    /// <summary>Resolved direction for this run.</summary>
    public OfficeTextDirection Direction { get; }
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
        states.Push(new DirectionalState(resolvedBase, false, false));
        OfficeTextDirection lastDirection = resolvedBase;
        IReadOnlyList<string> elements = OfficeTextElements.Split(text);
        for (int index = 0; index < elements.Count; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            string element = elements[index];
            if (element.Length == 1 && OfficeTextElements.ContainsBidiControl(element)) {
                Flush(runs, value, lastDirection);
                TryApplyControl(elements, ref index, element[0], states, ref lastDirection);
                continue;
            }

            DirectionalState state = states.Peek();
            OfficeTextDirection direction = state.Override
                ? state.Direction
                : ResolveElementDirection(element, state.Direction, lastDirection);
            if (value.Length > 0 && direction != lastDirection) Flush(runs, value, lastDirection);
            value.Append(element);
            lastDirection = direction;
        }
        Flush(runs, value, lastDirection);
        return runs.AsReadOnly();
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
        OfficeTextDirection resolvedBase = ResolveBaseDirection(text!, baseDirection);
        IEnumerable<OfficeBidiTextRun> paintRuns = resolvedBase == OfficeTextDirection.RightToLeft
            ? runs.Reverse()
            : runs;
        var visual = new StringBuilder(text!.Length);
        foreach (OfficeBidiTextRun run in paintRuns) {
            cancellationToken.ThrowIfCancellationRequested();
            if (run.Direction == OfficeTextDirection.RightToLeft) {
                IReadOnlyList<string> elements = OfficeTextElements.Split(run.Text);
                for (int index = elements.Count - 1; index >= 0; index--) visual.Append(elements[index]);
            } else {
                visual.Append(run.Text);
            }
        }
        return visual.ToString();
    }

    internal static IReadOnlyList<T> ToVisualOrder<T>(
        string directionalText,
        IReadOnlyList<T> visibleElements,
        OfficeTextDirection baseDirection,
        CancellationToken cancellationToken) {
        IReadOnlyList<OfficeBidiTextRun> runs = ResolveRuns(directionalText, baseDirection, cancellationToken);
        var groups = new List<IReadOnlyList<T>>(runs.Count);
        int elementIndex = 0;
        foreach (OfficeBidiTextRun run in runs) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = OfficeTextElements.Split(run.Text).Count;
            if (count > visibleElements.Count - elementIndex) return Array.Empty<T>();
            var group = new List<T>(count);
            for (int index = 0; index < count; index++) group.Add(visibleElements[elementIndex++]);
            if (run.Direction == OfficeTextDirection.RightToLeft) group.Reverse();
            groups.Add(group);
        }
        if (elementIndex != visibleElements.Count) return Array.Empty<T>();
        if (ResolveBaseDirection(directionalText, baseDirection) == OfficeTextDirection.RightToLeft) groups.Reverse();
        return groups.SelectMany(static group => group).ToArray();
    }

    private static bool TryApplyControl(
        IReadOnlyList<string> elements,
        ref int index,
        char control,
        Stack<DirectionalState> states,
        ref OfficeTextDirection lastDirection) {
        switch (control) {
            case '\u061C':
            case '\u200F':
                lastDirection = OfficeTextDirection.RightToLeft;
                return true;
            case '\u200E':
                lastDirection = OfficeTextDirection.LeftToRight;
                return true;
            case '\u202A': Push(states, OfficeTextDirection.LeftToRight, false, false); return true;
            case '\u202B': Push(states, OfficeTextDirection.RightToLeft, false, false); return true;
            case '\u202D': Push(states, OfficeTextDirection.LeftToRight, true, false); return true;
            case '\u202E': Push(states, OfficeTextDirection.RightToLeft, true, false); return true;
            case '\u202C':
                if (states.Count > 1 && !states.Peek().Isolate) states.Pop();
                lastDirection = states.Peek().Direction;
                return true;
            case '\u2066': Push(states, OfficeTextDirection.LeftToRight, false, true); return true;
            case '\u2067': Push(states, OfficeTextDirection.RightToLeft, false, true); return true;
            case '\u2068':
                Push(states, ResolveFirstStrongIsolateDirection(elements, index + 1, states.Peek().Direction), false, true);
                return true;
            case '\u2069':
                while (states.Count > 1) {
                    DirectionalState state = states.Pop();
                    if (state.Isolate) break;
                }
                lastDirection = states.Peek().Direction;
                return true;
            default:
                return false;
        }
    }

    private static void Push(Stack<DirectionalState> states, OfficeTextDirection direction, bool @override, bool isolate) {
        if (states.Count < MaximumEmbeddingDepth) states.Push(new DirectionalState(direction, @override, isolate));
    }

    private static OfficeTextDirection ResolveFirstStrongIsolateDirection(
        IReadOnlyList<string> elements,
        int start,
        OfficeTextDirection fallback) {
        int nested = 0;
        for (int index = start; index < elements.Count; index++) {
            if (elements[index].Length == 1) {
                char control = elements[index][0];
                if (control is '\u2066' or '\u2067' or '\u2068') { nested++; continue; }
                if (control == '\u2069') {
                    if (nested == 0) break;
                    nested--;
                    continue;
                }
                if (OfficeTextElements.ContainsBidiControl(elements[index])) continue;
            }
            if (nested > 0) continue;
            OfficeTextDirection direction = OfficeTextElements.ResolveBaseDirection(elements[index]);
            if (direction != OfficeTextDirection.Auto) return direction;
        }
        return fallback;
    }

    private static OfficeTextDirection ResolveElementDirection(
        string element,
        OfficeTextDirection embeddingDirection,
        OfficeTextDirection lastDirection) {
        OfficeTextDirection direction = OfficeTextElements.ResolveBaseDirection(element);
        if (direction != OfficeTextDirection.Auto) return direction;
        for (int index = 0; index < element.Length; index++) {
            if (char.IsDigit(element[index])) return OfficeTextDirection.LeftToRight;
        }
        return lastDirection == OfficeTextDirection.Auto ? embeddingDirection : lastDirection;
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

    private static void Flush(List<OfficeBidiTextRun> runs, StringBuilder value, OfficeTextDirection direction) {
        if (value.Length == 0) return;
        runs.Add(new OfficeBidiTextRun(value.ToString(), direction));
        value.Clear();
    }

    private readonly struct DirectionalState {
        internal DirectionalState(OfficeTextDirection direction, bool @override, bool isolate) {
            Direction = direction;
            Override = @override;
            Isolate = isolate;
        }
        internal OfficeTextDirection Direction { get; }
        internal bool Override { get; }
        internal bool Isolate { get; }
    }
}
