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
public static partial class OfficeBidiTextResolver {
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
        IReadOnlyList<string> elements = OfficeTextElements.Split(text);
        if (!elements.Any(static element => IsParagraphBoundary(element))) {
            return ResolveParagraphRuns(text!, baseDirection, cancellationToken);
        }

        var runs = new List<OfficeBidiTextRun>();
        var paragraph = new StringBuilder();
        foreach (string element in elements) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!IsParagraphBoundary(element)) {
                paragraph.Append(element);
                continue;
            }

            AppendParagraphRuns(runs, paragraph, baseDirection, cancellationToken);
            runs.Add(new OfficeBidiTextRun(element, OfficeTextDirection.LeftToRight, 0, runs.Count));
        }
        AppendParagraphRuns(runs, paragraph, baseDirection, cancellationToken);
        return runs.AsReadOnly();
    }

    // Levels come from the Unicode Bidirectional Algorithm (UAX #9). Runs group text elements of
    // one resolved level; bidi formatting characters do not paint and are omitted from run text.
    private static IReadOnlyList<OfficeBidiTextRun> ResolveParagraphRuns(
        string text,
        OfficeTextDirection baseDirection,
        CancellationToken cancellationToken) {
        int[] scalars = OfficeUnicodeBidi.ToScalars(text, out _);
        int? paragraphLevel = baseDirection switch {
            OfficeTextDirection.RightToLeft => 1,
            OfficeTextDirection.LeftToRight => 0,
            _ => null
        };
        OfficeUnicodeBidi.Paragraph paragraph = OfficeUnicodeBidi.Resolve(scalars, paragraphLevel, cancellationToken);
        byte[] levels = paragraph.GetLineLevels(0, scalars.Length);

        var runs = new List<OfficeBidiTextRun>();
        var value = new StringBuilder();
        int lastLevel = -1;
        int scalarIndex = 0;
        IReadOnlyList<string> elements = OfficeTextElements.Split(text);
        for (int index = 0; index < elements.Count; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            string element = elements[index];
            int elementScalars = CountScalars(element);
            int level = scalarIndex < levels.Length ? levels[scalarIndex] : paragraph.ParagraphLevel;
            scalarIndex += elementScalars;
            if (element.Length > 0 && OfficeTextElements.ContainsBidiControl(element) && IsOnlyBidiControls(element)) continue;
            if (value.Length > 0 && level != lastLevel) {
                Flush(runs, value, (lastLevel & 1) == 1 ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight, lastLevel);
            }
            value.Append(element);
            lastLevel = level;
        }
        if (value.Length > 0) {
            Flush(runs, value, (lastLevel & 1) == 1 ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight, lastLevel);
        }
        return runs.AsReadOnly();
    }

    private static int CountScalars(string element) {
        int count = 0;
        for (int index = 0; index < element.Length; index++) {
            if (char.IsHighSurrogate(element[index]) && index + 1 < element.Length && char.IsLowSurrogate(element[index + 1])) index++;
            count++;
        }
        return count;
    }

    private static bool IsOnlyBidiControls(string element) {
        foreach (char unit in element) {
            if (!OfficeTextElements.ContainsBidiControl(unit.ToString())) return false;
        }
        return true;
    }

    private static void AppendParagraphRuns(
        List<OfficeBidiTextRun> destination,
        StringBuilder paragraph,
        OfficeTextDirection baseDirection,
        CancellationToken cancellationToken) {
        if (paragraph.Length == 0) return;
        foreach (OfficeBidiTextRun run in ResolveParagraphRuns(paragraph.ToString(), baseDirection, cancellationToken)) {
            destination.Add(new OfficeBidiTextRun(
                run.Text,
                run.Direction,
                run.EmbeddingLevel,
                destination.Count));
        }
        paragraph.Clear();
    }

    private static bool IsParagraphBoundary(string element) =>
        element is "\r" or "\n" or "\r\n" or "\u2028" or "\u2029";

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

    private static void Flush(List<OfficeBidiTextRun> runs, StringBuilder value, OfficeTextDirection direction, int level) {
        if (value.Length == 0) return;
        runs.Add(new OfficeBidiTextRun(value.ToString(), direction, level, runs.Count));
        value.Clear();
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
