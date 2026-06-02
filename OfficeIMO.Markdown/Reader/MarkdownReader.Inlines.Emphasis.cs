namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private enum FrameKind {
        Root,
        Italic,
        Bold,
        Strike,
        Highlight,
    }

    private sealed class InlineFrame {
        public InlineFrame(FrameKind kind, char marker, int openLen, InlineSequence seq, int openIndex) {
            Kind = kind;
            Marker = marker;
            OpenLen = openLen;
            Seq = seq;
            OpenIndex = openIndex;
        }

        public FrameKind Kind { get; }
        public char Marker { get; }
        public int OpenLen { get; }
        public InlineSequence Seq { get; }
        public int OpenIndex { get; }
    }

    private static bool TryCloseFrame(Stack<InlineFrame> stack, char marker, int remaining, out int consumed) {
        consumed = 0;
        if (stack == null || stack.Count <= 1) return false;
        var top = stack.Peek();
        if (top.Marker != marker) return false;

        // Close the innermost matching frame only; this avoids crossing.
        if (top.Kind == FrameKind.Italic && remaining >= 1) {
            stack.Pop();
            var node = new ItalicSequenceInline(top.Seq);
            stack.Peek().Seq.AddRaw(node);
            consumed = 1;
            return true;
        }
        if (top.Kind == FrameKind.Bold && remaining >= 2) {
            stack.Pop();
            var node = new BoldSequenceInline(top.Seq);
            stack.Peek().Seq.AddRaw(node);
            consumed = 2;
            return true;
        }
        if (top.Kind == FrameKind.Strike && remaining >= top.OpenLen) {
            stack.Pop();
            var node = new StrikethroughSequenceInline(top.Seq);
            stack.Peek().Seq.AddRaw(node);
            consumed = top.OpenLen;
            return true;
        }
        if (top.Kind == FrameKind.Highlight && remaining >= 2) {
            stack.Pop();
            var node = new HighlightSequenceInline(top.Seq);
            stack.Peek().Seq.AddRaw(node);
            consumed = 2;
            return true;
        }
        return false;
    }

    private static bool ShouldTreatSingleMarkerAsLiteralInsideBold(string text, int start, char marker, int runLen, Stack<InlineFrame> stack) {
        if (runLen != 1) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Bold || top.Marker != marker || top.OpenLen != 2) return false;

        int nextDoubleClose = FindNextClosingDelimiterRunIndex(text, start + 1, marker, requiredRunLength: 2);
        if (nextDoubleClose >= 0) {
            int trailingSingleClose = FindNextClosingDelimiterRunIndex(text, nextDoubleClose + 2, marker, requiredRunLength: 1);
            if (trailingSingleClose >= 0) return false;
        }

        int nextRun = FindNextDelimiterRunLength(text, start + 1, marker);
        return nextRun == 2;
    }

    private static bool ShouldTreatDelimiterRunAsLiteral(string text, int start, char marker, int runLen, Stack<InlineFrame> stack, bool splitDoubleRunIntoDualItalic, out int literalRunLength) {
        literalRunLength = 0;
        if (runLen != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;

        var frames = stack.ToArray();
        if (frames.Length >= 2) {
            var parent = frames[1];
            // Keep the leading triple-delimiter path available for rebalancing into <em><strong>... later.
            if (parent.Kind == FrameKind.Bold && parent.Marker == marker && parent.OpenLen == 2 && parent.Seq.Nodes.Count == 0) return false;

            if (parent.Kind == FrameKind.Bold && parent.Marker == marker && parent.OpenLen == 2 && parent.Seq.Nodes.Count > 0) {
                int trailingSingleClose = FindNextClosingDelimiterRunIndex(text, start + 2, marker, requiredRunLength: 1);
                if (trailingSingleClose >= 0) return false;
            }
        }

        if (splitDoubleRunIntoDualItalic) return false;

        int nextRun = FindNextDelimiterRunLength(text, start + 2, marker);
        if (nextRun != 1) return false;

        literalRunLength = 2;
        return true;
    }

    private static int FindNextDelimiterRunLength(string text, int start, char marker) {
        if (string.IsNullOrEmpty(text)) return 0;
        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int run = 1;
            while (i + run < text.Length && text[i + run] == marker) run++;
            return run;
        }
        return 0;
    }

    private static bool TryRebalanceLeadingBoldInsideItalic(Stack<InlineFrame> stack, char marker, int remaining, out int consumed) {
        consumed = 0;
        if (remaining < 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (stack == null || stack.Count < 3) return false;

        var frames = stack.ToArray();
        var top = frames[0];
        var parent = frames[1];
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;
        if (parent.Kind != FrameKind.Bold || parent.Marker != marker || parent.OpenLen != 2) return false;
        if (parent.Seq.Nodes.Count != 0) return false;

        stack.Pop();
        stack.Pop();

        var italic = new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, parent.OpenIndex);
        italic.Seq.AddRaw(new BoldSequenceInline(top.Seq));
        stack.Push(italic);
        consumed = 2;
        return true;
    }

    private static bool TryRebalanceParentBoldWithInnerItalicIntoDualItalic(string text, int start, Stack<InlineFrame> stack, char marker, int remaining, out int consumed) {
        consumed = 0;
        if (remaining != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count < 3) return false;

        var frames = stack.ToArray();
        var top = frames[0];
        var parent = frames[1];
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;
        if (parent.Kind != FrameKind.Bold || parent.Marker != marker || parent.OpenLen != 2) return false;
        if (parent.Seq.Nodes.Count == 0) return false;

        int trailingSingleClose = FindNextClosingDelimiterRunIndex(text, start + 2, marker, requiredRunLength: 1);
        if (trailingSingleClose < 0) return false;

        stack.Pop();
        stack.Pop();

        var middle = new InlineSequence { AutoSpacing = false };
        foreach (var node in parent.Seq.Nodes) {
            middle.AddRaw(node);
        }

        middle.AddRaw(new ItalicSequenceInline(top.Seq));

        var outer = new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, parent.OpenIndex);
        outer.Seq.AddRaw(new ItalicSequenceInline(middle));
        stack.Push(outer);
        consumed = 2;
        return true;
    }

    private static bool ShouldPreferInnerBold(Stack<InlineFrame> stack, char marker, int remaining, bool canOpen, bool canClose) {
        if (!canOpen || !canClose || remaining != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        return top.Marker == marker && top.Kind == FrameKind.Italic;
    }

    private static bool ShouldSplitDoubleUnderscoreToLiteralAndItalic(string text, int start, int runLen, bool canOpen, bool canClose) {
        if (!canOpen || canClose) return false;
        if (runLen != 2) return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (text[start] != '_') return false;

        return !HasFutureClosingDelimiterRun(text, start + 2, '_', minimumRunLength: 2) &&
               HasFutureClosingDelimiterRun(text, start + 2, '_', minimumRunLength: 1);
    }

    private static int GetLiteralPrefixLengthForOddCloser(string text, int start, char marker, int runLen, bool canOpen, bool canClose) {
        if (!canOpen || canClose) return 0;
        if (runLen < 2 || (runLen % 2) != 0) return 0;
        if (marker != '*' && marker != '_') return 0;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return 0;

        for (int candidate = runLen - 1; candidate >= 1; candidate -= 2) {
            if (FindNextClosingDelimiterRunIndex(text, start + runLen, marker, requiredRunLength: candidate) < 0) continue;

            bool hasSameOrLongerEvenCloser = false;
            for (int even = runLen; even >= candidate + 1; even -= 2) {
                if (FindNextClosingDelimiterRunIndex(text, start + runLen, marker, requiredRunLength: even) >= 0) {
                    hasSameOrLongerEvenCloser = true;
                    break;
                }
            }

            if (!hasSameOrLongerEvenCloser) {
                return runLen - candidate;
            }
        }

        return 0;
    }

    private static bool HasFutureClosingDelimiterRun(string text, int start, char marker, int minimumRunLength) {
        if (string.IsNullOrEmpty(text)) return false;
        if (minimumRunLength <= 0) minimumRunLength = 1;

        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, out _, out bool canClose);
            if (canClose && runLen >= minimumRunLength) return true;

            i += runLen - 1;
        }

        return false;
    }

    private static bool ShouldTreatMixedSingleMarkerAsLiteral(string text, int start, char marker, int runLen, bool canOpen, bool canClose, Stack<InlineFrame> stack) {
        if (!canOpen || canClose) return false;
        if (runLen != 1) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Italic || top.OpenLen != 1) return false;
        if (top.Marker == marker) return false;

        int outerClose = FindNextClosingDelimiterIndex(text, start + 1, top.Marker, minimumRunLength: 1);
        if (outerClose < 0) return false;

        int innerClose = FindNextClosingDelimiterIndex(text, start + 1, marker, minimumRunLength: 1);
        return innerClose < 0 || outerClose < innerClose;
    }

    private static bool ShouldSplitDoubleRunIntoDualItalic(string text, int start, char marker, int runLen, Stack<InlineFrame> stack) {
        if (runLen != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;

        int singleClose = FindNextClosingDelimiterRunIndex(text, start + 2, marker, requiredRunLength: 1);
        if (singleClose < 0) return false;

        int doubleClose = FindNextClosingDelimiterRunIndex(text, singleClose + 1, marker, requiredRunLength: 2);
        if (doubleClose < 0) return false;

        int afterSingle = singleClose + 1;
        return afterSingle < text.Length && char.IsWhiteSpace(text[afterSingle]);
    }

    private static int FindNextClosingDelimiterIndex(string text, int start, char marker, int minimumRunLength) {
        if (string.IsNullOrEmpty(text)) return -1;
        if (minimumRunLength <= 0) minimumRunLength = 1;

        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, out _, out bool canClose);
            if (canClose && runLen >= minimumRunLength) return i;

            i += runLen - 1;
        }

        return -1;
    }

    private static int FindNextClosingDelimiterRunIndex(string text, int start, char marker, int requiredRunLength) {
        if (string.IsNullOrEmpty(text)) return -1;
        if (requiredRunLength <= 0) requiredRunLength = 1;

        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, out _, out bool canClose);
            if (canClose && runLen == requiredRunLength) return i;

            i += runLen - 1;
        }

        return -1;
    }

    private static void GetDelimiterFlags(string text, int start, char marker, int runLen, out bool canOpen, out bool canClose) {
        canOpen = false;
        canClose = false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return;

        char prev = start > 0 ? text[start - 1] : '\0';
        int nextIndex = start + runLen;
        char next = nextIndex < text.Length ? text[nextIndex] : '\0';

        bool prevWs = prev == '\0' || char.IsWhiteSpace(prev);
        bool nextWs = next == '\0' || char.IsWhiteSpace(next);
        bool prevPunct = prev != '\0' && IsPunctuationOrSymbol(prev);
        bool nextPunct = next != '\0' && IsPunctuationOrSymbol(next);

        bool leftFlanking = !nextWs && (!nextPunct || prevWs || prevPunct);
        bool rightFlanking = !prevWs && (!prevPunct || nextWs || nextPunct);

        if (marker == '~') {
            // Pragmatic GFM-like: "~~" opens/closes when not adjacent to whitespace on the relevant side.
            canOpen = !nextWs;
            canClose = !prevWs;
            return;
        }

        if (marker == '=') {
            // Pragmatic mark/highlight handling: "==" opens/closes when it hugs non-whitespace text.
            canOpen = runLen >= 2 && !nextWs;
            canClose = runLen >= 2 && !prevWs;
            return;
        }

        if (marker == '*') {
            canOpen = leftFlanking;
            canClose = rightFlanking;
            return;
        }

        // '_' is more restrictive (avoid intraword emphasis like foo_bar_baz).
        if (marker == '_') {
            canOpen = leftFlanking && (!rightFlanking || prevPunct || prevWs);
            canClose = rightFlanking && (!leftFlanking || nextPunct || nextWs);
            return;
        }
    }

    private static bool IsPunctuationOrSymbol(char c) => char.IsPunctuation(c) || char.IsSymbol(c);
}
