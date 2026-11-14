namespace OfficeIMO.Markdown;

/// <summary>
/// Source points of joined paragraph text kept as runs: one run per joined line, whose points
/// advance one character at a time, and one single-point run per line separator. A point is
/// computed when asked for, so a long paragraph does not hold a point per character.
/// </summary>
internal sealed class MarkdownSourcePointRuns {
    private Run[] _runs = new Run[4];
    private int _runCount;

    internal int Count { get; private set; }

    internal void AddPoint(MarkdownSourcePoint point) {
        Add(new Run {
            Start = Count,
            Length = 1,
            Line = point.Line,
            FirstColumn = point.Column,
            FirstOffset = point.Offset
        });
    }

    /// <summary>Adds the points <see cref="MarkdownSourceTextMap.WriteSequentialPoints"/> would write and returns the same column.</summary>
    internal int AddSequential(MarkdownSourceTextMap sourceTextMap, string text, int line, int startColumn) {
        if (string.IsNullOrEmpty(text)) {
            return startColumn;
        }

        var range = sourceTextMap.CreateSequentialRange(line, startColumn);
        int[]? logicalColumns = null;
        int endColumn;
        if (text.IndexOf('\t') >= 0) {
            logicalColumns = new int[text.Length + 1];
            logicalColumns[0] = startColumn;
            for (var index = 0; index < text.Length; index++) {
                logicalColumns[index + 1] = MarkdownSourceColumns.AdvanceColumn(logicalColumns[index], text[index]);
            }

            endColumn = logicalColumns[text.Length];
        } else {
            endColumn = startColumn + text.Length;
        }

        Add(new Run {
            Start = Count,
            Length = text.Length,
            Line = range.FirstPoint.Line,
            FirstColumn = range.FirstPoint.Column,
            FirstOffset = range.FirstPoint.Offset,
            LogicalStartColumn = startColumn,
            MaximumColumn = range.MaximumColumn,
            MaximumOffset = range.MaximumOffset,
            LogicalColumns = logicalColumns,
            Sequential = true
        });
        return endColumn;
    }

    internal MarkdownSourcePoint Get(int index) {
        int low = 0;
        int high = _runCount - 1;
        while (low < high) {
            int middle = (low + high + 1) >> 1;
            if (_runs[middle].Start <= index) {
                low = middle;
            } else {
                high = middle - 1;
            }
        }

        ref readonly var run = ref _runs[low];
        int step = index - run.Start;
        if (step == 0 || !run.Sequential) {
            return new MarkdownSourcePoint(run.Line, run.FirstColumn, run.FirstOffset);
        }

        int logicalColumn = run.LogicalColumns != null
            ? run.LogicalColumns[step]
            : run.LogicalStartColumn + step;
        return new MarkdownSourcePoint(
            run.Line,
            Math.Min(run.MaximumColumn, Math.Max(1, logicalColumn)),
            Math.Min(run.MaximumOffset, run.FirstOffset + step));
    }

    private void Add(Run run) {
        if (_runCount == _runs.Length) {
            Array.Resize(ref _runs, _runs.Length * 2);
        }

        _runs[_runCount++] = run;
        Count += run.Length;
    }

    private struct Run {
        public int Start;
        public int Length;
        public int Line;
        public int FirstColumn;
        public int FirstOffset;
        public int LogicalStartColumn;
        public int MaximumColumn;
        public int MaximumOffset;
        public int[]? LogicalColumns;
        public bool Sequential;
    }
}
