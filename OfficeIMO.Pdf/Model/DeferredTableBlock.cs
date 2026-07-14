namespace OfficeIMO.Pdf;

/// <summary>
/// Stores a replayable table row factory and materializes only a bounded group of rows for layout.
/// </summary>
internal sealed class DeferredTableBlock : IPdfBlock {
    private readonly System.Func<System.Collections.Generic.IEnumerable<PdfTableCell[]>> _rowFactory;

    internal DeferredTableBlock(
        System.Func<System.Collections.Generic.IEnumerable<PdfTableCell[]>> rowFactory,
        int batchSize,
        PdfAlign align,
        PdfTableStyle? style) {
        Guard.NotNull(rowFactory, nameof(rowFactory));
        if (batchSize <= 0) {
            throw new System.ArgumentOutOfRangeException(nameof(batchSize), "Deferred table batch size must be greater than zero.");
        }

        Guard.LeftCenterRightAlign(align, nameof(align), "Deferred table");
        _rowFactory = rowFactory;
        BatchSize = batchSize;
        Align = align;
        Style = style?.Clone();
    }

    internal int BatchSize { get; }
    internal PdfAlign Align { get; }
    internal PdfTableStyle? Style { get; }

    internal System.Collections.Generic.IEnumerable<PdfTableCell[]> EnumerateRows() {
        System.Collections.Generic.IEnumerable<PdfTableCell[]>? rows = _rowFactory();
        if (rows == null) {
            throw new System.InvalidOperationException("Deferred table row factory returned null.");
        }

        foreach (PdfTableCell[]? row in rows) {
            if (row == null) {
                throw new System.ArgumentException("Deferred table rows cannot contain null entries.", nameof(_rowFactory));
            }

            for (int cellIndex = 0; cellIndex < row.Length; cellIndex++) {
                if (row[cellIndex] == null) {
                    throw new System.ArgumentException("Deferred table cells cannot contain null entries.", nameof(_rowFactory));
                }
            }

            yield return row;
        }
    }

    internal System.Collections.Generic.IEnumerable<DeferredTableBatch> CreateBatches(PdfTableStyle effectiveStyle) {
        Guard.NotNull(effectiveStyle, nameof(effectiveStyle));
        if (effectiveStyle.AutoFitColumns) {
            throw new System.ArgumentException("Deferred tables cannot use global automatic column fitting. Configure fixed or weighted column widths so layout remains bounded.", nameof(effectiveStyle));
        }

        int headerRowCount = effectiveStyle.HeaderRowCount;
        int footerRowCount = effectiveStyle.FooterRowCount;
        var headers = new System.Collections.Generic.List<IndexedTableRow>(headerRowCount);
        var trailingRows = new System.Collections.Generic.Queue<IndexedTableRow>(footerRowCount + 1);
        var bodyRows = new System.Collections.Generic.List<IndexedTableRow>(BatchSize);
        bool emittedBatch = false;
        int sourceRowIndex = 0;
        int? resolvedColumnCount = null;

        using (System.Collections.Generic.IEnumerator<PdfTableCell[]> enumerator = EnumerateRows().GetEnumerator()) {
            while (headers.Count < headerRowCount && enumerator.MoveNext()) {
                headers.Add(new IndexedTableRow(sourceRowIndex++, enumerator.Current));
            }

            if (headers.Count > 0 && headers.Count < headerRowCount) {
                throw new System.ArgumentException("Deferred table header row count exceeds the number of supplied rows.", nameof(effectiveStyle));
            }

            while (enumerator.MoveNext()) {
                trailingRows.Enqueue(new IndexedTableRow(sourceRowIndex++, enumerator.Current));
                if (trailingRows.Count <= footerRowCount) {
                    continue;
                }

                IndexedTableRow bodyRow = trailingRows.Dequeue();
                if (bodyRows.Count == BatchSize) {
                    DeferredTableBatch batch = CreateBatch(headers, bodyRows, System.Array.Empty<IndexedTableRow>(), effectiveStyle, isFirst: !emittedBatch, isLast: false);
                    ValidateColumnCount(batch, ref resolvedColumnCount);
                    yield return batch;
                    emittedBatch = true;
                    bodyRows = new System.Collections.Generic.List<IndexedTableRow>(BatchSize);
                }

                bodyRows.Add(bodyRow);
            }
        }

        if (headers.Count == 0 && bodyRows.Count == 0 && trailingRows.Count == 0) {
            yield break;
        }

        if (trailingRows.Count < footerRowCount) {
            throw new System.ArgumentException("Deferred table header and footer row counts exceed the number of supplied rows.", nameof(effectiveStyle));
        }

        IndexedTableRow[] footers = trailingRows.ToArray();
        DeferredTableBatch finalBatch = CreateBatch(headers, bodyRows, footers, effectiveStyle, isFirst: !emittedBatch, isLast: true);
        ValidateColumnCount(finalBatch, ref resolvedColumnCount);
        yield return finalBatch;
    }

    private static void ValidateColumnCount(DeferredTableBatch batch, ref int? expectedColumnCount) {
        if (!expectedColumnCount.HasValue) {
            expectedColumnCount = batch.Table.ColumnCount;
            return;
        }

        if (batch.Table.ColumnCount != expectedColumnCount.Value) {
            throw new System.ArgumentException("Deferred table batches must resolve to a consistent column count.");
        }
    }

    private DeferredTableBatch CreateBatch(
        System.Collections.Generic.IReadOnlyList<IndexedTableRow> headers,
        System.Collections.Generic.IReadOnlyList<IndexedTableRow> bodyRows,
        System.Collections.Generic.IReadOnlyList<IndexedTableRow> footers,
        PdfTableStyle effectiveStyle,
        bool isFirst,
        bool isLast) {
        var rows = new System.Collections.Generic.List<PdfTableCell[]>(headers.Count + bodyRows.Count + footers.Count);
        var sourceIndexes = new System.Collections.Generic.List<int>(rows.Capacity);
        AddRows(headers, rows, sourceIndexes);
        AddRows(bodyRows, rows, sourceIndexes);
        AddRows(footers, rows, sourceIndexes);

        PdfTableStyle batchStyle = CreateBatchStyle(effectiveStyle, sourceIndexes, isFirst, isLast, footers.Count);
        var table = new TableBlock(rows, Align, batchStyle);
        int bodyRowOffset = bodyRows.Count == 0 ? 0 : bodyRows[0].SourceIndex - headers.Count;
        return new DeferredTableBatch(table, isFirst, isLast, bodyRowOffset);
    }

    private static void AddRows(
        System.Collections.Generic.IReadOnlyList<IndexedTableRow> source,
        System.Collections.Generic.List<PdfTableCell[]> rows,
        System.Collections.Generic.List<int> sourceIndexes) {
        for (int index = 0; index < source.Count; index++) {
            rows.Add(source[index].Cells);
            sourceIndexes.Add(source[index].SourceIndex);
        }
    }

    private static PdfTableStyle CreateBatchStyle(
        PdfTableStyle source,
        System.Collections.Generic.IReadOnlyList<int> sourceIndexes,
        bool isFirst,
        bool isLast,
        int footerRowCount) {
        PdfTableStyle style = source.Clone();
        style.Caption = isFirst ? source.Caption : null;
        style.SpacingBefore = isFirst ? source.SpacingBefore : 0D;
        style.SpacingAfter = isLast ? source.SpacingAfter : 0D;
        style.FooterRowCount = isLast ? footerRowCount : 0;
        style.MinimumBodyRowsOnLastPage = isLast ? source.MinimumBodyRowsOnLastPage : 0;
        style.KeepTogether = isFirst && isLast && source.KeepTogether;
        style.KeepWithNext = isLast && source.KeepWithNext;
        style.RowMinHeights = RemapList(source.RowMinHeights, sourceIndexes);
        style.FixedRowHeights = RemapList(source.FixedRowHeights, sourceIndexes);
        style.RowAllowBreakAcrossPages = RemapList(source.RowAllowBreakAcrossPages, sourceIndexes);
        style.CellFills = RemapDictionary(source.CellFills, sourceIndexes);
        style.CellDataBars = RemapDictionary(source.CellDataBars, sourceIndexes);
        style.CellIcons = RemapDictionary(source.CellIcons, sourceIndexes);
        style.CellBorders = RemapDictionary(source.CellBorders, sourceIndexes);
        style.CellPaddings = RemapDictionary(source.CellPaddings, sourceIndexes);
        style.CellAlignments = RemapDictionary(source.CellAlignments, sourceIndexes);
        style.CellVerticalAlignments = RemapDictionary(source.CellVerticalAlignments, sourceIndexes);
        return style;
    }

    private static System.Collections.Generic.List<T>? RemapList<T>(
        System.Collections.Generic.IReadOnlyList<T>? source,
        System.Collections.Generic.IReadOnlyList<int> sourceIndexes) {
        if (source == null) {
            return null;
        }

        var mapped = new System.Collections.Generic.List<T>(sourceIndexes.Count);
        for (int index = 0; index < sourceIndexes.Count; index++) {
            int sourceIndex = sourceIndexes[index];
            mapped.Add(sourceIndex < source.Count ? source[sourceIndex] : default!);
        }

        return mapped;
    }

    private static System.Collections.Generic.Dictionary<(int Row, int Column), TValue>? RemapDictionary<TValue>(
        System.Collections.Generic.IReadOnlyDictionary<(int Row, int Column), TValue>? source,
        System.Collections.Generic.IReadOnlyList<int> sourceIndexes) {
        if (source == null) {
            return null;
        }

        var sourceToLocal = new System.Collections.Generic.Dictionary<int, int>();
        for (int localIndex = 0; localIndex < sourceIndexes.Count; localIndex++) {
            sourceToLocal[sourceIndexes[localIndex]] = localIndex;
        }

        var mapped = new System.Collections.Generic.Dictionary<(int Row, int Column), TValue>();
        foreach (System.Collections.Generic.KeyValuePair<(int Row, int Column), TValue> entry in source) {
            if (sourceToLocal.TryGetValue(entry.Key.Row, out int localRow)) {
                mapped[(localRow, entry.Key.Column)] = entry.Value;
            }
        }

        return mapped;
    }

    private readonly struct IndexedTableRow {
        internal IndexedTableRow(int sourceIndex, PdfTableCell[] cells) {
            SourceIndex = sourceIndex;
            Cells = cells;
        }

        internal int SourceIndex { get; }
        internal PdfTableCell[] Cells { get; }
    }
}

internal sealed class DeferredTableBatch {
    internal DeferredTableBatch(TableBlock table, bool isFirst, bool isLast, int bodyRowOffset) {
        Table = table;
        IsFirst = isFirst;
        IsLast = isLast;
        BodyRowOffset = bodyRowOffset;
    }

    internal TableBlock Table { get; }
    internal bool IsFirst { get; }
    internal bool IsLast { get; }
    internal int BodyRowOffset { get; }
}
