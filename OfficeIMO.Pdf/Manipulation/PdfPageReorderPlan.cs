namespace OfficeIMO.Pdf;

/// <summary>Describes a complete page permutation before a PDF page-tree mutation is applied.</summary>
public sealed class PdfPageReorderPlan {
    private readonly int[] _positions;

    private PdfPageReorderPlan(int[] pages) {
        SourcePageNumbers = Array.AsReadOnly(pages);
        _positions = new int[pages.Length];
        for (int index = 0; index < pages.Length; index++) {
            _positions[pages[index] - 1] = index + 1;
            if (pages[index] != index + 1) HasChanges = true;
        }
    }

    /// <summary>One-based source page numbers in their proposed output order.</summary>
    public IReadOnlyList<int> SourcePageNumbers { get; }

    /// <summary>Whether the proposed order differs from the source order.</summary>
    public bool HasChanges { get; }

    /// <summary>Returns the one-based output position of a source page.</summary>
    public int GetOutputPageNumber(int sourcePageNumber) {
        if (sourcePageNumber < 1 || sourcePageNumber > _positions.Length)
            throw new ArgumentOutOfRangeException(nameof(sourcePageNumber));
        return _positions[sourcePageNumber - 1];
    }

    /// <summary>
    /// Plans moving selected pages before a source page, preserving their original relative order.
    /// Use page count + 1 to move to the end. The destination cannot be selected.
    /// </summary>
    public static PdfPageReorderPlan Move(int pageCount, int insertBeforePageNumber, params int[] pageNumbers) {
        var selected = ValidateSelection(pageCount, pageNumbers);
        if (insertBeforePageNumber < 1 || (long)insertBeforePageNumber > (long)pageCount + 1)
            throw new ArgumentOutOfRangeException(nameof(insertBeforePageNumber));
        if (selected.Contains(insertBeforePageNumber))
            throw new ArgumentException("Insert-before page cannot be one of the moved pages.", nameof(insertBeforePageNumber));
        var moving = Enumerable.Range(1, pageCount).Where(selected.Contains).ToArray();
        var remaining = Enumerable.Range(1, pageCount).Where(page => !selected.Contains(page)).ToList();
        int position = remaining.FindIndex(page => page >= insertBeforePageNumber);
        remaining.InsertRange(position < 0 ? remaining.Count : position, moving);
        return new PdfPageReorderPlan(remaining.ToArray());
    }

    /// <summary>
    /// Plans shifting each selected run by one position. Runs already at the requested edge remain there.
    /// </summary>
    public static PdfPageReorderPlan Shift(int pageCount, bool towardStart, params int[] pageNumbers) {
        var selected = ValidateSelection(pageCount, pageNumbers);
        int[] order = Enumerable.Range(1, pageCount).ToArray();
        int index = towardStart ? 1 : pageCount - 2;
        int step = towardStart ? 1 : -1;
        for (; index >= 0 && index < pageCount; index += step) {
            int adjacent = index - step;
            if (selected.Contains(order[index]) && !selected.Contains(order[adjacent]))
                (order[index], order[adjacent]) = (order[adjacent], order[index]);
        }
        return new PdfPageReorderPlan(order);
    }

    private static HashSet<int> ValidateSelection(int pageCount, int[] pageNumbers) {
        Guard.PositiveInteger(pageCount, nameof(pageCount));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        if (pageNumbers.Length == 0) throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        var selected = new HashSet<int>();
        foreach (int page in pageNumbers) {
            if (page < 1 || page > pageCount) throw new ArgumentOutOfRangeException(nameof(pageNumbers));
            if (!selected.Add(page)) throw new ArgumentException("Page numbers must be unique.", nameof(pageNumbers));
        }
        return selected;
    }
}
