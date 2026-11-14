using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Aggregate functions that a <see cref="PdfTableAggregate"/> can compute over table body rows.</summary>
public enum PdfTableAggregateFunction {
    /// <summary>Sum of the numeric values.</summary>
    Sum,
    /// <summary>Number of rows that have a value.</summary>
    Count,
    /// <summary>Smallest numeric value. Empty when no row has a value.</summary>
    Minimum,
    /// <summary>Largest numeric value. Empty when no row has a value.</summary>
    Maximum
}

/// <summary>
/// Declares a numeric aggregate over one column of a table's body rows. The aggregate is computed per page
/// fragment after layout and is shown in summary rows through its tokens.
/// </summary>
public sealed class PdfTableAggregate {
    private string? _format;

    /// <summary>Creates an aggregate over a table column.</summary>
    /// <param name="name">Name used in tokens such as <c>{page:name}</c>. Letters, digits, '_', '-' and '.' only.</param>
    /// <param name="column">Zero-based logical column whose body cell text supplies the values.</param>
    /// <param name="function">Aggregate function. Defaults to <see cref="PdfTableAggregateFunction.Sum"/>.</param>
    public PdfTableAggregate(string name, int column, PdfTableAggregateFunction function = PdfTableAggregateFunction.Sum) {
        Guard.NotNull(name, nameof(name));
        if (name.Length == 0 || name.Any(ch => !char.IsLetterOrDigit(ch) && ch != '_' && ch != '-' && ch != '.'))
            throw new ArgumentException("Table aggregate names must be non-empty and contain only letters, digits, '_', '-' or '.'.", nameof(name));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column), "Table aggregate column cannot be negative.");
        if (function is < PdfTableAggregateFunction.Sum or > PdfTableAggregateFunction.Maximum) throw new ArgumentOutOfRangeException(nameof(function));
        Name = name;
        Column = column;
        Function = function;
    }

    /// <summary>Name used in tokens.</summary>
    public string Name { get; }
    /// <summary>Zero-based logical column whose body cell text supplies the values.</summary>
    public int Column { get; }
    /// <summary>Aggregate function.</summary>
    public PdfTableAggregateFunction Function { get; }

    /// <summary>
    /// Standard or custom .NET numeric format applied to the value, for example <c>N2</c> or <c>C</c>.
    /// When null, sums, minimums and maximums use <c>N2</c> and counts use <c>N0</c>.
    /// </summary>
    public string? Format {
        get => _format;
        set {
            if (value != null) _ = 0M.ToString(value, CultureInfo.InvariantCulture);
            _format = value;
        }
    }

    /// <summary>
    /// Culture used to parse body cell text and to format values. When null, the invariant culture is used so
    /// output does not depend on the machine that renders the document.
    /// </summary>
    public IFormatProvider? FormatProvider { get; set; }

    /// <summary>
    /// Optional value source that replaces cell text parsing. It receives the zero-based body row index
    /// (the first row after the header rows is 0) and returns the row's value, or null when the row has none.
    /// </summary>
    public Func<int, decimal?>? ValueSelector { get; set; }

    /// <summary>Token for the value of the rows on the current page fragment.</summary>
    public string PageToken => "{page:" + Name + "}";
    /// <summary>Token for the value of the rows on all earlier page fragments (brought forward).</summary>
    public string BroughtForwardToken => "{broughtForward:" + Name + "}";
    /// <summary>
    /// Token for the value of all rows up to the end of the current page fragment. In a page footer row it is the
    /// carried-forward amount; in the table's footer rows on the last page it is the grand total.
    /// </summary>
    public string RunningToken => "{running:" + Name + "}";

    internal PdfTableAggregate Clone() => new PdfTableAggregate(Name, Column, Function) {
        _format = _format,
        FormatProvider = FormatProvider,
        ValueSelector = ValueSelector
    };

    internal IFormatProvider Culture => FormatProvider ?? CultureInfo.InvariantCulture;
}

/// <summary>
/// Per-page table summaries: column aggregates and the rows that show them at the bottom of every page fragment
/// that continues on the next page (for example "Carried forward" or "Page subtotal") and at the top of every
/// continuation fragment (for example "Brought forward").
/// </summary>
/// <remarks>
/// Values reflect the body rows that actually landed on each page after layout. A body row counts toward the page
/// where its first line is drawn. Summary rows are measured with the values they will show, so they never change
/// which rows fit on a page after the values are known. Cells in summary rows and in the table's footer rows can
/// contain the tokens of any declared aggregate (<see cref="PdfTableAggregate.PageToken"/>,
/// <see cref="PdfTableAggregate.BroughtForwardToken"/> and <see cref="PdfTableAggregate.RunningToken"/>).
/// Summary rows use the table's footer styling.
/// </remarks>
public sealed class PdfTablePageSummary {
    /// <summary>Aggregates available to summary row tokens.</summary>
    public System.Collections.Generic.List<PdfTableAggregate> Aggregates { get; } = new();

    /// <summary>
    /// Rows drawn at the bottom of every page fragment that continues on a later page. They are not drawn on the
    /// table's final fragment, where the table's own footer rows close the table.
    /// </summary>
    public System.Collections.Generic.List<PdfTableCell[]> PageFooterRows { get; } = new();

    /// <summary>
    /// Rows drawn at the top of every continuation fragment, after the repeated header rows. They are drawn only
    /// when at least one body row appeared on an earlier page. They are laid out before the page's rows are known,
    /// so their cells can use <see cref="PdfTableAggregate.BroughtForwardToken"/> but not page or running tokens.
    /// </summary>
    public System.Collections.Generic.List<PdfTableCell[]> ContinuationHeaderRows { get; } = new();

    /// <summary>Declares a sum aggregate and returns it so its tokens can be used in rows.</summary>
    public PdfTableAggregate Sum(string name, int column, string? format = null, IFormatProvider? formatProvider = null) =>
        Add(new PdfTableAggregate(name, column) { Format = format, FormatProvider = formatProvider });

    /// <summary>Adds an aggregate and returns it.</summary>
    public PdfTableAggregate Add(PdfTableAggregate aggregate) {
        Guard.NotNull(aggregate, nameof(aggregate));
        if (Aggregates.Any(existing => string.Equals(existing.Name, aggregate.Name, StringComparison.Ordinal)))
            throw new ArgumentException("A table aggregate with the same name is already declared.", nameof(aggregate));
        Aggregates.Add(aggregate);
        return aggregate;
    }

    /// <summary>Adds a row drawn at the bottom of every page fragment that continues on a later page.</summary>
    public PdfTablePageSummary AddPageFooterRow(params string[] cells) => AddPageFooterRow(ToCells(cells));

    /// <summary>Adds a row drawn at the bottom of every page fragment that continues on a later page.</summary>
    public PdfTablePageSummary AddPageFooterRow(params PdfTableCell[] cells) {
        PageFooterRows.Add(SnapshotRow(cells));
        return this;
    }

    /// <summary>Adds a row drawn at the top of every continuation fragment.</summary>
    public PdfTablePageSummary AddContinuationHeaderRow(params string[] cells) => AddContinuationHeaderRow(ToCells(cells));

    /// <summary>Adds a row drawn at the top of every continuation fragment.</summary>
    public PdfTablePageSummary AddContinuationHeaderRow(params PdfTableCell[] cells) {
        ContinuationHeaderRows.Add(SnapshotRow(cells));
        return this;
    }

    internal PdfTablePageSummary Clone() {
        var clone = new PdfTablePageSummary();
        foreach (PdfTableAggregate aggregate in Aggregates) clone.Aggregates.Add(aggregate.Clone());
        foreach (PdfTableCell[] row in PageFooterRows) clone.PageFooterRows.Add(SnapshotRow(row));
        foreach (PdfTableCell[] row in ContinuationHeaderRows) clone.ContinuationHeaderRows.Add(SnapshotRow(row));
        return clone;
    }

    private static PdfTableCell[] ToCells(string[] cells) {
        Guard.NotNull(cells, nameof(cells));
        return cells.Select(text => new PdfTableCell(text)).ToArray();
    }

    private static PdfTableCell[] SnapshotRow(PdfTableCell[] cells) {
        Guard.NotNull(cells, nameof(cells));
        if (cells.Any(cell => cell is null)) throw new ArgumentException("Summary rows cannot contain null cells.", nameof(cells));
        return cells.Select(cell => cell.Clone()).ToArray();
    }
}
