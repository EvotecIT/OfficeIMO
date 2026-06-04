using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Represents caller-ordered one-based PDF page selections.
/// </summary>
public sealed class PdfPageSelection : IEquatable<PdfPageSelection> {
    private readonly PdfPageRange[] _ranges;

    private PdfPageSelection(PdfPageRange[] ranges) {
        _ranges = (PdfPageRange[])ranges.Clone();
        Ranges = Array.AsReadOnly(_ranges);
    }

    /// <summary>
    /// Inclusive one-based ranges in caller order.
    /// </summary>
    public IReadOnlyList<PdfPageRange> Ranges { get; }

    /// <summary>
    /// Total selected page count before operation-specific de-duplication.
    /// </summary>
    public int PageCount {
        get {
            int count = 0;
            for (int i = 0; i < _ranges.Length; i++) {
                count += _ranges[i].PageCount;
            }

            return count;
        }
    }

    /// <summary>
    /// Creates a selection from one or more one-based page numbers.
    /// </summary>
    public static PdfPageSelection From(params int[] pageNumbers) {
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        if (pageNumbers.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        }

        var ranges = new PdfPageRange[pageNumbers.Length];
        for (int i = 0; i < pageNumbers.Length; i++) {
            ranges[i] = new PdfPageRange(pageNumbers[i], pageNumbers[i]);
        }

        return new PdfPageSelection(ranges);
    }

    /// <summary>
    /// Creates a selection from one inclusive one-based page range.
    /// </summary>
    public static PdfPageSelection From(PdfPageRange pageRange) {
        return new PdfPageSelection(new[] { pageRange });
    }

    /// <summary>
    /// Creates a selection from one or more inclusive one-based page ranges.
    /// </summary>
    public static PdfPageSelection FromRanges(params PdfPageRange[] pageRanges) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        if (pageRanges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", nameof(pageRanges));
        }

        return new PdfPageSelection(pageRanges);
    }

    /// <summary>
    /// Parses comma- or semicolon-separated one-based page ranges such as <c>1</c>, <c>1-3</c>, or <c>3,1-2</c>.
    /// </summary>
    public static PdfPageSelection Parse(string pageRanges) {
        return new PdfPageSelection(PdfPageRange.ParseMany(pageRanges));
    }

    /// <summary>
    /// Attempts to parse comma- or semicolon-separated one-based page ranges.
    /// </summary>
    public static bool TryParse(string? pageRanges, out PdfPageSelection? selection) {
        selection = null;
        if (!PdfPageRange.TryParseMany(pageRanges, out var ranges)) {
            return false;
        }

        selection = new PdfPageSelection(ranges);
        return true;
    }

    internal PdfPageRange[] ToRanges() {
        return (PdfPageRange[])_ranges.Clone();
    }

    internal int[] ToPageNumbers(int pageCount, string paramName) {
        return PdfPageRange.ExpandMany(_ranges, pageCount, paramName);
    }

    /// <inheritdoc />
    public bool Equals(PdfPageSelection? other) {
        if (ReferenceEquals(this, other)) {
            return true;
        }

        if (other is null || _ranges.Length != other._ranges.Length) {
            return false;
        }

        for (int i = 0; i < _ranges.Length; i++) {
            if (_ranges[i] != other._ranges[i]) {
                return false;
            }
        }

        return true;
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) {
        return obj is PdfPageSelection other && Equals(other);
    }

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            for (int i = 0; i < _ranges.Length; i++) {
                hash = (hash * 31) + _ranges[i].GetHashCode();
            }

            return hash;
        }
    }

    /// <inheritdoc />
    public override string ToString() {
        var parts = new string[_ranges.Length];
        for (int i = 0; i < _ranges.Length; i++) {
            PdfPageRange range = _ranges[i];
            parts[i] = range.FirstPage == range.LastPage
                ? range.FirstPage.ToString(CultureInfo.InvariantCulture)
                : range.ToString();
        }

        return string.Join(",", parts);
    }
}
