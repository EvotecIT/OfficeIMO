namespace OfficeIMO.Pdf;

/// <summary>
/// Represents an inclusive one-based PDF page range.
/// </summary>
public readonly struct PdfPageRange : System.IEquatable<PdfPageRange> {
    private static readonly char[] RangeListSeparators = { ',', ';' };
    private static readonly string[] DotRangeSeparator = { ".." };
    private static readonly char[] DashRangeSeparator = { '-' };

    /// <summary>
    /// Creates an inclusive one-based PDF page range.
    /// </summary>
    public PdfPageRange(int firstPage, int lastPage) {
        if (firstPage < 1) {
            throw new System.ArgumentOutOfRangeException(nameof(firstPage), "First page must be 1 or greater.");
        }

        if (lastPage < firstPage) {
            throw new System.ArgumentOutOfRangeException(nameof(lastPage), "Last page must be greater than or equal to first page.");
        }

        FirstPage = firstPage;
        LastPage = lastPage;
    }

    /// <summary>
    /// First page in the inclusive one-based range.
    /// </summary>
    public int FirstPage { get; }

    /// <summary>
    /// Last page in the inclusive one-based range.
    /// </summary>
    public int LastPage { get; }

    /// <summary>
    /// Number of pages in the range.
    /// </summary>
    public int PageCount => LastPage - FirstPage + 1;

    /// <summary>
    /// Creates an inclusive one-based PDF page range.
    /// </summary>
    public static PdfPageRange From(int firstPage, int lastPage) {
        return new PdfPageRange(firstPage, lastPage);
    }

    /// <summary>
    /// Parses one inclusive one-based page range from text such as <c>3</c>, <c>1-3</c>, or <c>1..3</c>.
    /// </summary>
    public static PdfPageRange Parse(string pageRange) {
        Guard.NotNull(pageRange, nameof(pageRange));

        string token = pageRange.Trim();
        if (token.Length == 0) {
            throw new System.ArgumentException("Page range text cannot be empty.", nameof(pageRange));
        }

        return ParseToken(token, nameof(pageRange));
    }

    /// <summary>
    /// Attempts to parse one inclusive one-based page range from text such as <c>3</c>, <c>1-3</c>, or <c>1..3</c>.
    /// </summary>
    public static bool TryParse(string? pageRange, out PdfPageRange range) {
        range = default;
        if (pageRange is null) {
            return false;
        }

        string token = pageRange.Trim();
        if (token.Length == 0) {
            return false;
        }

        try {
            range = ParseToken(token, nameof(pageRange));
            return true;
        } catch (System.ArgumentException) {
            return false;
        } catch (System.FormatException) {
            return false;
        } catch (System.OverflowException) {
            return false;
        }
    }

    /// <summary>
    /// Parses comma- or semicolon-separated inclusive one-based page ranges, preserving caller order.
    /// </summary>
    public static PdfPageRange[] ParseMany(string pageRanges) {
        Guard.NotNull(pageRanges, nameof(pageRanges));

        string trimmed = pageRanges.Trim();
        if (trimmed.Length == 0) {
            throw new System.ArgumentException("Page range text cannot be empty.", nameof(pageRanges));
        }

        string[] tokens = trimmed.Split(RangeListSeparators, System.StringSplitOptions.None);
        var ranges = new System.Collections.Generic.List<PdfPageRange>(tokens.Length);
        for (int i = 0; i < tokens.Length; i++) {
            string token = tokens[i].Trim();
            if (token.Length == 0) {
                throw new System.ArgumentException("Page range list contains an empty range.", nameof(pageRanges));
            }

            ranges.Add(ParseToken(token, nameof(pageRanges)));
        }

        return ranges.ToArray();
    }

    /// <summary>
    /// Attempts to parse comma- or semicolon-separated inclusive one-based page ranges, preserving caller order.
    /// </summary>
    public static bool TryParseMany(string? pageRanges, out PdfPageRange[] ranges) {
        ranges = System.Array.Empty<PdfPageRange>();
        if (pageRanges is null) {
            return false;
        }

        try {
            ranges = ParseMany(pageRanges);
            return true;
        } catch (System.ArgumentException) {
            ranges = System.Array.Empty<PdfPageRange>();
            return false;
        } catch (System.FormatException) {
            ranges = System.Array.Empty<PdfPageRange>();
            return false;
        } catch (System.OverflowException) {
            ranges = System.Array.Empty<PdfPageRange>();
            return false;
        }
    }

    internal int[] ToPageNumbers() {
        int[] pages = new int[PageCount];
        for (int i = 0; i < pages.Length; i++) {
            pages[i] = FirstPage + i;
        }

        return pages;
    }

    internal static int[] ExpandMany(PdfPageRange[]? pageRanges, int pageCount, string paramName) {
        if (pageRanges is null) {
            throw new System.ArgumentNullException(paramName);
        }

        if (pageRanges.Length == 0) {
            throw new System.ArgumentException("At least one page range must be specified.", paramName);
        }

        var pages = new System.Collections.Generic.List<int>();
        for (int i = 0; i < pageRanges.Length; i++) {
            if (pageRanges[i].FirstPage < 1 || pageRanges[i].LastPage < pageRanges[i].FirstPage) {
                throw new System.ArgumentOutOfRangeException(paramName, "Page ranges must be inclusive one-based ranges.");
            }

            if (pageRanges[i].LastPage > pageCount) {
                throw new System.ArgumentOutOfRangeException(paramName, "Page range cannot exceed the document page count.");
            }

            for (int pageNumber = pageRanges[i].FirstPage; pageNumber <= pageRanges[i].LastPage; pageNumber++) {
                pages.Add(pageNumber);
            }
        }

        return pages.ToArray();
    }

    /// <summary>
    /// Compares two page ranges for equality.
    /// </summary>
    public static bool operator ==(PdfPageRange left, PdfPageRange right) {
        return left.Equals(right);
    }

    /// <summary>
    /// Compares two page ranges for inequality.
    /// </summary>
    public static bool operator !=(PdfPageRange left, PdfPageRange right) {
        return !left.Equals(right);
    }

    /// <inheritdoc />
    public bool Equals(PdfPageRange other) {
        return FirstPage == other.FirstPage && LastPage == other.LastPage;
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) {
        return obj is PdfPageRange other && Equals(other);
    }

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            return (FirstPage * 397) ^ LastPage;
        }
    }

    /// <inheritdoc />
    public override string ToString() {
        return FirstPage.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + LastPage.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static PdfPageRange ParseToken(string token, string paramName) {
        string[] parts;
        if (token.Contains("..")) {
            parts = token.Split(DotRangeSeparator, System.StringSplitOptions.None);
        } else if (token.Contains('-')) {
            parts = token.Split(DashRangeSeparator, System.StringSplitOptions.None);
        } else {
            int page = ParsePageNumber(token, paramName);
            return new PdfPageRange(page, page);
        }

        if (parts.Length != 2) {
            throw new System.ArgumentException("Page range must be a single page or an inclusive first-last range.", paramName);
        }

        int firstPage = ParsePageNumber(parts[0].Trim(), paramName);
        int lastPage = ParsePageNumber(parts[1].Trim(), paramName);
        return new PdfPageRange(firstPage, lastPage);
    }

    private static int ParsePageNumber(string text, string paramName) {
        if (text.Length == 0) {
            throw new System.ArgumentException("Page number cannot be empty.", paramName);
        }

        if (!int.TryParse(text, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int pageNumber)) {
            throw new System.FormatException("Page number must be an invariant-culture integer.");
        }

        return pageNumber;
    }
}
