using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a document-relative, caller-ordered page selection expression.
/// </summary>
/// <remarks>
/// Selectors are resolved against a document page count. Supported terms include absolute pages and
/// ranges (<c>1</c>, <c>1-3</c>, <c>5..2</c>), end-relative pages (<c>last</c>, <c>last-2</c>),
/// <c>all</c>, <c>odd</c>, <c>even</c>, and exclusions prefixed with <c>!</c>.
/// </remarks>
public sealed class PdfPageSelector : IEquatable<PdfPageSelector> {
    private static readonly char[] TermSeparators = { ',', ';' };
    private readonly string _expression;
    private readonly SelectorTerm[] _terms;

    private PdfPageSelector(string expression, SelectorTerm[] terms) {
        _expression = expression;
        _terms = terms;
    }

    /// <summary>Original normalized selector expression.</summary>
    public string Expression => _expression;

    /// <summary>Parses a document-relative page selector.</summary>
    public static PdfPageSelector Parse(string expression) {
        Guard.NotNull(expression, nameof(expression));
        string normalized = expression.Trim();
        if (normalized.Length == 0) {
            throw new ArgumentException("Page selector expression cannot be empty.", nameof(expression));
        }

        string[] rawTerms = normalized.Split(TermSeparators, StringSplitOptions.None);
        var terms = new SelectorTerm[rawTerms.Length];
        for (int i = 0; i < rawTerms.Length; i++) {
            string token = rawTerms[i].Trim();
            if (token.Length == 0) {
                throw new ArgumentException("Page selector expression contains an empty term.", nameof(expression));
            }

            bool exclude = token[0] == '!';
            if (exclude) {
                token = token.Substring(1).Trim();
                if (token.Length == 0) {
                    throw new ArgumentException("Page selector exclusion cannot be empty.", nameof(expression));
                }
            }

            terms[i] = ParseTerm(token, exclude, nameof(expression));
        }

        return new PdfPageSelector(normalized, terms);
    }

    /// <summary>Attempts to parse a document-relative page selector.</summary>
    public static bool TryParse(string? expression, out PdfPageSelector? selector) {
        selector = null;
        if (expression is null) {
            return false;
        }

        try {
            selector = Parse(expression);
            return true;
        } catch (ArgumentException) {
            return false;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

    /// <summary>Resolves the selector to caller-ordered, one-based page numbers.</summary>
    public IReadOnlyList<int> Resolve(int pageCount) {
        if (pageCount < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageCount), "Document page count must be 1 or greater.");
        }

        var excluded = new HashSet<int>();
        bool hasIncludes = false;
        for (int i = 0; i < _terms.Length; i++) {
            SelectorTerm term = _terms[i];
            if (!term.Exclude) {
                hasIncludes = true;
                continue;
            }

            foreach (int page in term.Resolve(pageCount)) {
                excluded.Add(page);
            }
        }

        var pages = new List<int>();
        if (!hasIncludes) {
            for (int page = 1; page <= pageCount; page++) {
                if (!excluded.Contains(page)) {
                    pages.Add(page);
                }
            }
        } else {
            for (int i = 0; i < _terms.Length; i++) {
                SelectorTerm term = _terms[i];
                if (term.Exclude) {
                    continue;
                }

                foreach (int page in term.Resolve(pageCount)) {
                    if (!excluded.Contains(page)) {
                        pages.Add(page);
                    }
                }
            }
        }

        if (pages.Count == 0) {
            throw new InvalidOperationException("Page selector resolved to an empty page set.");
        }

        return pages.AsReadOnly();
    }

    /// <summary>Resolves the selector to an absolute page selection.</summary>
    public PdfPageSelection ResolveSelection(int pageCount) {
        return PdfPageSelection.From(Resolve(pageCount).ToArray());
    }

    /// <inheritdoc />
    public bool Equals(PdfPageSelector? other) {
        return other is not null && string.Equals(_expression, other._expression, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) {
        return obj is PdfPageSelector other && Equals(other);
    }

    /// <inheritdoc />
    public override int GetHashCode() {
        return StringComparer.OrdinalIgnoreCase.GetHashCode(_expression);
    }

    /// <inheritdoc />
    public override string ToString() {
        return _expression;
    }

    private static SelectorTerm ParseTerm(string token, bool exclude, string paramName) {
        if (token.Equals("all", StringComparison.OrdinalIgnoreCase)) {
            return SelectorTerm.All(exclude);
        }

        if (token.Equals("odd", StringComparison.OrdinalIgnoreCase)) {
            return SelectorTerm.Odd(exclude);
        }

        if (token.Equals("even", StringComparison.OrdinalIgnoreCase)) {
            return SelectorTerm.Even(exclude);
        }

        int dotRange = token.IndexOf("..", StringComparison.Ordinal);
        if (dotRange >= 0) {
            if (token.IndexOf("..", dotRange + 2, StringComparison.Ordinal) >= 0) {
                throw new ArgumentException("Page selector range must contain exactly two endpoints.", paramName);
            }

            string first = token.Substring(0, dotRange).Trim();
            string last = token.Substring(dotRange + 2).Trim();
            return SelectorTerm.Range(ParseEndpoint(first, paramName), ParseEndpoint(last, paramName), exclude);
        }

        int numericDash = FindNumericRangeDash(token);
        if (numericDash > 0) {
            string first = token.Substring(0, numericDash).Trim();
            string last = token.Substring(numericDash + 1).Trim();
            return SelectorTerm.Range(ParseEndpoint(first, paramName), ParseEndpoint(last, paramName), exclude);
        }

        return SelectorTerm.Page(ParseEndpoint(token, paramName), exclude);
    }

    private static int FindNumericRangeDash(string token) {
        for (int i = 1; i < token.Length - 1; i++) {
            if (token[i] == '-' && char.IsDigit(token[i - 1]) && char.IsDigit(token[i + 1])) {
                return i;
            }
        }

        return -1;
    }

    private static PageEndpoint ParseEndpoint(string text, string paramName) {
        if (text.Length == 0) {
            throw new ArgumentException("Page selector endpoint cannot be empty.", paramName);
        }

        if (int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int absolutePage)) {
            if (absolutePage < 1) {
                throw new ArgumentOutOfRangeException(paramName, "Absolute page numbers must be 1 or greater.");
            }

            return PageEndpoint.Absolute(absolutePage);
        }

        string lowered = text.ToLowerInvariant();
        string alias;
        if (lowered.StartsWith("last", StringComparison.Ordinal)) {
            alias = "last";
        } else if (lowered.StartsWith("end", StringComparison.Ordinal)) {
            alias = "end";
        } else if (lowered.Length > 0 && lowered[0] == 'z') {
            alias = "z";
        } else {
            throw new FormatException("Page selector endpoint must be a positive page number or last-page alias.");
        }

        string suffix = lowered.Substring(alias.Length);
        if (suffix.Length == 0) {
            return PageEndpoint.FromEnd(0);
        }

        if (!TryParseLastPageOffset(suffix, out int offset)) {
            throw new FormatException("Last-page offsets must use the form last-N.");
        }

        return PageEndpoint.FromEnd(offset);
    }

    private static bool TryParseLastPageOffset(string suffix, out int offset) {
        offset = 0;
        if (suffix.Length < 2 || suffix[0] != '-') {
            return false;
        }

        try {
            checked {
                for (int i = 1; i < suffix.Length; i++) {
                    char digit = suffix[i];
                    if (digit < '0' || digit > '9') {
                        return false;
                    }

                    offset = (offset * 10) + digit - '0';
                }
            }
        } catch (OverflowException) {
            return false;
        }

        return true;
    }

    private readonly struct PageEndpoint {
        private PageEndpoint(int value, bool fromEnd) {
            Value = value;
            IsFromEnd = fromEnd;
        }

        public int Value { get; }
        public bool IsFromEnd { get; }

        public static PageEndpoint Absolute(int page) => new PageEndpoint(page, fromEnd: false);
        public static PageEndpoint FromEnd(int offset) => new PageEndpoint(offset, fromEnd: true);

        public int Resolve(int pageCount) {
            int page = IsFromEnd ? pageCount - Value : Value;
            if (page < 1 || page > pageCount) {
                throw new ArgumentOutOfRangeException(nameof(pageCount), "Page selector endpoint falls outside the document page count.");
            }

            return page;
        }
    }

    private enum SelectorTermKind {
        Page,
        Range,
        All,
        Odd,
        Even
    }

    private readonly struct SelectorTerm {
        private SelectorTerm(SelectorTermKind kind, PageEndpoint first, PageEndpoint last, bool exclude) {
            Kind = kind;
            First = first;
            Last = last;
            Exclude = exclude;
        }

        public SelectorTermKind Kind { get; }
        public PageEndpoint First { get; }
        public PageEndpoint Last { get; }
        public bool Exclude { get; }

        public static SelectorTerm Page(PageEndpoint page, bool exclude) => new SelectorTerm(SelectorTermKind.Page, page, page, exclude);
        public static SelectorTerm Range(PageEndpoint first, PageEndpoint last, bool exclude) => new SelectorTerm(SelectorTermKind.Range, first, last, exclude);
        public static SelectorTerm All(bool exclude) => new SelectorTerm(SelectorTermKind.All, default, default, exclude);
        public static SelectorTerm Odd(bool exclude) => new SelectorTerm(SelectorTermKind.Odd, default, default, exclude);
        public static SelectorTerm Even(bool exclude) => new SelectorTerm(SelectorTermKind.Even, default, default, exclude);

        public IEnumerable<int> Resolve(int pageCount) {
            if (Kind == SelectorTermKind.All) {
                for (int page = 1; page <= pageCount; page++) {
                    yield return page;
                }

                yield break;
            }

            if (Kind == SelectorTermKind.Odd || Kind == SelectorTermKind.Even) {
                int first = Kind == SelectorTermKind.Odd ? 1 : 2;
                for (int page = first; page <= pageCount; page += 2) {
                    yield return page;
                }

                yield break;
            }

            int resolvedFirst = First.Resolve(pageCount);
            int resolvedLast = Last.Resolve(pageCount);
            int direction = resolvedFirst <= resolvedLast ? 1 : -1;
            for (int page = resolvedFirst; ; page += direction) {
                yield return page;
                if (page == resolvedLast) {
                    yield break;
                }
            }
        }
    }
}
