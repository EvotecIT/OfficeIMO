using System.Text;

#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.TextAnalysis;
#else
namespace OfficeIMO.TextAnalysis;
#endif

/// <summary>How adjacent query clauses without an explicit operator combine.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
enum OfficeTextQueryOperator {
    /// <summary>Every clause must match.</summary>
    And,
    /// <summary>Any clause may match; documents matching more clauses score higher.</summary>
    Or
}

/// <summary>The kind of an <see cref="OfficeTextQuery"/> node.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
enum OfficeTextQueryKind {
    /// <summary>One word.</summary>
    Term,
    /// <summary>Words that must appear consecutively in one field.</summary>
    Phrase,
    /// <summary>Any term that starts with the given text.</summary>
    Prefix,
    /// <summary>All children must match.</summary>
    And,
    /// <summary>At least one child must match.</summary>
    Or
}

/// <summary>
/// A query over an <see cref="OfficeTextIndex"/>: terms, phrases, prefixes and AND/OR combinations of them.
/// Query text is analysed with the index analyser when the query runs.
/// </summary>
/// <remarks>
/// <see cref="Parse(string, OfficeTextQueryOperator)"/> accepts words, <c>"quoted phrases"</c>, <c>prefix*</c>,
/// the upper-case operators <c>AND</c> and <c>OR</c>, and parentheses. AND binds tighter than OR. A word that the
/// analyser splits into several tokens, such as <c>e-mail</c> or a run of Han ideographs, is matched as a phrase.
/// </remarks>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
sealed class OfficeTextQuery {
    private OfficeTextQuery(OfficeTextQueryKind kind, string? text, IReadOnlyList<OfficeTextQuery> children) {
        Kind = kind;
        Text = text;
        Children = children;
    }

    /// <summary>Node kind.</summary>
    public OfficeTextQueryKind Kind { get; }

    /// <summary>Source text for term, phrase and prefix nodes; null for AND and OR.</summary>
    public string? Text { get; }

    /// <summary>Child clauses of AND and OR nodes; empty for leaves.</summary>
    public IReadOnlyList<OfficeTextQuery> Children { get; }

    /// <summary>Number of term, phrase and prefix clauses in the query.</summary>
    public int ClauseCount => Children.Count == 0 ? 1 : Children.Sum(static child => child.ClauseCount);

    /// <summary>Matches one word.</summary>
    public static OfficeTextQuery Term(string text) => Leaf(OfficeTextQueryKind.Term, text);

    /// <summary>Matches words that appear consecutively in one field.</summary>
    public static OfficeTextQuery Phrase(string text) => Leaf(OfficeTextQueryKind.Phrase, text);

    /// <summary>Matches any term starting with <paramref name="text"/> after normalization and folding.</summary>
    public static OfficeTextQuery Prefix(string text) => Leaf(OfficeTextQueryKind.Prefix, text);

    /// <summary>Matches documents matching every clause.</summary>
    public static OfficeTextQuery And(params OfficeTextQuery[] clauses) => Combine(OfficeTextQueryKind.And, clauses);

    /// <summary>Matches documents matching at least one clause.</summary>
    public static OfficeTextQuery Or(params OfficeTextQuery[] clauses) => Combine(OfficeTextQueryKind.Or, clauses);

    /// <summary>
    /// Parses query syntax. Returns null when the text contains no clauses, for example only operators or quotes.
    /// </summary>
    public static OfficeTextQuery? Parse(string text, OfficeTextQueryOperator defaultOperator = OfficeTextQueryOperator.And) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        var parser = new Parser(Tokenize(text), defaultOperator);
        return parser.ParseAll();
    }

    /// <inheritdoc />
    public override string ToString() {
        switch (Kind) {
            case OfficeTextQueryKind.Term: return Text!;
            case OfficeTextQueryKind.Phrase: return "\"" + Text + "\"";
            case OfficeTextQueryKind.Prefix: return Text + "*";
            default:
                string separator = Kind == OfficeTextQueryKind.And ? " AND " : " OR ";
                return "(" + string.Join(separator, Children.Select(static child => child.ToString())) + ")";
        }
    }

    private static OfficeTextQuery Leaf(OfficeTextQueryKind kind, string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        return new OfficeTextQuery(kind, text, Array.Empty<OfficeTextQuery>());
    }

    private static OfficeTextQuery Combine(OfficeTextQueryKind kind, OfficeTextQuery[] clauses) {
        if (clauses == null) throw new ArgumentNullException(nameof(clauses));
        if (clauses.Length == 0) throw new ArgumentException("At least one clause is required.", nameof(clauses));
        if (clauses.Any(static clause => clause == null)) throw new ArgumentException("Clauses must not be null.", nameof(clauses));
        if (clauses.Length == 1) return clauses[0];
        var flattened = new List<OfficeTextQuery>();
        foreach (OfficeTextQuery clause in clauses) {
            if (clause.Kind == kind) flattened.AddRange(clause.Children);
            else flattened.Add(clause);
        }
        return new OfficeTextQuery(kind, null, flattened.AsReadOnly());
    }

    private enum SyntaxKind { Word, Prefix, Phrase, And, Or, Open, Close }

    private readonly struct Syntax {
        internal Syntax(SyntaxKind kind, string text) {
            Kind = kind;
            Text = text;
        }

        internal SyntaxKind Kind { get; }
        internal string Text { get; }
    }

    private static List<Syntax> Tokenize(string text) {
        var tokens = new List<Syntax>();
        int index = 0;
        while (index < text.Length) {
            char character = text[index];
            if (char.IsWhiteSpace(character)) {
                index++;
            } else if (character == '"') {
                int end = text.IndexOf('"', index + 1);
                if (end < 0) end = text.Length;
                tokens.Add(new Syntax(SyntaxKind.Phrase, text.Substring(index + 1, end - index - 1)));
                index = Math.Min(text.Length, end + 1);
            } else if (character == '(') {
                tokens.Add(new Syntax(SyntaxKind.Open, "("));
                index++;
            } else if (character == ')') {
                tokens.Add(new Syntax(SyntaxKind.Close, ")"));
                index++;
            } else {
                var builder = new StringBuilder();
                while (index < text.Length && !char.IsWhiteSpace(text[index]) &&
                       text[index] != '"' && text[index] != '(' && text[index] != ')') {
                    builder.Append(text[index++]);
                }
                string word = builder.ToString();
                if (word == "AND") tokens.Add(new Syntax(SyntaxKind.And, word));
                else if (word == "OR") tokens.Add(new Syntax(SyntaxKind.Or, word));
                else if (word.Length > 1 && word[word.Length - 1] == '*') tokens.Add(new Syntax(SyntaxKind.Prefix, word.TrimEnd('*')));
                else tokens.Add(new Syntax(SyntaxKind.Word, word));
            }
        }
        return tokens;
    }

    private sealed class Parser {
        private readonly List<Syntax> _tokens;
        private readonly OfficeTextQueryOperator _defaultOperator;
        private int _index;

        internal Parser(List<Syntax> tokens, OfficeTextQueryOperator defaultOperator) {
            _tokens = tokens;
            _defaultOperator = defaultOperator;
        }

        internal OfficeTextQuery? ParseAll() {
            var parts = new List<OfficeTextQuery>();
            while (_index < _tokens.Count) {
                OfficeTextQuery? part = ParseOr();
                if (part != null) parts.Add(part);
                if (_index < _tokens.Count && _tokens[_index].Kind == SyntaxKind.Close) _index++; // unbalanced ')'
            }
            return parts.Count == 0 ? null : Join(_defaultOperator == OfficeTextQueryOperator.And
                ? OfficeTextQueryKind.And : OfficeTextQueryKind.Or, parts);
        }

        private OfficeTextQuery? ParseOr() {
            var parts = new List<OfficeTextQuery>();
            while (true) {
                OfficeTextQuery? part = ParseAnd();
                if (part != null) parts.Add(part);
                if (Peek(SyntaxKind.Or)) {
                    _index++;
                    continue;
                }
                if (_defaultOperator == OfficeTextQueryOperator.Or && StartsClause()) continue;
                break;
            }
            return Join(OfficeTextQueryKind.Or, parts);
        }

        private OfficeTextQuery? ParseAnd() {
            var parts = new List<OfficeTextQuery>();
            while (true) {
                OfficeTextQuery? part = ParseUnit();
                if (part != null) parts.Add(part);
                if (Peek(SyntaxKind.And)) {
                    _index++;
                    continue;
                }
                if (_defaultOperator == OfficeTextQueryOperator.And && StartsClause()) continue;
                break;
            }
            return Join(OfficeTextQueryKind.And, parts);
        }

        private OfficeTextQuery? ParseUnit() {
            while (_index < _tokens.Count) {
                Syntax token = _tokens[_index];
                switch (token.Kind) {
                    case SyntaxKind.Word:
                        _index++;
                        return Term(token.Text);
                    case SyntaxKind.Prefix:
                        _index++;
                        return Prefix(token.Text);
                    case SyntaxKind.Phrase:
                        _index++;
                        return token.Text.Trim().Length == 0 ? null : Phrase(token.Text);
                    case SyntaxKind.Open:
                        _index++;
                        OfficeTextQuery? inner = ParseOr();
                        if (Peek(SyntaxKind.Close)) _index++;
                        return inner;
                    case SyntaxKind.And:
                    case SyntaxKind.Or:
                        _index++; // a leading or doubled operator has nothing to join
                        continue;
                    default:
                        return null;
                }
            }
            return null;
        }

        private bool Peek(SyntaxKind kind) => _index < _tokens.Count && _tokens[_index].Kind == kind;

        private bool StartsClause() =>
            _index < _tokens.Count && _tokens[_index].Kind is SyntaxKind.Word or SyntaxKind.Prefix or SyntaxKind.Phrase or SyntaxKind.Open;

        private static OfficeTextQuery? Join(OfficeTextQueryKind kind, List<OfficeTextQuery> parts) {
            if (parts.Count == 0) return null;
            return kind == OfficeTextQueryKind.And ? And(parts.ToArray()) : Or(parts.ToArray());
        }
    }
}
