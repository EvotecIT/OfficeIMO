#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.TextAnalysis;
#else
namespace OfficeIMO.TextAnalysis;
#endif

/// <summary>
/// In-memory inverted index with positional postings, ranked by BM25F: term, phrase, prefix and AND/OR queries with
/// optional per-field weights, highlights and snippets.
/// </summary>
/// <remarks>
/// <para>
/// For a single-term clause <c>t</c> and document <c>d</c> the score is
/// <c>idf(t) · tf' · (k1 + 1) / (tf' + k1)</c>, where
/// <c>tf' = Σ_fields weight(f) · tf(t, d, f) / (1 − b + b · length(d, f) / averageLength(f))</c> and
/// <c>idf(t) = ln(1 + (N − df(t) + 0.5) / (df(t) + 0.5))</c>. With one field of weight 1 this is classic BM25.
/// Field length is its token count; the average is over all documents. A phrase uses its occurrence count as
/// <c>tf</c> and the sum of its terms' idf. A prefix scores as the sum of the terms it expands to. AND and OR add the
/// scores of their matching clauses.
/// </para>
/// <para>
/// Hits are ordered by descending score and then by insertion order, so equal scores rank deterministically.
/// Adding documents is not thread-safe; concurrent searches are safe once no documents are being added.
/// </para>
/// </remarks>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
sealed class OfficeTextIndex {
    private readonly OfficeTextIndexOptions _options;
    private readonly OfficeTextAnalyzer _analyzer;
    private readonly Dictionary<string, List<Posting>> _postings = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _fieldSlots = new(StringComparer.Ordinal);
    private readonly List<string> _fieldNames = new();
    private readonly List<double> _fieldWeights = new();
    private readonly List<long> _fieldLengthTotals = new();
    private readonly List<DocumentEntry> _documents = new();
    private readonly object _termOrderLock = new();
    private string[]? _sortedTerms;
    private long _tokenCount;

    /// <summary>Creates an empty index. The options are copied.</summary>
    public OfficeTextIndex(OfficeTextIndexOptions? options = null) {
        OfficeTextIndexOptions source = options ?? new OfficeTextIndexOptions();
        if (double.IsNaN(source.K1) || double.IsInfinity(source.K1) || source.K1 < 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "K1 must be a finite value of zero or more.");
        }
        if (double.IsNaN(source.B) || source.B < 0 || source.B > 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "B must be between 0 and 1.");
        }
        if (source.MaximumDocuments < 1 || source.MaximumTokens < 1 || source.MaximumFieldCharacters < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "Index limits must be positive.");
        }
        foreach (KeyValuePair<string, double> weight in source.FieldWeights) {
            if (double.IsNaN(weight.Value) || double.IsInfinity(weight.Value) || weight.Value < 0) {
                throw new ArgumentOutOfRangeException(nameof(options), "Field weights must be finite values of zero or more.");
            }
        }
        _options = new OfficeTextIndexOptions {
            K1 = source.K1,
            B = source.B,
            Analyzer = (source.Analyzer ?? new OfficeTextAnalyzerOptions()).Clone(),
            MaximumDocuments = source.MaximumDocuments,
            MaximumTokens = source.MaximumTokens,
            MaximumFieldCharacters = source.MaximumFieldCharacters
        };
        foreach (KeyValuePair<string, double> weight in source.FieldWeights) _options.FieldWeights[weight.Key] = weight.Value;
        _analyzer = new OfficeTextAnalyzer(_options.Analyzer);
    }

    /// <summary>Analyser used for indexed and query text.</summary>
    public OfficeTextAnalyzer Analyzer => _analyzer;

    /// <summary>Number of indexed documents.</summary>
    public int DocumentCount => _documents.Count;

    /// <summary>Number of distinct terms.</summary>
    public int TermCount => _postings.Count;

    /// <summary>Number of tokens held by the index.</summary>
    public long TokenCount => _tokenCount;

    /// <summary>True when a document or token limit truncated or rejected input.</summary>
    public bool LimitReached { get; private set; }

    /// <summary>Indexes one document with one field named <c>body</c>.</summary>
    public bool Add(string id, string text) => Add(id, new OfficeTextField("body", text));

    /// <summary>
    /// Indexes one document. Returns false, and sets <see cref="LimitReached"/>, when the document was rejected because
    /// a limit was already reached. A document that reaches the token limit is indexed up to it and returns true.
    /// </summary>
    public bool Add(string id, params OfficeTextField[] fields) {
        if (id == null) throw new ArgumentNullException(nameof(id));
        if (fields == null) throw new ArgumentNullException(nameof(fields));
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (OfficeTextField field in fields) {
            if (field.Name == null) throw new ArgumentException("Every field needs a name.", nameof(fields));
            if (!names.Add(field.Name)) throw new ArgumentException("Field '" + field.Name + "' is listed twice.", nameof(fields));
        }
        if (_documents.Count >= _options.MaximumDocuments || _tokenCount >= _options.MaximumTokens) {
            LimitReached = true;
            return false;
        }

        int documentIndex = _documents.Count;
        var entry = new DocumentEntry(id, fields.Length);
        for (int fieldIndex = 0; fieldIndex < fields.Length; fieldIndex++) {
            OfficeTextField field = fields[fieldIndex];
            int slot = GetFieldSlot(field.Name);
            string text = field.Text;
            if (text.Length > _options.MaximumFieldCharacters) {
                int cut = _options.MaximumFieldCharacters;
                if (char.IsLowSurrogate(text[cut]) && char.IsHighSurrogate(text[cut - 1])) cut--;
                text = text.Substring(0, cut);
                LimitReached = true;
            }
            IReadOnlyList<OfficeTextToken> tokens = _analyzer.Analyze(text);
            int count = tokens.Count;
            long available = _options.MaximumTokens - _tokenCount;
            if (count > available) {
                count = (int)available;
                LimitReached = true;
            }
            entry.Slots[fieldIndex] = slot;
            entry.Texts[fieldIndex] = text;
            entry.Lengths[fieldIndex] = count;
            _fieldLengthTotals[slot] += count;
            _tokenCount += count;

            var positions = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            var order = new List<string>();
            for (int tokenIndex = 0; tokenIndex < count; tokenIndex++) {
                string term = tokens[tokenIndex].Term;
                if (!positions.TryGetValue(term, out List<int>? list)) {
                    list = new List<int>();
                    positions.Add(term, list);
                    order.Add(term);
                }
                list.Add(tokens[tokenIndex].Position);
            }
            foreach (string term in order) {
                if (!_postings.TryGetValue(term, out List<Posting>? postings)) {
                    postings = new List<Posting>();
                    _postings.Add(term, postings);
                    _sortedTerms = null;
                }
                postings.Add(new Posting(documentIndex, slot, positions[term].ToArray()));
            }
        }
        _documents.Add(entry);
        return true;
    }

    /// <summary>
    /// Parses <paramref name="query"/> with <see cref="OfficeTextQuery.Parse"/>, using
    /// <see cref="OfficeTextIndexSearchOptions.DefaultOperator"/>, and searches.
    /// </summary>
    public OfficeTextIndexSearchResult Search(string query, OfficeTextIndexSearchOptions options) {
        if (query == null) throw new ArgumentNullException(nameof(query));
        OfficeTextIndexSearchOptions effective = options ?? throw new ArgumentNullException(nameof(options));
        OfficeTextQuery? parsed = OfficeTextQuery.Parse(query, effective.DefaultOperator);
        if (parsed == null) {
            return new OfficeTextIndexSearchResult(null, Array.Empty<OfficeTextIndexHit>(), 0, false, false, false);
        }
        return Search(parsed, effective);
    }

    /// <summary>Parses and searches <paramref name="query"/> with default search options.</summary>
    public OfficeTextIndexSearchResult Search(string query) => Search(query, new OfficeTextIndexSearchOptions());

    /// <summary>Evaluates <paramref name="query"/> with default search options.</summary>
    public OfficeTextIndexSearchResult Search(OfficeTextQuery query) => Search(query, new OfficeTextIndexSearchOptions());

    /// <summary>Evaluates <paramref name="query"/> and returns ranked hits.</summary>
    public OfficeTextIndexSearchResult Search(OfficeTextQuery query, OfficeTextIndexSearchOptions options) {
        if (query == null) throw new ArgumentNullException(nameof(query));
        OfficeTextIndexSearchOptions effective = options ?? throw new ArgumentNullException(nameof(options));
        if (effective.MaximumResults < 1 || effective.MaximumClauses < 1 || effective.MaximumPrefixExpansions < 1 ||
            effective.MaximumPostingsVisited < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "Search limits must be positive.");
        }
        if (query.ClauseCount > effective.MaximumClauses) {
            throw new ArgumentException(
                "The query has " + query.ClauseCount + " clauses; the limit is " + effective.MaximumClauses + ".",
                nameof(query));
        }

        var evaluation = new Evaluation(this, effective);
        Node? root = evaluation.Compile(query);
        Dictionary<int, double> scores = root == null ? new Dictionary<int, double>() : evaluation.Evaluate(root);
        var ranked = scores.ToList();
        ranked.Sort(static (left, right) => {
            int byScore = right.Value.CompareTo(left.Value);
            return byScore != 0 ? byScore : left.Key.CompareTo(right.Key);
        });
        int returned = Math.Min(ranked.Count, effective.MaximumResults);
        var hits = new OfficeTextIndexHit[returned];
        for (int index = 0; index < returned; index++) {
            int document = ranked[index].Key;
            IReadOnlyList<OfficeTextHighlight> highlights = effective.IncludeHighlights && root != null
                ? CollectHighlights(document, root)
                : Array.Empty<OfficeTextHighlight>();
            hits[index] = new OfficeTextIndexHit(_documents[document].Id, document, ranked[index].Value, highlights);
        }
        return new OfficeTextIndexSearchResult(query, hits, ranked.Count, evaluation.PrefixExpansionLimited,
            evaluation.WorkLimitReached, root != null);
    }

    /// <summary>
    /// Returns up to <paramref name="maximumCharacters"/> of <paramref name="field"/> from the hit's document, centred
    /// on its first highlight in that field, or the field start when it has none. Returns null when the document has
    /// no such field.
    /// </summary>
    public OfficeTextSnippet? CreateSnippet(OfficeTextIndexHit hit, string field, int maximumCharacters = 280) {
        if (hit == null) throw new ArgumentNullException(nameof(hit));
        if (field == null) throw new ArgumentNullException(nameof(field));
        if (maximumCharacters < 1) throw new ArgumentOutOfRangeException(nameof(maximumCharacters));
        if ((uint)hit.DocumentIndex >= (uint)_documents.Count) throw new ArgumentException("The hit is not from this index.", nameof(hit));
        DocumentEntry entry = _documents[hit.DocumentIndex];
        int fieldIndex = Array.FindIndex(entry.Slots, slot => string.Equals(_fieldNames[slot], field, StringComparison.Ordinal));
        if (fieldIndex < 0) return null;
        string text = entry.Texts[fieldIndex];
        OfficeTextHighlight[] inField = hit.Highlights.Where(h => string.Equals(h.Field, field, StringComparison.Ordinal)).ToArray();

        int anchorStart = inField.Length > 0 ? inField[0].Start : 0;
        int anchorLength = inField.Length > 0 ? inField[0].Length : 0;
        int context = Math.Max(0, (maximumCharacters - anchorLength) / 2);
        int start = Math.Max(0, anchorStart - context);
        int end = Math.Min(text.Length, start + maximumCharacters);
        start = Math.Max(0, Math.Min(start, end - maximumCharacters));
        SnapToGraphemeClusters(text, ref start, ref end, anchorStart, anchorStart + anchorLength);

        var relative = new List<OfficeTextHighlight>();
        foreach (OfficeTextHighlight highlight in inField) {
            int highlightStart = Math.Max(start, highlight.Start);
            int highlightEnd = Math.Min(end, highlight.Start + highlight.Length);
            if (highlightEnd > highlightStart) relative.Add(new OfficeTextHighlight(field, highlightStart - start, highlightEnd - highlightStart));
        }
        return new OfficeTextSnippet(text.Substring(start, end - start), start, start > 0, end < text.Length, relative.AsReadOnly());
    }

    private static void SnapToGraphemeClusters(string text, ref int start, ref int end, int anchorStart, int anchorEnd) {
        if (start == 0 && end == text.Length) return;
        const int Margin = 64;
        int windowStart = Math.Max(0, start - Margin);
        int windowEnd = Math.Min(text.Length, end + Margin);
        if (windowStart > 0 && char.IsLowSurrogate(text[windowStart])) windowStart--;
        if (windowEnd < text.Length && char.IsLowSurrogate(text[windowEnd])) windowEnd++;
        IReadOnlyList<int> boundaries = UnicodeTextSegmenter.GetGraphemeClusterBoundaries(text.Substring(windowStart, windowEnd - windowStart));
        int snappedStart = start;
        int snappedEnd = end;
        if (start > 0) {
            int candidate = int.MaxValue;
            foreach (int offset in boundaries) {
                if (offset + windowStart >= start) {
                    candidate = offset + windowStart;
                    break;
                }
            }
            snappedStart = Math.Min(candidate == int.MaxValue ? start : candidate, anchorStart);
        }
        if (end < text.Length) {
            int candidate = -1;
            foreach (int offset in boundaries) {
                if (offset + windowStart <= end) candidate = offset + windowStart;
                else break;
            }
            snappedEnd = Math.Max(candidate < 0 ? end : candidate, Math.Min(anchorEnd, text.Length));
        }
        if (snappedEnd > snappedStart) {
            start = snappedStart;
            end = snappedEnd;
        }
    }

    private int GetFieldSlot(string name) {
        if (_fieldSlots.TryGetValue(name, out int slot)) return slot;
        slot = _fieldNames.Count;
        _fieldSlots.Add(name, slot);
        _fieldNames.Add(name);
        _fieldWeights.Add(_options.FieldWeights.TryGetValue(name, out double weight) ? weight : 1d);
        _fieldLengthTotals.Add(0);
        return slot;
    }

    private string[] GetSortedTerms() {
        lock (_termOrderLock) {
            if (_sortedTerms == null) {
                string[] terms = _postings.Keys.ToArray();
                Array.Sort(terms, StringComparer.Ordinal);
                _sortedTerms = terms;
            }
            return _sortedTerms;
        }
    }

    private double InverseDocumentFrequency(int documentFrequency) {
        double count = _documents.Count;
        return Math.Log(1d + (count - documentFrequency + 0.5d) / (documentFrequency + 0.5d));
    }

    private static int CountDocuments(List<Posting> postings) {
        int count = 0;
        int last = -1;
        foreach (Posting posting in postings) {
            if (posting.Document != last) {
                count++;
                last = posting.Document;
            }
        }
        return count;
    }

    private double FieldNorm(int document, int slot) {
        DocumentEntry entry = _documents[document];
        int length = 0;
        for (int index = 0; index < entry.Slots.Length; index++) {
            if (entry.Slots[index] == slot) {
                length = entry.Lengths[index];
                break;
            }
        }
        double average = (double)_fieldLengthTotals[slot] / _documents.Count;
        if (average <= 0) return 1d;
        return 1d - _options.B + _options.B * length / average;
    }

    private double Saturate(double weightedFrequency, double idf) =>
        weightedFrequency <= 0 ? 0 : idf * weightedFrequency * (_options.K1 + 1d) / (weightedFrequency + _options.K1);

    private IReadOnlyList<OfficeTextHighlight> CollectHighlights(int document, Node root) {
        var terms = new HashSet<string>(StringComparer.Ordinal);
        var phrases = new List<string[]>();
        root.CollectHighlightTargets(terms, phrases);
        DocumentEntry entry = _documents[document];
        var highlights = new List<OfficeTextHighlight>();
        for (int fieldIndex = 0; fieldIndex < entry.Slots.Length; fieldIndex++) {
            if (entry.Lengths[fieldIndex] == 0) continue;
            IReadOnlyList<OfficeTextToken> tokens = _analyzer.Analyze(entry.Texts[fieldIndex]);
            int count = Math.Min(tokens.Count, entry.Lengths[fieldIndex]);
            var ranges = new List<(int Start, int End)>();
            for (int tokenIndex = 0; tokenIndex < count; tokenIndex++) {
                OfficeTextToken token = tokens[tokenIndex];
                if (terms.Contains(token.Term)) ranges.Add((token.Start, token.Start + token.Length));
                foreach (string[] phrase in phrases) {
                    if (tokenIndex + phrase.Length > count) continue;
                    bool match = true;
                    for (int offset = 0; offset < phrase.Length && match; offset++) {
                        match = string.Equals(tokens[tokenIndex + offset].Term, phrase[offset], StringComparison.Ordinal);
                    }
                    if (match) {
                        OfficeTextToken last = tokens[tokenIndex + phrase.Length - 1];
                        ranges.Add((token.Start, last.Start + last.Length));
                    }
                }
            }
            ranges.Sort();
            string name = _fieldNames[entry.Slots[fieldIndex]];
            int mergedStart = -1;
            int mergedEnd = -1;
            foreach ((int start, int end) in ranges) {
                if (mergedStart >= 0 && start <= mergedEnd) {
                    mergedEnd = Math.Max(mergedEnd, end);
                    continue;
                }
                if (mergedStart >= 0) highlights.Add(new OfficeTextHighlight(name, mergedStart, mergedEnd - mergedStart));
                mergedStart = start;
                mergedEnd = end;
            }
            if (mergedStart >= 0) highlights.Add(new OfficeTextHighlight(name, mergedStart, mergedEnd - mergedStart));
        }
        return highlights.AsReadOnly();
    }

    private sealed class Posting {
        internal Posting(int document, int slot, int[] positions) {
            Document = document;
            Slot = slot;
            Positions = positions;
        }

        internal int Document { get; }
        internal int Slot { get; }
        internal int[] Positions { get; }
    }

    private sealed class DocumentEntry {
        internal DocumentEntry(string id, int fieldCount) {
            Id = id;
            Slots = new int[fieldCount];
            Texts = new string[fieldCount];
            Lengths = new int[fieldCount];
        }

        internal string Id { get; }
        internal int[] Slots { get; }
        internal string[] Texts { get; }
        internal int[] Lengths { get; }
    }

    private abstract class Node {
        internal abstract void CollectHighlightTargets(HashSet<string> terms, List<string[]> phrases);
    }

    private sealed class TermNode : Node {
        internal TermNode(IReadOnlyList<string> terms) => Terms = terms;

        // One term, or the expansion of a prefix.
        internal IReadOnlyList<string> Terms { get; }

        internal override void CollectHighlightTargets(HashSet<string> terms, List<string[]> phrases) {
            foreach (string term in Terms) terms.Add(term);
        }
    }

    private sealed class PhraseNode : Node {
        internal PhraseNode(string[] terms) => Terms = terms;

        internal string[] Terms { get; }

        internal override void CollectHighlightTargets(HashSet<string> terms, List<string[]> phrases) => phrases.Add(Terms);
    }

    private sealed class BooleanNode : Node {
        internal BooleanNode(bool conjunction, IReadOnlyList<Node> children) {
            Conjunction = conjunction;
            Children = children;
        }

        internal bool Conjunction { get; }
        internal IReadOnlyList<Node> Children { get; }

        internal override void CollectHighlightTargets(HashSet<string> terms, List<string[]> phrases) {
            foreach (Node child in Children) child.CollectHighlightTargets(terms, phrases);
        }
    }

    private sealed class Evaluation {
        private readonly OfficeTextIndex _index;
        private readonly OfficeTextIndexSearchOptions _options;
        private long _visited;

        internal Evaluation(OfficeTextIndex index, OfficeTextIndexSearchOptions options) {
            _index = index;
            _options = options;
        }

        internal bool PrefixExpansionLimited { get; private set; }

        internal bool WorkLimitReached { get; private set; }

        // Clauses with no analysable text are dropped, so "the - word" searches for "the" and "word".
        internal Node? Compile(OfficeTextQuery query) {
            switch (query.Kind) {
                case OfficeTextQueryKind.Term:
                case OfficeTextQueryKind.Phrase: {
                    string[] terms = _index._analyzer.Analyze(query.Text!).Select(static token => token.Term).ToArray();
                    if (terms.Length == 0) return null;
                    return terms.Length == 1 ? new TermNode(terms) : new PhraseNode(terms);
                }
                case OfficeTextQueryKind.Prefix: {
                    string[] words = _index._analyzer.Analyze(query.Text!).Select(static token => token.Term).ToArray();
                    if (words.Length == 0) return null;
                    Node prefix = new TermNode(Expand(words[words.Length - 1]));
                    if (words.Length == 1) return prefix;
                    var parts = words.Take(words.Length - 1).Select(static term => (Node)new TermNode(new[] { term })).ToList();
                    parts.Add(prefix);
                    return new BooleanNode(true, parts);
                }
                default: {
                    var children = query.Children.Select(Compile).Where(static child => child != null).Select(static child => child!).ToList();
                    if (children.Count == 0) return null;
                    return children.Count == 1 ? children[0] : new BooleanNode(query.Kind == OfficeTextQueryKind.And, children);
                }
            }
        }

        private IReadOnlyList<string> Expand(string prefix) {
            string[] sorted = _index.GetSortedTerms();
            int low = 0;
            int high = sorted.Length;
            while (low < high) {
                int middle = (low + high) >> 1;
                if (string.CompareOrdinal(sorted[middle], prefix) < 0) low = middle + 1;
                else high = middle;
            }
            var terms = new List<string>();
            for (int index = low; index < sorted.Length && sorted[index].StartsWith(prefix, StringComparison.Ordinal); index++) {
                if (terms.Count == _options.MaximumPrefixExpansions) {
                    PrefixExpansionLimited = true;
                    break;
                }
                terms.Add(sorted[index]);
            }
            return terms;
        }

        internal Dictionary<int, double> Evaluate(Node node) {
            switch (node) {
                case TermNode term: return EvaluateTerms(term.Terms);
                case PhraseNode phrase: return EvaluatePhrase(phrase.Terms);
                default: return EvaluateBoolean((BooleanNode)node);
            }
        }

        private Dictionary<int, double> EvaluateTerms(IReadOnlyList<string> terms) {
            var scores = new Dictionary<int, double>();
            foreach (string term in terms) {
                if (!_index._postings.TryGetValue(term, out List<Posting>? postings)) continue;
                double idf = _index.InverseDocumentFrequency(CountDocuments(postings));
                int document = -1;
                double weighted = 0;
                foreach (Posting posting in postings) {
                    if (!Visit()) break;
                    if (posting.Document != document) {
                        if (document >= 0) AddScore(scores, document, _index.Saturate(weighted, idf));
                        document = posting.Document;
                        weighted = 0;
                    }
                    weighted += _index._fieldWeights[posting.Slot] * posting.Positions.Length / _index.FieldNorm(posting.Document, posting.Slot);
                }
                if (document >= 0) AddScore(scores, document, _index.Saturate(weighted, idf));
            }
            return scores;
        }

        private Dictionary<int, double> EvaluatePhrase(string[] terms) {
            var scores = new Dictionary<int, double>();
            var lists = new List<Posting>[terms.Length];
            double idf = 0;
            for (int index = 0; index < terms.Length; index++) {
                if (!_index._postings.TryGetValue(terms[index], out List<Posting>? postings)) return scores;
                lists[index] = postings;
                idf += _index.InverseDocumentFrequency(CountDocuments(postings));
            }
            int slotCount = _index._fieldNames.Count;
            var byLocation = new Dictionary<long, Posting>[terms.Length];
            for (int index = 1; index < terms.Length; index++) {
                byLocation[index] = new Dictionary<long, Posting>();
                foreach (Posting posting in lists[index]) {
                    if (!Visit()) return scores;
                    byLocation[index][(long)posting.Document * slotCount + posting.Slot] = posting;
                }
            }

            int document = -1;
            double weighted = 0;
            foreach (Posting first in lists[0]) {
                if (!Visit()) break;
                long key = (long)first.Document * slotCount + first.Slot;
                var others = new Posting[terms.Length];
                bool present = true;
                for (int index = 1; index < terms.Length && present; index++) {
                    present = byLocation[index].TryGetValue(key, out Posting? other);
                    others[index] = other!;
                }
                if (!present) continue;
                int frequency = 0;
                foreach (int position in first.Positions) {
                    bool match = true;
                    for (int index = 1; index < terms.Length && match; index++) {
                        match = Array.BinarySearch(others[index].Positions, position + index) >= 0;
                    }
                    if (match) frequency++;
                }
                if (frequency == 0) continue;
                if (first.Document != document) {
                    if (document >= 0) AddScore(scores, document, _index.Saturate(weighted, idf));
                    document = first.Document;
                    weighted = 0;
                }
                weighted += _index._fieldWeights[first.Slot] * frequency / _index.FieldNorm(first.Document, first.Slot);
            }
            if (document >= 0) AddScore(scores, document, _index.Saturate(weighted, idf));
            return scores;
        }

        private Dictionary<int, double> EvaluateBoolean(BooleanNode node) {
            var results = node.Children.Select(Evaluate).ToList();
            var combined = new Dictionary<int, double>();
            if (node.Conjunction) {
                Dictionary<int, double> smallest = results.OrderBy(static result => result.Count).First();
                foreach (int document in smallest.Keys.OrderBy(static document => document)) {
                    double score = 0;
                    bool all = true;
                    foreach (Dictionary<int, double> result in results) {
                        if (!result.TryGetValue(document, out double value)) {
                            all = false;
                            break;
                        }
                        score += value;
                    }
                    if (all) combined[document] = score;
                }
            } else {
                var documents = new SortedSet<int>();
                foreach (Dictionary<int, double> result in results) documents.UnionWith(result.Keys);
                foreach (int document in documents) {
                    double score = 0;
                    foreach (Dictionary<int, double> result in results) {
                        if (result.TryGetValue(document, out double value)) score += value;
                    }
                    combined[document] = score;
                }
            }
            return combined;
        }

        private bool Visit() {
            if (WorkLimitReached) return false;
            if (++_visited > _options.MaximumPostingsVisited) {
                WorkLimitReached = true;
                return false;
            }
            return true;
        }

        private static void AddScore(Dictionary<int, double> scores, int document, double score) {
            scores.TryGetValue(document, out double existing);
            scores[document] = existing + score;
        }
    }
}
