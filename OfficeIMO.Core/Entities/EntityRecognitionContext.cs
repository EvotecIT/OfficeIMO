using System;
using System.Collections.Generic;
using System.Threading;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Access to one segment, its shared tokens and the candidate sink.</summary>
public sealed class EntityRecognitionContext {
    private readonly EntityCatalog _catalog;
    private readonly EntityFinderOptions _options;
    private readonly HashSet<string>? _languages;
    private readonly HashSet<string> _enabled;
    private readonly List<EntityToken> _tokens;
    private readonly string? _label;
    private readonly List<ScoredEntity> _candidates;
    private readonly bool _clipped;

    internal EntityRecognitionContext(EntityCatalog catalog, EntityFinderOptions options, HashSet<string> enabled,
        HashSet<string>? languages, EntityTextSegment segment, int segmentIndex, string text, bool clipped,
        List<EntityToken> tokens, List<ScoredEntity> candidates, CancellationToken cancellationToken) {
        _catalog = catalog;
        _options = options;
        _enabled = enabled;
        _languages = languages;
        _tokens = tokens;
        _candidates = candidates;
        _clipped = clipped;
        _label = segment.Label == null ? null : EntityContextMatcher.Fold(segment.Label);
        Segment = segment;
        SegmentIndex = segmentIndex;
        Text = text;
        CancellationToken = cancellationToken;
    }

    /// <summary>The searched portion of the segment.</summary>
    public string Text { get; }
    /// <summary>Segment metadata.</summary>
    public EntityTextSegment Segment { get; }
    /// <summary>Zero-based segment index in the current find operation.</summary>
    public int SegmentIndex { get; }
    /// <summary>Cancellation token provided by the caller.</summary>
    public CancellationToken CancellationToken { get; }

    /// <summary>Enumerates shared tokens that have any of the requested shapes.</summary>
    public IEnumerable<EntityToken> GetTokens(EntityTokenShape shapes) {
        for (int index = 0; index < _tokens.Count; index++) {
            if ((index & 4095) == 0) CancellationToken.ThrowIfCancellationRequested();
            EntityToken token = _tokens[index];
            if ((token.Shape & shapes) != 0) yield return token;
        }
    }

    /// <summary>Reports whether the active finder options permit this scheme.</summary>
    public bool IsEnabled(EntityScheme scheme) => _enabled.Contains(scheme.Id);

    /// <summary>Submits a valid reading within the searched segment.</summary>
    public void Add(EntityCandidate candidate) {
        if (candidate == null) throw new ArgumentNullException(nameof(candidate));
        if (candidate.Start > Text.Length - candidate.Length)
            throw new ArgumentException("The candidate extends outside the searched text.", nameof(candidate));
        if (!candidate.Validation.IsValid || _clipped && candidate.Start + candidate.Length == Text.Length) return;

        EntityScheme scheme = candidate.Validation.Scheme;
        if (!IsEnabled(scheme) || !_catalog.TryGetScheme(scheme.Id, out EntityScheme? registered) || !ReferenceEquals(scheme, registered)) return;

        FindContext(scheme, candidate.Start, out bool labeled, out EntityKeyword? keyword);
        bool structured = Segment.StructuredSource && candidate.Start == 0 && candidate.Length == Text.Length && !_clipped;
        if (candidate.RequiresContext && !labeled && keyword == null && !structured) return;

        EntityEvidence evidence = candidate.Evidence & ~(EntityEvidence.Label | EntityEvidence.ContextKeyword |
            EntityEvidence.StructuredSource | EntityEvidence.AmbiguousReading | EntityEvidence.CheckDigitPassed |
            EntityEvidence.NoCheckDefined | EntityEvidence.CommunityDocumentedCheck);
        evidence |= candidate.Validation.CheckDigit == EntityCheckOutcome.Passed ? EntityEvidence.CheckDigitPassed : EntityEvidence.NoCheckDefined;
        if (candidate.Validation.CheckDigit == EntityCheckOutcome.Passed &&
            candidate.Validation.CheckProvenance == EntityCheckProvenance.CommunityDocumented)
            evidence |= EntityEvidence.CommunityDocumentedCheck;
        if (labeled) evidence |= EntityEvidence.Label;
        if (keyword != null) evidence |= EntityEvidence.ContextKeyword;
        if (structured) evidence |= EntityEvidence.StructuredSource;

        _candidates.Add(new ScoredEntity(candidate, evidence, keyword, _catalog.RegistrationOrder[scheme.Id]));
    }

    internal bool HasRecognizerFor(string schemeId) {
        if (!_enabled.Contains(schemeId)) return false;
        foreach (IEntityRecognizer recognizer in _catalog.Recognizers) {
            foreach (EntityScheme scheme in recognizer.Schemes) {
                if (string.Equals(scheme.Id, schemeId, StringComparison.OrdinalIgnoreCase)) return true;
            }
        }
        return false;
    }
    internal bool RejectDigitRun(EntityToken token) => EntityDigitGuards.Reject(Text, token, _clipped);

    internal bool HasContext(EntityScheme scheme, int start) {
        FindContext(scheme, start, out bool labeled, out EntityKeyword? keyword);
        return labeled || keyword != null || Segment.StructuredSource;
    }

    private void FindContext(EntityScheme scheme, int start, out bool labeled, out EntityKeyword? keyword) {
        keyword = null;
        _catalog.Keywords.TryGetValue(scheme.Id, out List<EntityKeyword>? words);
        labeled = _label != null && (EntityContextMatcher.Find(words, _languages, _label) != null ||
            _label == EntityContextMatcher.Fold(scheme.Id) || _label == EntityContextMatcher.Fold(scheme.DisplayName));

        if (words == null || start == 0 || _options.ContextWindow == 0) return;
        int from = Math.Max(0, start - _options.ContextWindow);
        if (from > 0 && EntityContextMatcher.IsWord(Text[from - 1])) {
            while (from < start && EntityContextMatcher.IsWord(Text[from])) from++;
        }
        keyword = EntityContextMatcher.Find(words, _languages, EntityContextMatcher.Fold(Text.Substring(from, start - from)));
    }
}
