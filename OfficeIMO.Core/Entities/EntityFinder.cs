using System;
using System.Collections.Generic;
using System.Threading;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Finds and scores validated entities in text with opt-in continuation across source segments.</summary>
public sealed class EntityFinder {
    private readonly EntityCatalog _catalog;
    private readonly EntityFinderOptions _options;
    private readonly HashSet<string> _enabled = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string>? _languages;
    private readonly List<IEntityRecognizer> _recognizers = new();
    private readonly EntityTokenShape _shapes;

    /// <summary>Constructs a reusable finder. A null catalog uses all built-in packs.</summary>
    public EntityFinder(EntityCatalog? catalog = null, EntityFinderOptions? options = null) {
        _catalog = catalog ?? EntityCatalog.Default;
        options ??= new EntityFinderOptions();
        _options = Snapshot(options);
        _languages = options.Languages == null ? null : new HashSet<string>(options.Languages, StringComparer.OrdinalIgnoreCase);

        foreach (EntityScheme scheme in _catalog.Schemes) {
            if (options.Kinds.Count > 0 && !options.Kinds.Contains(scheme.Kind) ||
                options.Countries.Count > 0 && scheme.Country != null && !options.Countries.Contains(scheme.Country) ||
                options.IncludeSchemes.Count > 0 && !options.IncludeSchemes.Contains(scheme.Id) ||
                options.ExcludeSchemes.Contains(scheme.Id)) continue;
            _enabled.Add(scheme.Id);
        }
        foreach (IEntityRecognizer recognizer in _catalog.Recognizers) {
            foreach (EntityScheme scheme in recognizer.Schemes) {
                if (!_enabled.Contains(scheme.Id)) continue;
                _recognizers.Add(recognizer);
                _shapes |= recognizer.Shapes;
                break;
            }
        }
    }

    /// <summary>Searches a single string as segment zero.</summary>
    public EntityFindResult Find(string text) => Find(text, CancellationToken.None);

    /// <summary>Searches a single string as segment zero with cancellation.</summary>
    public EntityFindResult Find(string text, CancellationToken cancellationToken) =>
        Find(new[] { new EntityTextSegment(text) }, cancellationToken);

    /// <summary>Searches the segments in order under the configured limits.</summary>
    public EntityFindResult Find(IEnumerable<EntityTextSegment> segments) => Find(segments, CancellationToken.None);

    /// <summary>Searches the segments in order with cancellation.</summary>
    public EntityFindResult Find(IEnumerable<EntityTextSegment> segments, CancellationToken cancellationToken) {
        if (segments == null) throw new ArgumentNullException(nameof(segments));
        cancellationToken.ThrowIfCancellationRequested();

        var matches = new List<EntityMatch>();
        if (_options.MaxMatches == 0 || _options.MaxInputCharacters == 0)
            return new EntityFindResult(matches, truncated: true, droppedOverlaps: 0);
        int remaining = _options.MaxInputCharacters;
        int segmentIndex = 0;
        int dropped = 0;
        using IEnumerator<EntityTextSegment> source = segments.GetEnumerator();
        bool hasSegment = source.MoveNext();
        while (hasSegment) {
            cancellationToken.ThrowIfCancellationRequested();
            var chain = new EntitySegmentChain(source.Current, segmentIndex);
            bool clipped = false;
            do {
                cancellationToken.ThrowIfCancellationRequested();
                EntityTextSegment segment = source.Current;
                int length = segment.Text?.Length ?? 0;
                int take = Math.Min(remaining, length);
                chain.Append(segment, segmentIndex, take);
                remaining -= take;
                if (remaining == 0) {
                    clipped = take < length || HasContinuedText(source, cancellationToken);
                    break;
                }

                hasSegment = source.MoveNext();
                segmentIndex++;
            } while (hasSegment && source.Current.ContinuesPrevious);

            if (chain.Length > 0) {
                List<EntityMatch> found = FindChain(chain, clipped, cancellationToken, out int segmentDropped);
                dropped += segmentDropped;
                int take = Math.Min(found.Count, _options.MaxMatches - matches.Count);
                for (int index = 0; index < take; index++) matches.Add(found[index]);
            }

            if (remaining == 0 || matches.Count == _options.MaxMatches)
                return new EntityFindResult(matches, truncated: true, dropped);
        }

        return new EntityFindResult(matches, truncated: false, dropped);
    }

    private static bool HasContinuedText(IEnumerator<EntityTextSegment> source, CancellationToken cancellationToken) {
        while (source.MoveNext() && source.Current.ContinuesPrevious) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.IsNullOrEmpty(source.Current.Text)) return true;
        }
        return false;
    }

    private List<EntityMatch> FindChain(EntitySegmentChain chain, bool clipped, CancellationToken cancellationToken, out int dropped) {
        string text = chain.GetText();
        List<EntityToken> tokens = EntityTokenizer.Tokenize(text, _shapes, cancellationToken);
        var candidates = new List<ScoredEntity>();
        var context = new EntityRecognitionContext(_catalog, _options, _enabled, _languages, chain,
            text, clipped, tokens, candidates, cancellationToken);
        foreach (IEntityRecognizer recognizer in _recognizers) {
            cancellationToken.ThrowIfCancellationRequested();
            recognizer.Recognize(context);
        }

        int originalCount = candidates.Count;
        for (int index = 0; index < originalCount; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            ScoredEntity reading = candidates[index];
            reading.Scheme.RecognizeContained(context, reading);
        }
        return EntityOverlapResolver.Resolve(candidates, _options, chain, cancellationToken, out dropped);
    }

    private static EntityFinderOptions Snapshot(EntityFinderOptions options) {
        if (double.IsNaN(options.MinimumConfidence) || options.MinimumConfidence < 0 || options.MinimumConfidence > 1)
            throw new ArgumentOutOfRangeException(nameof(options.MinimumConfidence));
        if (options.MaxInputCharacters < 0) throw new ArgumentOutOfRangeException(nameof(options.MaxInputCharacters));
        if (options.MaxMatches < 0) throw new ArgumentOutOfRangeException(nameof(options.MaxMatches));
        if (options.ContextWindow < 0) throw new ArgumentOutOfRangeException(nameof(options.ContextWindow));
        if (options.OverlapPolicy is not (EntityOverlapPolicy.BestOnly or EntityOverlapPolicy.All))
            throw new ArgumentOutOfRangeException(nameof(options.OverlapPolicy));
        return new EntityFinderOptions {
            MinimumConfidence = options.MinimumConfidence,
            MaxInputCharacters = options.MaxInputCharacters,
            MaxMatches = options.MaxMatches,
            ContextWindow = options.ContextWindow,
            OverlapPolicy = options.OverlapPolicy,
        };
    }
}
