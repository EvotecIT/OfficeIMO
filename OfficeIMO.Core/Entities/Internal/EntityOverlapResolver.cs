using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Entities.Internal;

internal static class EntityOverlapResolver {
    internal static List<EntityMatch> Resolve(List<ScoredEntity> candidates, EntityFinderOptions options, EntitySegmentChain chain, CancellationToken cancellationToken, out int dropped) {
        dropped = 0;
        var matches = new List<EntityMatch>();
        if (candidates.Count == 0) return matches;

        candidates.Sort(CompareByPosition);
        Deduplicate(candidates);
        var alternatives = new Dictionary<ScoredEntity, List<ScoredEntity>>();
        var distinct = RankSameSpans(candidates, options.OverlapPolicy, alternatives, cancellationToken);

        if (options.OverlapPolicy == EntityOverlapPolicy.All) {
            foreach (ScoredEntity reading in distinct) {
                if (reading.Confidence >= options.MinimumConfidence) matches.Add(new EntityMatch(reading, chain));
            }
            return matches;
        }

        var children = new Dictionary<ScoredEntity, List<ScoredEntity>>();
        HashSet<ScoredEntity> contained = FindContained(distinct, options.MinimumConfidence, children, cancellationToken);
        var pending = new SortedSet<ScoredEntity>(Comparer<ScoredEntity>.Create(CompareByPriority));
        foreach (ScoredEntity reading in distinct) {
            if (reading.Confidence >= options.MinimumConfidence && !contained.Contains(reading)) pending.Add(reading);
        }

        var orderedStarts = new SortedSet<int>();
        var winners = new Dictionary<int, ScoredEntity>();
        var maskingFallbacks = new Dictionary<ScoredEntity, List<ScoredEntity>>();
        foreach (KeyValuePair<ScoredEntity, List<ScoredEntity>> group in alternatives) {
            foreach (ScoredEntity reading in group.Value) {
                if (reading.Confidence >= options.MinimumConfidence && reading.Scheme.Sensitivity != group.Key.Scheme.Sensitivity)
                    AddReading(maskingFallbacks, group.Key, reading);
            }
        }

        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            ScoredEntity reading = pending.Min!;
            pending.Remove(reading);
            ScoredEntity? winner = FindOverlappingWinner(reading, orderedStarts, winners);
            if (winner == null) {
                orderedStarts.Add(reading.Start);
                winners.Add(reading.Start, reading);
                continue;
            }

            dropped++;
            if (maskingFallbacks.TryGetValue(reading, out List<ScoredEntity>? inheritedFallbacks)) {
                foreach (ScoredEntity fallback in inheritedFallbacks) {
                    if (NeedsMaskingFallback(winner, fallback)) AddReading(maskingFallbacks, winner, fallback);
                }
            }
            if (NeedsMaskingFallback(winner, reading)) AddReading(maskingFallbacks, winner, reading);
            if (children.TryGetValue(reading, out List<ScoredEntity>? nested)) {
                foreach (ScoredEntity child in nested) {
                    if (child.Confidence >= options.MinimumConfidence) pending.Add(child);
                }
            }
        }
        var built = new Dictionary<ScoredEntity, EntityMatch>();
        // Contained readings follow their parents in source order, so materialize from the end.
        for (int index = distinct.Count - 1; index >= 0; index--) {
            ScoredEntity reading = distinct[index];
            if (!contained.Contains(reading) &&
                (!winners.TryGetValue(reading.Start, out ScoredEntity? winner) || !ReferenceEquals(reading, winner))) continue;
            built.Add(reading, ToMatch(reading, chain, alternatives, children, maskingFallbacks, built));
        }
        foreach (int start in orderedStarts) matches.Add(built[winners[start]]);
        return matches;
    }

    private static bool NeedsMaskingFallback(ScoredEntity winner, ScoredEntity reading) =>
        reading.Scheme.Sensitivity != winner.Scheme.Sensitivity || reading.Start < winner.Start || reading.End > winner.End;

    private static ScoredEntity? FindOverlappingWinner(ScoredEntity reading, SortedSet<int> starts, Dictionary<int, ScoredEntity> winners) {
        if (starts.Count == 0) return null;
        int previousStart = starts.GetViewBetween(-1, reading.Start).Max;
        if (winners.TryGetValue(previousStart, out ScoredEntity? predecessor) &&
            previousStart <= reading.Start && predecessor.End > reading.Start) return predecessor;

        int followingStart = starts.GetViewBetween(reading.Start + 1, reading.End).Min;
        return followingStart > reading.Start && followingStart < reading.End ? winners[followingStart] : null;
    }

    private static void Deduplicate(List<ScoredEntity> readings) {
        int count = 1;
        for (int index = 1; index < readings.Count; index++) {
            ScoredEntity previous = readings[count - 1];
            ScoredEntity current = readings[index];
            if (previous.Start != current.Start || previous.End != current.End || previous.RegistrationOrder != current.RegistrationOrder) {
                readings[count++] = current;
            } else if (current.RawConfidence > previous.RawConfidence) {
                readings[count - 1] = current;
            }
        }
        if (count < readings.Count) readings.RemoveRange(count, readings.Count - count);
    }

    private static List<ScoredEntity> RankSameSpans(List<ScoredEntity> candidates, EntityOverlapPolicy policy,
        Dictionary<ScoredEntity, List<ScoredEntity>> alternatives, CancellationToken cancellationToken) {
        var distinct = new List<ScoredEntity>(candidates.Count);
        for (int start = 0; start < candidates.Count;) {
            cancellationToken.ThrowIfCancellationRequested();
            ScoredEntity winner = candidates[start];
            int end = start + 1;
            while (end < candidates.Count && candidates[end].Start == winner.Start && candidates[end].End == winner.End) end++;

            if (end - start > 1) {
                for (int index = start; index < end; index++) {
                    ScoredEntity candidate = candidates[index];
                    candidate.Evidence |= EntityEvidence.AmbiguousReading;
                    if (!candidate.HasContext) candidate.RawConfidence -= 0.15 * (end - start - 1);
                }
                for (int index = start + 1; index < end; index++) {
                    if (CompareByPriority(candidates[index], winner) < 0) winner = candidates[index];
                }
                if (policy == EntityOverlapPolicy.BestOnly) {
                    var others = new List<ScoredEntity>(end - start - 1);
                    for (int index = start; index < end; index++) {
                        if (!ReferenceEquals(candidates[index], winner)) others.Add(candidates[index]);
                    }
                    others.Sort(CompareByPriority);
                    alternatives.Add(winner, others);
                }
            }
            if (policy == EntityOverlapPolicy.All) {
                for (int index = start; index < end; index++) distinct.Add(candidates[index]);
            } else {
                distinct.Add(winner);
            }
            start = end;
        }
        return distinct;
    }

    private static HashSet<ScoredEntity> FindContained(List<ScoredEntity> readings, double threshold,
        Dictionary<ScoredEntity, List<ScoredEntity>> children, CancellationToken cancellationToken) {
        var contained = new HashSet<ScoredEntity>();
        var owners = new Dictionary<ScoredEntity, ScoredEntity>();

        for (int index = 0; index < readings.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            ScoredEntity outer = readings[index];
            if (outer.Confidence < threshold || outer.Scheme.Contains.Count == 0) continue;
            for (int innerIndex = index + 1; innerIndex < readings.Count && readings[innerIndex].Start < outer.End; innerIndex++) {
                ScoredEntity inner = readings[innerIndex];
                if (inner.End > outer.End || !DeclaresContained(outer.Scheme, inner.Scheme)) continue;
                // Choose the smallest enclosing declaration so nested entities have one immediate owner.
                if (owners.TryGetValue(inner, out ScoredEntity? existing) && existing.Candidate.Length <= outer.Candidate.Length) continue;
                owners[inner] = outer;
            }
        }
        foreach (KeyValuePair<ScoredEntity, ScoredEntity> relation in owners) {
            contained.Add(relation.Key);
            AddReading(children, relation.Value, relation.Key);
        }
        foreach (List<ScoredEntity> group in children.Values) group.Sort(CompareByPosition);
        return contained;
    }

    private static void AddReading(Dictionary<ScoredEntity, List<ScoredEntity>> groups, ScoredEntity owner, ScoredEntity reading) {
        if (!groups.TryGetValue(owner, out List<ScoredEntity>? group)) {
            group = new List<ScoredEntity>();
            groups.Add(owner, group);
        }
        group.Add(reading);
    }

    private static bool DeclaresContained(EntityScheme outer, EntityScheme inner) {
        foreach (string id in outer.Contains) {
            if (string.Equals(id, inner.Id, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static EntityMatch ToMatch(ScoredEntity reading, EntitySegmentChain chain,
        Dictionary<ScoredEntity, List<ScoredEntity>> alternatives, Dictionary<ScoredEntity, List<ScoredEntity>> children,
        Dictionary<ScoredEntity, List<ScoredEntity>> maskingFallbacks, Dictionary<ScoredEntity, EntityMatch> built) {
        EntityMatch[]? otherMatches = alternatives.TryGetValue(reading, out List<ScoredEntity>? other)
            ? other.ConvertAll(item => new EntityMatch(item, chain)).ToArray() : null;
        EntityMatch[]? childMatches = children.TryGetValue(reading, out List<ScoredEntity>? inner)
            ? inner.ConvertAll(item => built[item]).ToArray() : null;
        EntityMatch[]? fallbackMatches = maskingFallbacks.TryGetValue(reading, out List<ScoredEntity>? fallbacks)
            ? fallbacks.ConvertAll(item => new EntityMatch(item, chain)).ToArray() : null;
        return new EntityMatch(reading, chain, otherMatches, childMatches, fallbackMatches);
    }

    private static int CompareByPosition(ScoredEntity left, ScoredEntity right) {
        int result = left.Start.CompareTo(right.Start);
        if (result != 0) return result;
        result = right.End.CompareTo(left.End);
        return result != 0 ? result : left.RegistrationOrder.CompareTo(right.RegistrationOrder);
    }

    private static int CompareByPriority(ScoredEntity left, ScoredEntity right) {
        int result = right.Confidence.CompareTo(left.Confidence);
        if (result != 0) return result;
        result = right.Candidate.Length.CompareTo(left.Candidate.Length);
        if (result != 0) return result;
        result = EntityScoring.BaseConfidence(right.Candidate.Strength).CompareTo(EntityScoring.BaseConfidence(left.Candidate.Strength));
        if (result != 0) return result;
        result = left.RegistrationOrder.CompareTo(right.RegistrationOrder);
        return result != 0 ? result : left.Start.CompareTo(right.Start);
    }
}
