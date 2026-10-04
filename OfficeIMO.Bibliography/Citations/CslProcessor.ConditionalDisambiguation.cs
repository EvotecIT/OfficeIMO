using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

public sealed partial class CslProcessor {
    private CslDisambiguationValue[] DisambiguationValues(CslRecord[] records, IReadOnlyList<CslDisambiguationForm> forms,
        CslEvaluator evaluator, XElement layout, bool observeConditions) {
        var values = new List<CslDisambiguationValue>();
        foreach (CslRecord record in records) foreach (CslDisambiguationForm form in forms) {
            CslContext context = CreateDisambiguationContext(record, form);
            if (observeConditions) context.ObservedConditions = new List<string>();
            string text = evaluator.Evaluate(layout, context).Plain;
            values.Add(new CslDisambiguationValue(record, form, text, context.ObservedConditions));
        }
        return values.ToArray();
    }

    /// <summary>Tries conditional detail in style order, retaining only choices that reduce collisions between different works.</summary>
    private void SelectConditionalDetail(CslRecord[] records, IReadOnlyList<CslDisambiguationForm> forms, CslEvaluator evaluator, XElement layout, bool reassignSuffixes = false) {
        CslDisambiguationValue[] values = DisambiguationValues(records, forms, evaluator, layout, true);
        while (true) {
            List<CslAmbiguity> ambiguous = Ambiguities(values);
            bool changed = false;
            foreach (CslConditionAttempt attempt in ConditionalAttempts(ambiguous.SelectMany(group => group.Values), evaluator)) {
                CslDisambiguationValue[]? improved = TryConditionalDetail(attempt, values, records, forms, evaluator, layout, 0, reassignSuffixes);
                if (improved == null) continue;
                values = improved;
                changed = true;
                break;
            }
            if (changed) continue;
            CslDisambiguationValue[]? combination = CombinedConditionalDetail(values, records, forms, evaluator, layout, reassignSuffixes);
            if (combination == null) return;
            values = combination;
        }
    }

    /// <summary>Handles groups whose literal and variable branches only help together, then removes unnecessary choices.</summary>
    private CslDisambiguationValue[]? CombinedConditionalDetail(CslDisambiguationValue[] before, CslRecord[] records,
        IReadOnlyList<CslDisambiguationForm> forms, CslEvaluator evaluator, XElement layout, bool reassignSuffixes) {
        ulong targetScore = AmbiguityScore(before, evaluator);
        CslDisambiguationValue[] values = before;
        var trials = new List<CslConditionAttempt>();
        string[]? originalSuffixes = SnapshotSuffixes(records, reassignSuffixes);
        bool accepted = false;
        try {
            while (true) {
                CslConditionAttempt? attempt = ConditionalAttempts(Ambiguities(values).SelectMany(group => group.Values), evaluator).FirstOrDefault();
                if (attempt == null) return null;
                foreach (var choice in attempt.Choices) choice.Record.ActiveConditions.Add(choice.Key);
                trials.Add(attempt);
                values = ConditionalValues(records, forms, evaluator, layout, reassignSuffixes);
                if (AmbiguityScore(values, evaluator) >= targetScore) continue;
                // Preserve the earliest useful branches when several choices
                // can supply the same detail. Removal must not worsen the result.
                for (int index = trials.Count - 1; index >= 0; index--) {
                    CslConditionAttempt removable = trials[index];
                    string[]? priorSuffixes = SnapshotSuffixes(records, reassignSuffixes);
                    foreach (var choice in removable.Choices) choice.Record.ActiveConditions.Remove(choice.Key);
                    CslDisambiguationValue[] reduced = ConditionalValues(records, forms, evaluator, layout, reassignSuffixes);
                    if (AmbiguityScore(reduced, evaluator) <= AmbiguityScore(values, evaluator)) {
                        trials.RemoveAt(index);
                        values = reduced;
                    } else {
                        foreach (var choice in removable.Choices) choice.Record.ActiveConditions.Add(choice.Key);
                        RestoreSuffixes(records, priorSuffixes);
                    }
                }
                accepted = true;
                return values;
            }
        } finally {
            if (!accepted) {
                foreach (CslConditionAttempt trial in trials) foreach (var choice in trial.Choices) choice.Record.ActiveConditions.Remove(choice.Key);
                RestoreSuffixes(records, originalSuffixes);
            }
        }
    }

    private CslDisambiguationValue[]? TryConditionalDetail(CslConditionAttempt attempt, CslDisambiguationValue[] before, CslRecord[] records,
        IReadOnlyList<CslDisambiguationForm> forms, CslEvaluator evaluator, XElement layout, int depth, bool reassignSuffixes, ulong? requiredScore = null) {
        evaluator.PerformOperation();
        if (depth >= _style.MaximumDepth) throw new InvalidDataException("CSL conditional disambiguation exceeds MaximumNestingDepth.");
        var selected = new List<(CslRecord Record, string Key)>(attempt.Choices);
        string[]? originalSuffixes = SnapshotSuffixes(records, reassignSuffixes);
        foreach (var choice in selected) choice.Record.ActiveConditions.Add(choice.Key);
        bool accepted = false;
        try {
            ulong targetScore = requiredScore ?? AmbiguityScore(before, evaluator);
            CslDisambiguationValue[] after;
            while (true) {
                after = ConditionalValues(records, forms, evaluator, layout, reassignSuffixes);
                if (AmbiguityScore(after, evaluator) < targetScore) {
                    accepted = true;
                    return after;
                }
                // Detail can reveal a collision with a previously unique work.
                // Try the same branch there before deciding the original trial
                // was unhelpful. Each wave adds previously inactive choices.
                var propagated = ConditionalAttempts(Ambiguities(after).SelectMany(group => group.Values), evaluator)
                    .Where(candidate => candidate.Identity == attempt.Identity).SelectMany(candidate => candidate.Choices).ToArray();
                if (propagated.Length == 0) break;
                foreach (var choice in propagated) { choice.Record.ActiveConditions.Add(choice.Key); selected.Add(choice); }
            }
            // An outer condition may expose useful nested detail without itself
            // separating the works. Retain the gate only when a descendant helps.
            var affected = new HashSet<CslRecord>(selected.Select(choice => choice.Record));
            var observed = before.Where(value => affected.Contains(value.Record))
                .GroupBy(value => value.Record).ToDictionary(group => group.Key,
                    group => new HashSet<string>(group.SelectMany(value => value.Conditions!), StringComparer.Ordinal));
            IEnumerable<CslConditionAttempt> nested = ConditionalAttempts(after.Where(value => observed.ContainsKey(value.Record)), evaluator);
            foreach (CslConditionAttempt child in nested) {
                var exposed = child.Choices.Where(choice => !observed[choice.Record].Contains(choice.Key)).ToArray();
                if (exposed.Length == 0) continue;
                CslDisambiguationValue[]? improved = TryConditionalDetail(new CslConditionAttempt(exposed), after,
                    records, forms, evaluator, layout, depth + 1, reassignSuffixes, targetScore);
                if (improved == null) continue;
                accepted = true;
                return improved;
            }
            return null;
        } finally {
            if (!accepted) {
                foreach (var choice in selected) choice.Record.ActiveConditions.Remove(choice.Key);
                RestoreSuffixes(records, originalSuffixes);
            }
        }
    }

    private CslDisambiguationValue[] ConditionalValues(CslRecord[] records, IReadOnlyList<CslDisambiguationForm> forms,
        CslEvaluator evaluator, XElement layout, bool reassignSuffixes) {
        if (reassignSuffixes) AssignYearSuffixes(records, forms, evaluator, layout);
        return DisambiguationValues(records, forms, evaluator, layout, true);
    }

    private static string[]? SnapshotSuffixes(CslRecord[] records, bool required) => required ? records.Select(record => record.YearSuffix).ToArray() : null;
    private static void RestoreSuffixes(CslRecord[] records, string[]? suffixes) {
        if (suffixes != null) for (int index = 0; index < records.Length; index++) records[index].YearSuffix = suffixes[index];
    }

    private static IEnumerable<CslConditionAttempt> ConditionalAttempts(IEnumerable<CslDisambiguationValue> values, CslEvaluator evaluator) {
        var candidates = new SortedDictionary<string, HashSet<(CslRecord Record, string Key)>>(StringComparer.Ordinal);
        foreach (CslDisambiguationValue value in values) foreach (string key in value.Conditions!) {
            evaluator.PerformOperation();
            if (value.Record.ActiveConditions.Contains(key)) continue;
            // The same branch should be tried together for all colliding forms;
            // unique forms never join that trial. Macro invocations stay distinct.
            string identity = key.Substring(2);
            if (!candidates.TryGetValue(identity, out var choices)) candidates.Add(identity, choices = new HashSet<(CslRecord Record, string Key)>());
            choices.Add((value.Record, key));
        }
        return candidates.Select(pair => new CslConditionAttempt(pair.Value.ToArray()));
    }

    private static ulong AmbiguityScore(IEnumerable<CslDisambiguationValue> values, CslEvaluator evaluator) {
        ulong score = 0;
        foreach (IGrouping<string, CslDisambiguationValue> group in values.GroupBy(value => value.Text, StringComparer.Ordinal)) {
            ulong preceding = 0;
            foreach (IGrouping<CslRecord, CslDisambiguationValue> work in group.GroupBy(value => value.Record)) {
                evaluator.PerformOperation();
                ulong count = (ulong)work.Count();
                // Count form pairs of different works, never pairs of the same
                // work. Separating first notes is progress even if shorts tie.
                score += preceding * count;
                preceding += count;
            }
        }
        return score;
    }

    private sealed class CslConditionAttempt {
        internal CslConditionAttempt((CslRecord Record, string Key)[] choices) { Choices = choices; }
        internal (CslRecord Record, string Key)[] Choices { get; }
        internal string Identity => Choices[0].Key.Substring(2);
    }
}
