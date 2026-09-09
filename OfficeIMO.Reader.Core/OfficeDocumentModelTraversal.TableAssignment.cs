using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    private readonly struct TableMatchKey {
        internal TableMatchKey(string payload, ReaderLocation? location) { Payload = payload; Location = location; }
        internal string Payload { get; }
        internal ReaderLocation? Location { get; }
    }

    // Match a complete projection set before yielding it. Indexed greedy matching handles
    // ordinary occurrences; an iterative augmenting path repairs ambiguous partial locations
    // without duplicating an occurrence or recursing through a large document.
    private static int[] AssignTableOccurrences(IReadOnlyList<TableMatchKey> queries, IReadOnlyList<TableMatchKey> candidates) {
        var available = new TableMatchIndex<int>();
        var freeByPayload = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int index = 0; index < candidates.Count; index++) {
            available.Add(candidates[index].Payload, candidates[index].Location, index);
            freeByPayload.TryGetValue(candidates[index].Payload, out int count);
            freeByPayload[candidates[index].Payload] = count + 1;
        }
        int[] assignment = Enumerable.Repeat(-1, queries.Count).ToArray();
        int[] owners = Enumerable.Repeat(-1, candidates.Count).ToArray();
        var unmatched = new List<int>();
        foreach (int query in Enumerable.Range(0, queries.Count)
            .OrderByDescending(index => LocationCoordinates(queries[index].Location).Count(value => value != null))) {
            if (available.TryTake(queries[query].Payload, queries[query].Location, out int candidate)) {
                assignment[query] = candidate;
                owners[candidate] = query;
                freeByPayload[queries[query].Payload]--;
            } else unmatched.Add(query);
        }
        if (unmatched.Count == 0) return assignment;
        var all = new TableMatchIndex<int>();
        for (int index = 0; index < candidates.Count; index++)
            all.Add(candidates[index].Payload, candidates[index].Location, index);
        foreach (int root in unmatched) {
            if (!freeByPayload.TryGetValue(queries[root].Payload, out int freeCount) || freeCount == 0) continue;
            var pending = new Queue<int>();
            var visited = new HashSet<int> { root };
            var predecessors = new Dictionary<int, int>();
            pending.Enqueue(root);
            bool augmented = false;
            while (pending.Count > 0 && !augmented) {
                int query = pending.Dequeue();
                foreach (int candidate in all.Candidates(queries[query].Payload, queries[query].Location)) {
                    if (predecessors.ContainsKey(candidate)) continue;
                    predecessors.Add(candidate, query);
                    if (owners[candidate] >= 0) {
                        if (visited.Add(owners[candidate])) pending.Enqueue(owners[candidate]);
                        continue;
                    }
                    int free = candidate;
                    while (free >= 0) {
                        int nextQuery = predecessors[free];
                        int previous = assignment[nextQuery];
                        assignment[nextQuery] = free;
                        owners[free] = nextQuery;
                        free = previous;
                    }
                    augmented = true;
                    freeByPayload[queries[root].Payload]--;
                    break;
                }
            }
        }
        return assignment;
    }
}
