using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    internal static IEnumerable<ReaderTable> Tables(OfficeDocumentReadResult document) {
        var seen = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var aggregateIdentityCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var canonicalMatches = new Dictionary<string, Queue<ReaderTable>>(StringComparer.Ordinal);
        void RegisterCanonical(ReaderTable original, ReaderTable projected) {
            string key = BuildTableIdentity(original, includeLocation: false, includeAnchor: false);
            if (!canonicalMatches.TryGetValue(key, out var matches)) canonicalMatches.Add(key, matches = new());
            matches.Enqueue(projected);
        }
        var pageMatches = new Dictionary<string, Queue<(ReaderTable Table, OfficeDocumentPage Page, int Index)>>(StringComparer.Ordinal);
        var pageReferences = new Dictionary<ReaderTable, (OfficeDocumentPage Page, int Index)>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var matchedPages = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var aggregateReferences = new HashSet<ReaderTable>(document.Tables ?? Array.Empty<ReaderTable>(), ReferenceIdentityComparer<ReaderTable>.Instance);
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            if (page?.Tables == null) continue;
            for (int index = 0; index < page.Tables.Count; index++) {
                ReaderTable table = page.Tables[index];
                if (table == null) continue;
                if (!pageReferences.ContainsKey(table)) pageReferences.Add(table, (page, index));
                string rawIdentity = BuildTableIdentity(table);
                string scopedIdentity = BuildTableIdentity(table, page, index);
                AddMatch(rawIdentity, table, page, index);
                if (rawIdentity != scopedIdentity) AddMatch(scopedIdentity, table, page, index);
                AddMatch("unscoped:" + BuildTableIdentity(table, includeLocation: false), table, page, index);
            }
        }
        void AddMatch(string identity, ReaderTable table, OfficeDocumentPage page, int index) {
            if (!pageMatches.TryGetValue(identity, out var matches)) {
                matches = new Queue<(ReaderTable, OfficeDocumentPage, int)>();
                pageMatches.Add(identity, matches);
            }
            matches.Enqueue((table, page, index));
        }
        foreach (ReaderTable table in document.Tables ?? System.Array.Empty<ReaderTable>()) {
            if (table == null || !seen.Add(table)) continue;
            ReaderTable projected = table;
            if (pageReferences.TryGetValue(table, out var pageReference)) {
                matchedPages.Add(table);
                projected = WithPageLocationFallback(table, pageReference.Page, pageReference.Index);
            } else {
                var keys = new List<string> { BuildTableIdentity(table) };
                if (table.Location?.Page == null && table.Location?.Slide == null && string.IsNullOrWhiteSpace(table.Location?.Sheet))
                    keys.Insert(0, "unscoped:" + BuildTableIdentity(table, includeLocation: false));
                foreach (string key in keys) {
                    if (!pageMatches.TryGetValue(key, out var matches)) continue;
                    int remaining = matches.Count;
                    bool found = false;
                    while (remaining-- > 0) {
                        var match = matches.Dequeue();
                        if (matchedPages.Contains(match.Table) || aggregateReferences.Contains(match.Table)) continue;
                        ReaderTable candidate = WithPageLocationFallback(match.Table, match.Page, match.Index);
                        ReaderTable proposed = WithLocationFallback(table, candidate.Location!, candidate.Location!.TableIndex);
                        // Use the canonical located identity: checking only page/path would erase
                        // distinct source blocks, ranges or heading positions with identical cells.
                        if (BuildTableIdentity(proposed) != BuildTableIdentity(candidate)) {
                            matches.Enqueue(match);
                            continue;
                        }
                        matchedPages.Add(match.Table);
                        projected = proposed;
                        found = true;
                        break;
                    }
                    if (found) break;
                }
            }
            // Matching counts retain repeated equal tables while reconciling aggregate/page copies
            // whose object references were separated by a Reader JSON round-trip.
            IncrementIdentity(aggregateIdentityCounts, BuildTableIdentity(projected, null, null));
            RegisterCanonical(table, projected);
            yield return projected;
        }
        foreach (OfficeDocumentPage page in document.Pages ?? System.Array.Empty<OfficeDocumentPage>()) {
            if (page?.Tables == null) continue;
            for (int tableIndex = 0; tableIndex < page.Tables.Count; tableIndex++) {
                ReaderTable table = page.Tables[tableIndex];
                if (table == null || matchedPages.Contains(table) || !seen.Add(table)) continue;
                ReaderTable scopedTable = WithPageLocationFallback(table, page, tableIndex);
                string identity = BuildTableIdentity(scopedTable, null, null);
                if (aggregateIdentityCounts.ContainsKey(identity)) continue;
                IncrementIdentity(aggregateIdentityCounts, identity);
                RegisterCanonical(table, scopedTable);
                yield return scopedTable;
            }
        }
        var chunkSeen = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        foreach (ReaderChunk chunk in document.Chunks ?? Array.Empty<ReaderChunk>()) {
            if (chunk?.Tables == null) continue;
            foreach (ReaderTable table in chunk.Tables) {
                if (table == null || !chunkSeen.Add(table)) continue;
                ReaderTable projected = WithLocationFallback(table, chunk.Location ?? new ReaderLocation(), null);
                string key = BuildTableIdentity(table, includeLocation: false, includeAnchor: false);
                bool matched = false;
                if (canonicalMatches.TryGetValue(key, out var matches)) {
                    int remaining = matches.Count;
                    while (remaining-- > 0) {
                        ReaderTable candidate = matches.Dequeue();
                        ReaderTable proposed = WithLocationFallback(projected, candidate.Location ?? new ReaderLocation(), candidate.Location?.TableIndex);
                        // Missing table ordinals inherit from the matched canonical table. A global chunk
                        // index is not comparable with a page-local ordinal after JSON separates references.
                        ReaderTable comparable = WithLocationFallback(candidate, projected.Location ?? new ReaderLocation(), projected.Location?.TableIndex);
                        // Additional coordinates on either projection are compatible; explicit disagreements are not.
                        if (BuildTableIdentity(proposed) == BuildTableIdentity(comparable)) {
                            matched = true;
                            break;
                        }
                        matches.Enqueue(candidate);
                    }
                }
                if (matched || !seen.Add(table)) continue;
                yield return table;
            }
        }
    }

}
