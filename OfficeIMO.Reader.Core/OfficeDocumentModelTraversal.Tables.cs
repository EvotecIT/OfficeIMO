using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    internal static IEnumerable<ReaderTable> Tables(OfficeDocumentReadResult document) {
        var seen = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var canonical = new List<ReaderTable>();
        var canonicalMatches = new TableMatchIndex<int>();
        var payloads = new Dictionary<ReaderTable, string>(ReferenceIdentityComparer<ReaderTable>.Instance);
        string Payload(ReaderTable table) {
            if (!payloads.TryGetValue(table, out var identity))
                payloads.Add(table, identity = BuildTableIdentity(table, includeLocation: false, includeAnchor: false));
            return identity;
        }
        void RegisterCanonical(ReaderTable original, ReaderTable projected) {
            canonicalMatches.Add(Payload(original), projected.Location, canonical.Count);
            canonical.Add(projected);
        }
        var pageMatches = new TableMatchIndex<ReaderTable>();
        var pageReferences = new Dictionary<ReaderTable, (OfficeDocumentPage Page, int Index)>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var matchedPages = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var aggregateReferences = new HashSet<ReaderTable>(document.Tables ?? Array.Empty<ReaderTable>(), ReferenceIdentityComparer<ReaderTable>.Instance);
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            if (page?.Tables == null) continue;
            for (int index = 0; index < page.Tables.Count; index++) {
                ReaderTable table = page.Tables[index];
                if (table == null) continue;
                if (!pageReferences.ContainsKey(table)) pageReferences.Add(table, (page, index));
                if (!aggregateReferences.Contains(table))
                    pageMatches.Add(Payload(table), WithPageLocationFallback(table, page, index).Location, table);
            }
        }
        foreach (ReaderTable table in document.Tables ?? System.Array.Empty<ReaderTable>()) {
            if (table == null || !seen.Add(table)) continue;
            ReaderTable projected = table;
            if (pageReferences.TryGetValue(table, out var pageReference)) {
                matchedPages.Add(table);
                projected = WithPageLocationFallback(table, pageReference.Page, pageReference.Index);
            } else {
                while (pageMatches.TryTake(Payload(table), table.Location, out var match)) {
                    if (!matchedPages.Add(match)) continue;
                    var scope = pageReferences[match];
                    ReaderTable candidate = WithPageLocationFallback(match, scope.Page, scope.Index);
                    projected = WithLocationFallback(table, candidate.Location!, candidate.Location!.TableIndex);
                    break;
                }
            }
            // Each matching page occurrence is consumed once, including after JSON separates references.
            RegisterCanonical(table, projected);
        }
        foreach (OfficeDocumentPage page in document.Pages ?? System.Array.Empty<OfficeDocumentPage>()) {
            if (page?.Tables == null) continue;
            for (int tableIndex = 0; tableIndex < page.Tables.Count; tableIndex++) {
                ReaderTable table = page.Tables[tableIndex];
                if (table == null || matchedPages.Contains(table) || !seen.Add(table)) continue;
                ReaderTable scopedTable = WithPageLocationFallback(table, page, tableIndex);
                RegisterCanonical(table, scopedTable);
            }
        }
        var chunkSeen = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        foreach (ReaderChunk chunk in document.Chunks ?? Array.Empty<ReaderChunk>()) {
            if (chunk?.Tables == null) continue;
            foreach (ReaderTable table in chunk.Tables) {
                if (table == null || !chunkSeen.Add(table)) continue;
                ReaderTable projected = WithLocationFallback(table, chunk.Location ?? new ReaderLocation(), null);
                if (canonicalMatches.TryTake(Payload(table), projected.Location, out int candidateIndex)) {
                    ReaderTable candidate = canonical[candidateIndex];
                    canonical[candidateIndex] = WithLocationFallback(candidate, projected.Location ?? new ReaderLocation(), projected.Location?.TableIndex);
                    continue;
                }
                if (!seen.Add(table)) continue;
                canonical.Add(projected);
            }
        }
        // Reconcile all projections before exposing tables so later chunk coordinates are retained.
        return canonical;
    }

}
