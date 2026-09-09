using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    internal static IEnumerable<ReaderTable> Tables(OfficeDocumentReadResult document) {
        var seen = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var canonical = new List<ReaderTable>();
        var canonicalReferences = new Dictionary<ReaderTable, int>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var canonicalPayloads = new List<string>();
        var chunkReferences = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var chunkProjections = new List<(ReaderTable Source, ReaderTable Projected)>();
        foreach (ReaderChunk chunk in document.Chunks ?? Array.Empty<ReaderChunk>())
            foreach (ReaderTable table in chunk?.Tables ?? Array.Empty<ReaderTable>())
                if (table != null && chunkReferences.Add(table))
                    chunkProjections.Add((table, WithLocationFallback(table, chunk!.Location ?? new ReaderLocation(), null)));
        var payloads = new Dictionary<ReaderTable, string>(ReferenceIdentityComparer<ReaderTable>.Instance);
        string Payload(ReaderTable table) {
            if (!payloads.TryGetValue(table, out var identity))
                payloads.Add(table, identity = BuildTableIdentity(table, includeLocation: false, includeAnchor: false, includeCoverage: false));
            return identity;
        }
        void RegisterCanonical(ReaderTable original, ReaderTable projected, ReaderTable? pageAlias = null) {
            canonicalReferences[original] = canonical.Count;
            if (pageAlias != null) canonicalReferences[pageAlias] = canonical.Count;
            canonicalPayloads.Add(Payload(original));
            canonical.Add(projected);
        }
        var pageReferences = new Dictionary<ReaderTable, (OfficeDocumentPage Page, int Index)>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var matchedPages = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var aggregateReferences = new HashSet<ReaderTable>(document.Tables ?? Array.Empty<ReaderTable>(), ReferenceIdentityComparer<ReaderTable>.Instance);
        var pageCandidates = new List<(ReaderTable Source, ReaderTable Projected)>();
        var pageKeys = new List<TableMatchKey>();
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            if (page?.Tables == null) continue;
            for (int index = 0; index < page.Tables.Count; index++) {
                ReaderTable table = page.Tables[index];
                if (table == null || pageReferences.ContainsKey(table)) continue;
                pageReferences.Add(table, (page, index));
                if (!aggregateReferences.Contains(table)) {
                    ReaderTable projected = WithPageLocationFallback(table, page, index);
                    pageCandidates.Add((table, projected));
                    pageKeys.Add(new(Payload(table), projected.Location));
                }
            }
        }
        var aggregateQueryIndices = new Dictionary<ReaderTable, int>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var aggregateKeys = new List<TableMatchKey>();
        foreach (ReaderTable table in document.Tables ?? Array.Empty<ReaderTable>()) {
            if (table == null || pageReferences.ContainsKey(table) || aggregateQueryIndices.ContainsKey(table)) continue;
            aggregateQueryIndices.Add(table, aggregateKeys.Count);
            aggregateKeys.Add(new(Payload(table), table.Location));
        }
        int[] pageAssignments = AssignTableOccurrences(aggregateKeys, pageKeys);
        foreach (ReaderTable table in document.Tables ?? Array.Empty<ReaderTable>()) {
            if (table == null || !seen.Add(table)) continue;
            ReaderTable projected = table;
            ReaderTable? pageAlias = null;
            if (pageReferences.TryGetValue(table, out var pageReference)) {
                matchedPages.Add(table);
                projected = WithPageLocationFallback(table, pageReference.Page, pageReference.Index);
            } else {
                int match = pageAssignments[aggregateQueryIndices[table]];
                if (match >= 0) {
                    var candidate = pageCandidates[match];
                    matchedPages.Add(candidate.Source);
                    projected = MergeTableProjection(table, candidate.Projected);
                    pageAlias = candidate.Source;
                }
            }
            RegisterCanonical(table, projected, pageAlias);
        }
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            if (page?.Tables == null) continue;
            for (int tableIndex = 0; tableIndex < page.Tables.Count; tableIndex++) {
                ReaderTable table = page.Tables[tableIndex];
                if (table == null || matchedPages.Contains(table) || !seen.Add(table)) continue;
                RegisterCanonical(table, WithPageLocationFallback(table, page, tableIndex));
            }
        }
        // Source references are authoritative matches; reserve them before comparing copies.
        var reserved = new HashSet<int>();
        var chunkMatches = new Dictionary<ReaderTable, int>(ReferenceIdentityComparer<ReaderTable>.Instance);
        foreach (var chunk in chunkProjections)
            if (canonicalReferences.TryGetValue(chunk.Source, out int index) && reserved.Add(index))
                chunkMatches.Add(chunk.Source, index);
        var availableCanonical = new List<int>();
        var canonicalKeys = new List<TableMatchKey>();
        for (int index = 0; index < canonical.Count; index++) {
            if (reserved.Contains(index)) continue;
            availableCanonical.Add(index);
            canonicalKeys.Add(new(canonicalPayloads[index], canonical[index].Location));
        }
        var chunkQueries = new List<ReaderTable>();
        var chunkKeys = new List<TableMatchKey>();
        foreach (var chunk in chunkProjections) {
            if (chunkMatches.ContainsKey(chunk.Source)) continue;
            chunkQueries.Add(chunk.Source);
            chunkKeys.Add(new(Payload(chunk.Source), chunk.Projected.Location));
        }
        int[] chunkAssignments = AssignTableOccurrences(chunkKeys, canonicalKeys);
        for (int index = 0; index < chunkAssignments.Length; index++)
            if (chunkAssignments[index] >= 0) chunkMatches.Add(chunkQueries[index], availableCanonical[chunkAssignments[index]]);
        foreach (var chunk in chunkProjections) {
            if (chunkMatches.TryGetValue(chunk.Source, out int candidateIndex)) {
                canonical[candidateIndex] = MergeTableProjection(canonical[candidateIndex], chunk.Projected);
                continue;
            }
            if (!canonicalReferences.ContainsKey(chunk.Source) && !seen.Add(chunk.Source)) continue;
            canonical.Add(chunk.Projected);
        }
        return canonical;
    }
}