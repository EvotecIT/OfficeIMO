using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Reader;

internal static class OfficeDocumentModelTraversal {
    internal static IEnumerable<OfficeDocumentBlock> Blocks(OfficeDocumentReadResult document) {
        var seen = new HashSet<OfficeDocumentBlock>(ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        foreach (OfficeDocumentBlock block in document.Blocks ?? System.Array.Empty<OfficeDocumentBlock>()) {
            if (block != null && seen.Add(block)) yield return block;
        }
        foreach (OfficeDocumentPage page in document.Pages ?? System.Array.Empty<OfficeDocumentPage>()) {
            if (page?.Blocks == null) continue;
            foreach (OfficeDocumentBlock block in page.Blocks) {
                if (block != null && seen.Add(block)) yield return block;
            }
        }
    }

    internal static IEnumerable<ReaderTable> Tables(OfficeDocumentReadResult document) {
        var seen = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        foreach (ReaderTable table in document.Tables ?? System.Array.Empty<ReaderTable>()) {
            if (table != null && seen.Add(table)) yield return table;
        }
        foreach (OfficeDocumentPage page in document.Pages ?? System.Array.Empty<OfficeDocumentPage>()) {
            if (page?.Tables == null) continue;
            foreach (ReaderTable table in page.Tables) {
                if (table != null && seen.Add(table)) yield return table;
            }
        }
        foreach (ReaderChunk chunk in document.Chunks ?? System.Array.Empty<ReaderChunk>()) {
            if (chunk?.Tables == null) continue;
            foreach (ReaderTable table in chunk.Tables) {
                if (table != null && seen.Add(table)) yield return table;
            }
        }
    }

    internal static IEnumerable<OfficeDocumentLink> Links(OfficeDocumentReadResult document) {
        var seen = new HashSet<OfficeDocumentLink>(ReferenceIdentityComparer<OfficeDocumentLink>.Instance);
        foreach (OfficeDocumentLink link in document.Links ?? System.Array.Empty<OfficeDocumentLink>()) {
            if (link != null && seen.Add(link)) yield return link;
        }
        foreach (OfficeDocumentPage page in document.Pages ?? System.Array.Empty<OfficeDocumentPage>()) {
            if (page?.Links == null) continue;
            foreach (OfficeDocumentLink link in page.Links) {
                if (link != null && seen.Add(link)) yield return link;
            }
        }
    }
}

internal sealed class ReferenceIdentityComparer<T> : IEqualityComparer<T> where T : class {
    internal static ReferenceIdentityComparer<T> Instance { get; } = new ReferenceIdentityComparer<T>();

    public bool Equals(T? x, T? y) => ReferenceEquals(x, y);

    public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
}
