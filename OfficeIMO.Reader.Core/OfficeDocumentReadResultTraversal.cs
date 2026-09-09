using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentReadResultExtensions {
    /// <summary>
    /// Enumerates blocks and tables together in canonical source order. Chunks supply text only when no blocks contain text.
    /// Tables inherit a matching source block's position through their stable anchor. Table rows remain part of their table.
    /// Unpositioned content retains its encounter order after positioned content in the same container.
    /// Results reference source payloads or location projections and are not immutable snapshots.
    /// </summary>
    public static IEnumerable<OfficeDocumentContentItem> EnumerateContent(this OfficeDocumentReadResult document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return OfficeDocumentModelTraversal.Content(document);
    }

    /// <summary>
    /// Enumerates document-level and page-level blocks in canonical source order, retaining each stable ID or anchor once.
    /// Blocks without either identity are deduplicated only by object reference, preserving unlabelled repeated text.
    /// Missing locations inherit their page context by reference or stable identity without changing the source.
    /// Results may refer to source objects or location projections; they are not immutable snapshots.
    /// </summary>
    public static IEnumerable<OfficeDocumentBlock> EnumerateBlocks(this OfficeDocumentReadResult document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return OfficeDocumentModelTraversal.Blocks(document);
    }

    /// <summary>
    /// Enumerates tables from document, page and chunk projections using the Reader's canonical duplicate handling.
    /// Page-only tables inherit missing location information from their page. Results may refer to source model objects.
    /// </summary>
    public static IEnumerable<ReaderTable> EnumerateTables(this OfficeDocumentReadResult document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return OfficeDocumentModelTraversal.Tables(document);
    }
}
