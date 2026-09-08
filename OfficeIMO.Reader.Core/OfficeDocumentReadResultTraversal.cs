using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentReadResultExtensions {
    /// <summary>
    /// Enumerates document-level and page-level blocks in canonical source order, retaining each stable ID or anchor once.
    /// Blocks without either identity are deduplicated only by object reference, preserving unlabelled repeated text.
    /// The returned blocks are the source model objects, not immutable copies.
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
