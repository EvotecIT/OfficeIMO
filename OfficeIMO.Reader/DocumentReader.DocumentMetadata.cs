using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static IReadOnlyList<OfficeDocumentMetadataEntry> BuildChunkDocumentMetadata(
        ReaderInputKind kind,
        IReadOnlyList<ReaderChunk> chunks,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<ReaderTable> tables,
        IReadOnlyList<ReaderVisual> visuals,
        IReadOnlyList<OfficeDocumentPage> pages,
        IReadOnlyList<OfficeDocumentAsset> assets) {
        var entries = new List<OfficeDocumentMetadataEntry>();
        AddCountMetadata(entries, "reader-chunk-count", "reader.summary", "ChunkCount", chunks.Count);
        AddCountMetadata(entries, "reader-block-count", "reader.summary", "BlockCount", blocks.Count);
        AddCountMetadata(entries, "reader-table-count", "reader.summary", "TableCount", tables.Count);
        AddCountMetadata(entries, "reader-visual-count", "reader.summary", "VisualCount", visuals.Count);
        AddCountMetadata(entries, "reader-asset-count", "reader.summary", "AssetCount", assets.Count);

        if (kind == ReaderInputKind.Excel) {
            AddCountMetadata(entries, "reader-sheet-count", "reader.container", "SheetCount", pages.Count(static page => string.Equals(page.Location.SourceBlockKind, "sheet", StringComparison.Ordinal)));
        } else if (kind == ReaderInputKind.PowerPoint) {
            AddCountMetadata(entries, "reader-slide-count", "reader.container", "SlideCount", pages.Count(static page => string.Equals(page.Location.SourceBlockKind, "slide", StringComparison.Ordinal)));
        } else if (kind == ReaderInputKind.Pdf) {
            AddCountMetadata(entries, "reader-page-count", "reader.container", "PageCount", pages.Count(static page => string.Equals(page.Location.SourceBlockKind, "page", StringComparison.Ordinal)));
        }

        return entries.Count == 0 ? Array.Empty<OfficeDocumentMetadataEntry>() : entries.AsReadOnly();
    }

    private static void AddCountMetadata(List<OfficeDocumentMetadataEntry> entries, string id, string category, string name, int count) {
        if (count == 0) {
            return;
        }

        entries.Add(new OfficeDocumentMetadataEntry {
            Id = id,
            Category = category,
            Name = name,
            Value = count.ToString(CultureInfo.InvariantCulture),
            ValueType = "count"
        });
    }
}
