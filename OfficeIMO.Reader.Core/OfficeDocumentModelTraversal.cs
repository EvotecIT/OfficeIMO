using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;

namespace OfficeIMO.Reader;

internal static class OfficeDocumentModelTraversal {
    internal static IEnumerable<OfficeDocumentBlock> Blocks(OfficeDocumentReadResult document) {
        var seen = new HashSet<OfficeDocumentBlock>(ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        var ordered = new List<OrderedBlock>();
        int insertionIndex = 0;
        foreach (OfficeDocumentBlock block in document.Blocks ?? System.Array.Empty<OfficeDocumentBlock>()) {
            if (block != null && seen.Add(block)) ordered.Add(new OrderedBlock(block, insertionIndex++));
        }
        foreach (OfficeDocumentPage page in document.Pages ?? System.Array.Empty<OfficeDocumentPage>()) {
            if (page?.Blocks == null) continue;
            foreach (OfficeDocumentBlock block in page.Blocks) {
                if (block != null && seen.Add(block)) ordered.Add(new OrderedBlock(block, insertionIndex++));
            }
        }
        ordered.Sort(CompareBlocks);
        for (int index = 0; index < ordered.Count; index++) yield return ordered[index].Block;
    }

    internal static IEnumerable<ReaderTable> Tables(OfficeDocumentReadResult document) {
        var seen = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var aggregateIdentityCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (ReaderTable table in document.Tables ?? System.Array.Empty<ReaderTable>()) {
            if (table == null || !seen.Add(table)) continue;
            IncrementIdentity(aggregateIdentityCounts, BuildTableIdentity(table, null, null));
            yield return table;
        }
        foreach (OfficeDocumentPage page in document.Pages ?? System.Array.Empty<OfficeDocumentPage>()) {
            if (page?.Tables == null) continue;
            for (int tableIndex = 0; tableIndex < page.Tables.Count; tableIndex++) {
                ReaderTable table = page.Tables[tableIndex];
                if (table == null || !seen.Add(table)) continue;
                ReaderTable scopedTable = WithPageLocationFallback(table, page, tableIndex);
                string identity = BuildTableIdentity(scopedTable, null, null);
                if (aggregateIdentityCounts.ContainsKey(identity)) continue;
                IncrementIdentity(aggregateIdentityCounts, identity);
                yield return scopedTable;
            }
        }
        var chunkIdentityCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        int fallbackTableIndex = 0;
        foreach (ReaderChunk chunk in document.Chunks ?? System.Array.Empty<ReaderChunk>()) {
            if (chunk?.Tables == null) continue;
            foreach (ReaderTable table in chunk.Tables) {
                if (table == null || !seen.Add(table)) {
                    fallbackTableIndex++;
                    continue;
                }
                string identity = BuildTableIdentity(table, chunk.Location, fallbackTableIndex++);
                int occurrence = IncrementIdentity(chunkIdentityCounts, identity);
                if (aggregateIdentityCounts.TryGetValue(identity, out int aggregateCount) && occurrence <= aggregateCount) continue;
                yield return table;
            }
        }
    }

    internal static IEnumerable<ReaderTable> TableInstances(OfficeDocumentReadResult document) {
        var seen = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        foreach (ReaderTable table in document.Tables ?? Array.Empty<ReaderTable>()) {
            if (table != null && seen.Add(table)) yield return table;
        }
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            foreach (ReaderTable table in page?.Tables ?? Array.Empty<ReaderTable>()) {
                if (table != null && seen.Add(table)) yield return table;
            }
        }
        foreach (ReaderChunk chunk in document.Chunks ?? Array.Empty<ReaderChunk>()) {
            foreach (ReaderTable table in chunk?.Tables ?? Array.Empty<ReaderTable>()) {
                if (table != null && seen.Add(table)) yield return table;
            }
        }
    }

    internal static IEnumerable<ReaderVisual> Visuals(OfficeDocumentReadResult document) {
        var seen = new HashSet<ReaderVisual>(ReferenceIdentityComparer<ReaderVisual>.Instance);
        var aggregateIdentityCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (ReaderVisual visual in document.Visuals ?? Array.Empty<ReaderVisual>()) {
            if (visual == null || !seen.Add(visual)) continue;
            IncrementIdentity(aggregateIdentityCounts, BuildVisualIdentity(visual, null));
            yield return visual;
        }

        var chunkIdentityCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (ReaderChunk chunk in document.Chunks ?? Array.Empty<ReaderChunk>()) {
            if (chunk?.Visuals == null) continue;
            foreach (ReaderVisual visual in chunk.Visuals) {
                if (visual == null || !seen.Add(visual)) continue;
                string identity = BuildVisualIdentity(visual, chunk.Location);
                int occurrence = IncrementIdentity(chunkIdentityCounts, identity);
                if (aggregateIdentityCounts.TryGetValue(identity, out int aggregateCount) && occurrence <= aggregateCount) continue;
                yield return visual;
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

    internal static IEnumerable<OfficeDocumentFormField> Forms(OfficeDocumentReadResult document) {
        var seen = new HashSet<OfficeDocumentFormField>(ReferenceIdentityComparer<OfficeDocumentFormField>.Instance);
        var seenIds = new HashSet<string>(System.StringComparer.Ordinal);
        foreach (OfficeDocumentFormField form in document.Forms ?? System.Array.Empty<OfficeDocumentFormField>()) {
            if (form != null && seen.Add(form) &&
                (string.IsNullOrWhiteSpace(form.Id) || seenIds.Add(form.Id))) {
                yield return form;
            }
        }
        foreach (OfficeDocumentPage page in document.Pages ?? System.Array.Empty<OfficeDocumentPage>()) {
            if (page?.Forms == null) continue;
            foreach (OfficeDocumentFormField form in page.Forms) {
                if (form != null && seen.Add(form) &&
                    (string.IsNullOrWhiteSpace(form.Id) || seenIds.Add(form.Id))) {
                    yield return form;
                }
            }
        }
        IReadOnlyList<ReaderChunk> chunks = document.Chunks ?? System.Array.Empty<ReaderChunk>();
        for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
            ReaderChunk chunk = chunks[chunkIndex];
            if (chunk?.FormFields == null) continue;
            for (int index = 0; index < chunk.FormFields.Count; index++) {
                ReaderFormField field = chunk.FormFields[index];
                if (field == null) continue;
                OfficeDocumentFormField form = ProjectChunkForm(chunk, field, chunkIndex, index);
                if (seenIds.Add(form.Id)) yield return form;
            }
        }
    }

    private static OfficeDocumentFormField ProjectChunkForm(
        ReaderChunk chunk,
        ReaderFormField field,
        int chunkIndex,
        int fieldIndex) {
        ReaderFormWidget? widget = field.Widgets == null || field.Widgets.Count == 0 ? null : field.Widgets[0];
        return new OfficeDocumentFormField {
            Id = BuildChunkFormId(chunk, field, chunkIndex, fieldIndex),
            Name = FirstNonEmpty(field.Name, field.PartialName, field.AlternateName, field.MappingName),
            Kind = string.IsNullOrWhiteSpace(field.FieldType) ? field.Kind.ToString().ToLowerInvariant() : field.FieldType!,
            Value = BuildChunkFormValue(field),
            IsReadOnly = field.IsReadOnly,
            IsRequired = field.IsRequired,
            Location = BuildChunkFormLocation(chunk.Location, widget?.PageNumber, field.PageNumbers),
            Region = widget == null ? null : new OfficeDocumentRegion {
                X = widget.X1,
                Y = widget.Y1,
                Width = widget.Width,
                Height = widget.Height
            }
        };
    }

    private static string BuildChunkFormId(
        ReaderChunk chunk,
        ReaderFormField field,
        int chunkIndex,
        int fieldIndex) {
        string? identity = FirstNonEmpty(field.Name, field.PartialName, field.MappingName, field.AlternateName);
        if (!string.IsNullOrWhiteSpace(identity)) return identity!;
        string chunkId = string.IsNullOrWhiteSpace(chunk.Id)
            ? "chunk-" + chunkIndex.ToString("D4", System.Globalization.CultureInfo.InvariantCulture)
            : chunk.Id;
        return chunkId + "-form-" + fieldIndex.ToString("D4", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string? BuildChunkFormValue(ReaderFormField field) {
        if (field.Value != null) return field.Value;
        if (field.Values == null || field.Values.Count == 0) return null;
        return string.Join("\n", field.Values);
    }

    private static string? FirstNonEmpty(params string?[] values) {
        for (int index = 0; index < values.Length; index++) {
            if (!string.IsNullOrWhiteSpace(values[index])) return values[index];
        }
        return null;
    }

    private static ReaderLocation BuildChunkFormLocation(
        ReaderLocation? source,
        int? widgetPage,
        IReadOnlyList<int>? fieldPages) {
        source ??= new ReaderLocation();
        return new ReaderLocation {
            Path = source.Path,
            BlockIndex = source.BlockIndex,
            SourceBlockIndex = source.SourceBlockIndex,
            StartLine = source.StartLine,
            EndLine = source.EndLine,
            NormalizedStartLine = source.NormalizedStartLine,
            NormalizedEndLine = source.NormalizedEndLine,
            HeadingPath = source.HeadingPath,
            HeadingSlug = source.HeadingSlug,
            SourceBlockKind = source.SourceBlockKind,
            BlockAnchor = source.BlockAnchor,
            Sheet = source.Sheet,
            A1Range = source.A1Range,
            Slide = source.Slide,
            Page = widgetPage ?? (fieldPages == null || fieldPages.Count == 0 ? source.Page : fieldPages[0]),
            TableIndex = source.TableIndex
        };
    }

    private static int CompareBlocks(OrderedBlock left, OrderedBlock right) {
        int comparison = string.CompareOrdinal(BuildContainerOrderKey(left.Block.Location), BuildContainerOrderKey(right.Block.Location));
        if (comparison != 0) return comparison;
        comparison = BuildBlockPosition(left.Block.Location).CompareTo(BuildBlockPosition(right.Block.Location));
        return comparison != 0 ? comparison : left.InsertionIndex.CompareTo(right.InsertionIndex);
    }

    private static string BuildContainerOrderKey(ReaderLocation? location) {
        if (location == null) return "9|";
        if (location.Page.HasValue) return "0|" + location.Page.Value.ToString("D10", CultureInfo.InvariantCulture);
        if (location.Slide.HasValue) return "1|" + location.Slide.Value.ToString("D10", CultureInfo.InvariantCulture);
        if (!string.IsNullOrWhiteSpace(location.Sheet)) return "2|" + location.Sheet;
        return "9|";
    }

    private static int BuildBlockPosition(ReaderLocation? location) =>
        location?.SourceBlockIndex
        ?? location?.BlockIndex
        ?? location?.StartLine
        ?? location?.NormalizedStartLine
        ?? int.MaxValue;

    internal static string BuildBlockIdentity(OfficeDocumentBlock block) {
        if (!string.IsNullOrWhiteSpace(block.Id)) return "id:" + block.Id;
        string? anchor = block.Location?.BlockAnchor;
        if (!string.IsNullOrWhiteSpace(anchor)) return "anchor:" + anchor;
        return BuildLocatedIdentity(block, block.Location, block.Kind, block.Text);
    }

    internal static string BuildAssetIdentity(OfficeDocumentAsset asset) {
        if (!string.IsNullOrWhiteSpace(asset.Id)) return "id:" + asset.Id;
        if (!string.IsNullOrWhiteSpace(asset.SourceObjectId)) return "source:" + asset.SourceObjectId;
        if (!string.IsNullOrWhiteSpace(asset.PayloadHash)) return "hash:" + asset.PayloadHash;
        string? anchor = asset.Location?.BlockAnchor;
        if (!string.IsNullOrWhiteSpace(anchor)) return "anchor:" + anchor;
        return BuildLocatedIdentity(asset, asset.Location, asset.FileName, asset.MediaType, asset.Kind);
    }

    internal static string BuildLinkIdentity(OfficeDocumentLink link) {
        if (!string.IsNullOrWhiteSpace(link.Id)) return "id:" + link.Id;
        string? anchor = link.Location?.BlockAnchor;
        if (!string.IsNullOrWhiteSpace(anchor)) return "anchor:" + anchor;
        return BuildLocatedIdentity(link, link.Location, link.Uri, link.DestinationName, link.RemoteFile, link.Text);
    }

    internal static string BuildFormIdentity(OfficeDocumentFormField form) {
        if (!string.IsNullOrWhiteSpace(form.Id)) return "id:" + form.Id;
        string? anchor = form.Location?.BlockAnchor;
        if (!string.IsNullOrWhiteSpace(anchor)) return "anchor:" + anchor;
        return BuildLocatedIdentity(form, form.Location, form.Name, form.Kind);
    }

    private static string BuildLocatedIdentity<T>(T instance, ReaderLocation? location, params string?[] values) where T : class {
        var builder = new StringBuilder();
        AppendLocationIdentity(builder, location, null, null);
        bool hasLocation = location != null && (
            !string.IsNullOrWhiteSpace(location.Path) ||
            !string.IsNullOrWhiteSpace(location.Sheet) ||
            location.Page.HasValue ||
            location.Slide.HasValue ||
            location.BlockIndex.HasValue ||
            location.SourceBlockIndex.HasValue ||
            location.StartLine.HasValue ||
            location.TableIndex.HasValue);
        if (!hasLocation) return "reference:" + RuntimeHelpers.GetHashCode(instance).ToString(CultureInfo.InvariantCulture);
        for (int index = 0; index < values.Length; index++) AppendIdentity(builder, values[index]);
        return builder.ToString();
    }

    private static int IncrementIdentity(IDictionary<string, int> counts, string identity) {
        counts.TryGetValue(identity, out int count);
        count++;
        counts[identity] = count;
        return count;
    }

    internal static string BuildTableIdentity(ReaderTable table, ReaderLocation? fallback = null, int? fallbackTableIndex = null) {
        var builder = new StringBuilder();
        AppendIdentity(builder, table.PayloadHash);
        AppendIdentity(builder, table.CallId);
        AppendIdentity(builder, table.Kind);
        AppendIdentity(builder, table.Title);
        AppendLocationIdentity(builder, table.Location, fallback, fallbackTableIndex);
        AppendIdentity(builder, table.Columns);
        foreach (IReadOnlyList<string> row in table.Rows ?? Array.Empty<IReadOnlyList<string>>()) AppendIdentity(builder, row);
        AppendIdentity(builder, table.TotalRowCount.ToString(CultureInfo.InvariantCulture));
        return builder.ToString();
    }

    internal static string BuildTableIdentity(ReaderTable table, OfficeDocumentPage page, int tableIndex) =>
        BuildTableIdentity(WithPageLocationFallback(table, page, tableIndex));

    private static ReaderTable WithPageLocationFallback(ReaderTable table, OfficeDocumentPage page, int tableIndex) {
        ReaderLocation fallback = BuildPageLocation(page);
        if (table.Location != null && !NeedsLocationFallback(table.Location)) return table;
        return new ReaderTable {
            Title = table.Title,
            Kind = table.Kind,
            CallId = table.CallId,
            Summary = table.Summary,
            PayloadHash = table.PayloadHash,
            Location = MergeLocation(table.Location, fallback, tableIndex),
            Columns = table.Columns,
            ColumnProfiles = table.ColumnProfiles,
            Diagnostics = table.Diagnostics,
            Rows = table.Rows,
            TotalRowCount = table.TotalRowCount,
            Truncated = table.Truncated
        };
    }

    private static ReaderLocation BuildPageLocation(OfficeDocumentPage page) {
        ReaderLocation source = page.Location ?? new ReaderLocation();
        return new ReaderLocation {
            Path = source.Path,
            BlockIndex = source.BlockIndex,
            SourceBlockIndex = source.SourceBlockIndex,
            StartLine = source.StartLine,
            EndLine = source.EndLine,
            NormalizedStartLine = source.NormalizedStartLine,
            NormalizedEndLine = source.NormalizedEndLine,
            HeadingPath = source.HeadingPath,
            HeadingSlug = source.HeadingSlug,
            SourceBlockKind = source.SourceBlockKind,
            BlockAnchor = source.BlockAnchor,
            Sheet = source.Sheet,
            A1Range = source.A1Range,
            Slide = source.Slide,
            Page = source.Page ?? page.Number,
            TableIndex = source.TableIndex
        };
    }

    private static bool NeedsLocationFallback(ReaderLocation location) {
        return string.IsNullOrWhiteSpace(location.Path)
            || (!location.Page.HasValue && !location.Slide.HasValue && string.IsNullOrWhiteSpace(location.Sheet));
    }

    private static ReaderLocation MergeLocation(ReaderLocation? location, ReaderLocation fallback, int fallbackTableIndex) {
        return new ReaderLocation {
            Path = Prefer(location?.Path, fallback.Path),
            BlockIndex = location?.BlockIndex ?? fallback.BlockIndex,
            SourceBlockIndex = location?.SourceBlockIndex ?? fallback.SourceBlockIndex,
            StartLine = location?.StartLine ?? fallback.StartLine,
            EndLine = location?.EndLine ?? fallback.EndLine,
            NormalizedStartLine = location?.NormalizedStartLine ?? fallback.NormalizedStartLine,
            NormalizedEndLine = location?.NormalizedEndLine ?? fallback.NormalizedEndLine,
            HeadingPath = Prefer(location?.HeadingPath, fallback.HeadingPath),
            HeadingSlug = Prefer(location?.HeadingSlug, fallback.HeadingSlug),
            SourceBlockKind = Prefer(location?.SourceBlockKind, fallback.SourceBlockKind),
            BlockAnchor = Prefer(location?.BlockAnchor, fallback.BlockAnchor),
            Sheet = Prefer(location?.Sheet, fallback.Sheet),
            A1Range = Prefer(location?.A1Range, fallback.A1Range),
            Slide = location?.Slide ?? fallback.Slide,
            Page = location?.Page ?? fallback.Page,
            TableIndex = location?.TableIndex ?? fallback.TableIndex ?? fallbackTableIndex
        };
    }

    private static string BuildVisualIdentity(ReaderVisual visual, ReaderLocation? fallback) {
        var builder = new StringBuilder();
        AppendIdentity(builder, visual.PayloadHash);
        AppendIdentity(builder, visual.Kind);
        AppendIdentity(builder, visual.Language);
        AppendIdentity(builder, visual.SourceName);
        AppendIdentity(builder, visual.MimeType);
        AppendIdentity(builder, visual.Content);
        AppendLocationIdentity(builder, visual.Location, fallback, null);
        return builder.ToString();
    }

    private static void AppendLocationIdentity(
        StringBuilder builder,
        ReaderLocation? location,
        ReaderLocation? fallback,
        int? fallbackTableIndex) {
        AppendIdentity(builder, Prefer(location?.Path, fallback?.Path));
        AppendIdentity(builder, Prefer(location?.Sheet, fallback?.Sheet));
        AppendIdentity(builder, Prefer(location?.A1Range, fallback?.A1Range));
        AppendIdentity(builder, Prefer(location?.HeadingPath, fallback?.HeadingPath));
        AppendIdentity(builder, Prefer(location?.HeadingSlug, fallback?.HeadingSlug));
        AppendIdentity(builder, Prefer(location?.SourceBlockKind, fallback?.SourceBlockKind));
        AppendIdentity(builder, Prefer(location?.BlockAnchor, fallback?.BlockAnchor));
        AppendIdentity(builder, (location?.Page ?? fallback?.Page)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.Slide ?? fallback?.Slide)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.BlockIndex ?? fallback?.BlockIndex)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.SourceBlockIndex ?? fallback?.SourceBlockIndex)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.StartLine ?? fallback?.StartLine)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.TableIndex ?? fallbackTableIndex ?? fallback?.TableIndex)?.ToString(CultureInfo.InvariantCulture));
    }

    private static string? Prefer(string? value, string? fallback) =>
        string.IsNullOrWhiteSpace(value) ? fallback : value;

    private static void AppendIdentity(StringBuilder builder, IReadOnlyList<string>? values) {
        if (values == null) {
            AppendIdentity(builder, (string?)null);
            return;
        }
        AppendIdentity(builder, values.Count.ToString(CultureInfo.InvariantCulture));
        for (int index = 0; index < values.Count; index++) AppendIdentity(builder, values[index]);
    }

    private static void AppendIdentity(StringBuilder builder, string? value) {
        if (value == null) {
            builder.Append("-1:");
            return;
        }
        builder.Append(value.Length.ToString(CultureInfo.InvariantCulture));
        builder.Append(':');
        builder.Append(value);
    }

    private readonly struct OrderedBlock {
        internal OrderedBlock(OfficeDocumentBlock block, int insertionIndex) {
            Block = block;
            InsertionIndex = insertionIndex;
        }

        internal OfficeDocumentBlock Block { get; }
        internal int InsertionIndex { get; }
    }
}

internal sealed class ReferenceIdentityComparer<T> : IEqualityComparer<T> where T : class {
    internal static ReferenceIdentityComparer<T> Instance { get; } = new ReferenceIdentityComparer<T>();

    public bool Equals(T? x, T? y) => ReferenceEquals(x, y);

    public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
}
