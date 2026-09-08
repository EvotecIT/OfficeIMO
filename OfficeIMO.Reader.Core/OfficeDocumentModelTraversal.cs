using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;

namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    // Mutation needs every distinct source instance, including separate aggregate/page copies of the same ID.
    internal static IEnumerable<OfficeDocumentBlock> BlockInstances(OfficeDocumentReadResult document) {
        var seen = new HashSet<OfficeDocumentBlock>(ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        foreach (OfficeDocumentBlock block in document.Blocks ?? Array.Empty<OfficeDocumentBlock>())
            if (block != null && seen.Add(block)) yield return block;
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>())
            foreach (OfficeDocumentBlock block in page?.Blocks ?? Array.Empty<OfficeDocumentBlock>())
                if (block != null && seen.Add(block)) yield return block;
    }

    internal static IEnumerable<OfficeDocumentBlock> Blocks(OfficeDocumentReadResult document) {
        var projections = new BlockProjectionIndex();
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            if (page == null) continue;
            foreach (OfficeDocumentBlock block in page.Blocks ?? Array.Empty<OfficeDocumentBlock>())
                if (block != null) projections.Add(block, page);
        }
        IEnumerable<OfficeDocumentBlock> candidates =
            (document.Blocks ?? System.Array.Empty<OfficeDocumentBlock>())
            .Concat((document.Pages ?? System.Array.Empty<OfficeDocumentPage>())
                .Where(page => page?.Blocks != null)
                .SelectMany(page => page.Blocks));
        OfficeDocumentBlock[] materialized = candidates.Where(block => block != null).ToArray();
        foreach (OfficeDocumentBlock block in OrderBlocks(materialized, projections.ResolveLocation)) {
            ReaderLocation? location = projections.ResolveLocation(block);
            // Project fallback locations without mutating the aggregate or page model. Aggregate content wins.
            yield return location != null && !ReferenceEquals(location, block.Location)
                ? new OfficeDocumentBlock { Id = block.Id, Kind = block.Kind, Text = block.Text, Level = block.Level,
                    Marker = block.Marker, Region = block.Region, Location = location }
                : block;
        }
    }

    internal static IReadOnlyList<OfficeDocumentBlock> OrderBlocks(
        IEnumerable<OfficeDocumentBlock> candidates) =>
        OrderBlocks(candidates, static block => block.Location);

    internal static IReadOnlyList<OfficeDocumentBlock> OrderBlocks(
        IEnumerable<OfficeDocumentBlock> candidates,
        Func<OfficeDocumentBlock, ReaderLocation?> locationSelector) {
        if (locationSelector == null) throw new ArgumentNullException(nameof(locationSelector));
        var seen = new HashSet<OfficeDocumentBlock>(ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        var identities = new HashSet<string>(StringComparer.Ordinal);
        var ordered = new List<OfficeDocumentBlock>();
        foreach (OfficeDocumentBlock block in candidates) {
            if (block != null && seen.Add(block)) {
                // IDs and anchors survive transport round-trips; unlabelled repeated text is not a duplicate.
                if ((!string.IsNullOrWhiteSpace(block.Id) || !string.IsNullOrWhiteSpace(block.Location?.BlockAnchor))
                    && !identities.Add(BuildBlockIdentity(block, locationSelector(block)))) continue;
                ordered.Add(block);
            }
        }
        return OrderSourceItems(ordered, locationSelector);
    }

    internal static IEnumerable<ReaderTable> Tables(OfficeDocumentReadResult document) {
        var seen = new HashSet<ReaderTable>(ReferenceIdentityComparer<ReaderTable>.Instance);
        var aggregateIdentityCounts = new Dictionary<string, int>(StringComparer.Ordinal);
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

    internal static IEnumerable<OfficeDocumentFormField> Forms(
        OfficeDocumentReadResult document,
        int maxCandidates = int.MaxValue,
        Action? candidateLimitExceeded = null) {
        if (maxCandidates <= 0) yield break;
        int candidateCount = 0;
        var seen = new HashSet<OfficeDocumentFormField>(ReferenceIdentityComparer<OfficeDocumentFormField>.Instance);
        var seenIds = new HashSet<string>(System.StringComparer.Ordinal);
        foreach (OfficeDocumentFormField form in document.Forms ?? System.Array.Empty<OfficeDocumentFormField>()) {
            if (candidateCount >= maxCandidates) {
                candidateLimitExceeded?.Invoke();
                yield break;
            }
            candidateCount++;
            if (form == null) continue;
            if (seen.Add(form) &&
                (string.IsNullOrWhiteSpace(form.Id) || seenIds.Add(form.Id))) {
                yield return form;
            }
        }
        foreach (OfficeDocumentPage page in document.Pages ?? System.Array.Empty<OfficeDocumentPage>()) {
            if (page?.Forms == null) continue;
            foreach (OfficeDocumentFormField form in page.Forms) {
                if (candidateCount >= maxCandidates) {
                    candidateLimitExceeded?.Invoke();
                    yield break;
                }
                candidateCount++;
                if (form == null) continue;
                if (seen.Add(form) &&
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
                if (candidateCount >= maxCandidates) {
                    candidateLimitExceeded?.Invoke();
                    yield break;
                }
                candidateCount++;
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
            HierarchyHeadingPath = source.HierarchyHeadingPath,
            HierarchyHeadingDisplayPath = source.HierarchyHeadingDisplayPath,
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

    private static string BuildContainerOrderKey(ReaderLocation? location, IReadOnlyDictionary<string, int> sheetOrder) {
        if (location == null) return "9|";
        if (location.Page.HasValue) return "0|" + location.Page.Value.ToString("D10", CultureInfo.InvariantCulture);
        if (location.Slide.HasValue) return "1|" + location.Slide.Value.ToString("D10", CultureInfo.InvariantCulture);
        if (!string.IsNullOrWhiteSpace(location.Sheet)) return "2|" + sheetOrder[location.Sheet!].ToString("D10", CultureInfo.InvariantCulture);
        return "9|";
    }

    private static int BuildBlockPosition(ReaderLocation? location) =>
        location?.SourceBlockIndex
        ?? location?.BlockIndex
        ?? location?.StartLine
        ?? location?.NormalizedStartLine
        ?? int.MaxValue;

    private static string BuildBlockProjectionKey(OfficeDocumentBlock block) =>
        !string.IsNullOrWhiteSpace(block.Id) ? "id:" + block.Id : "anchor:" + block.Location?.BlockAnchor;

    internal static string BuildBlockIdentity(OfficeDocumentBlock block) => BuildBlockIdentity(block, block.Location);

    private static string BuildBlockIdentity(OfficeDocumentBlock block, ReaderLocation? location) {
        if (!string.IsNullOrWhiteSpace(block.Id)) return "id:" + block.Id;
        string? anchor = location?.BlockAnchor;
        if (!string.IsNullOrWhiteSpace(anchor)) {
            var builder = new StringBuilder("anchor:");
            AppendIdentity(builder, anchor);
            AppendIdentity(builder, location?.Path);
            AppendIdentity(builder, location?.Page?.ToString(CultureInfo.InvariantCulture));
            AppendIdentity(builder, location?.Slide?.ToString(CultureInfo.InvariantCulture));
            AppendIdentity(builder, location?.Sheet);
            return builder.ToString();
        }
        return BuildLocatedIdentity(block, location, block.Kind, block.Text);
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

    internal static string BuildTableIdentity(ReaderTable table, ReaderLocation? fallback = null, int? fallbackTableIndex = null, bool includeLocation = true) {
        var builder = new StringBuilder();
        AppendIdentity(builder, table.PayloadHash);
        AppendIdentity(builder, table.CallId);
        AppendIdentity(builder, table.Kind);
        AppendIdentity(builder, table.Title);
        if (includeLocation) AppendLocationIdentity(builder, table.Location, fallback, fallbackTableIndex);
        else AppendIdentity(builder, table.Location?.BlockAnchor);
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
        return WithLocationFallback(table, fallback, tableIndex);
    }

    private static ReaderTable WithLocationFallback(ReaderTable table, ReaderLocation fallback, int? tableIndex) {
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

    internal static ReaderLocation BuildPageLocation(OfficeDocumentPage page) {
        ReaderLocation source = page.Location ?? new ReaderLocation();
        var fallback = new ReaderLocation {
            Path = source.Path,
            BlockIndex = source.BlockIndex,
            SourceBlockIndex = source.SourceBlockIndex,
            StartLine = source.StartLine,
            EndLine = source.EndLine,
            NormalizedStartLine = source.NormalizedStartLine,
            NormalizedEndLine = source.NormalizedEndLine,
            HeadingPath = source.HeadingPath,
            HierarchyHeadingPath = source.HierarchyHeadingPath,
            HierarchyHeadingDisplayPath = source.HierarchyHeadingDisplayPath,
            HeadingSlug = source.HeadingSlug,
            SourceBlockKind = source.SourceBlockKind,
            BlockAnchor = source.BlockAnchor,
            Sheet = source.Sheet,
            A1Range = source.A1Range,
            Slide = source.Slide,
            Page = source.Page,
            TableIndex = source.TableIndex
        };
        // A sheet or slide number denotes that container, not a PDF-style page number.
        fallback.Page = page.Location?.Page;
        string? kind = page.Location?.SourceBlockKind?.Trim();
        if (string.Equals(kind, "sheet", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(fallback.Sheet))
            fallback.Sheet = !string.IsNullOrWhiteSpace(page.Name) ? page.Name
                : page.Number > 0 ? "Sheet " + page.Number.Value.ToString(CultureInfo.InvariantCulture) : null;
        if (!fallback.Slide.HasValue && string.IsNullOrWhiteSpace(fallback.Sheet)) {
            int? number = page.Number > 0 ? page.Number : fallback.Page;
            if (string.Equals(kind, "slide", StringComparison.OrdinalIgnoreCase)) {
                fallback.Slide = number;
                fallback.Page = null;
            } else {
                fallback.Page = number;
            }
        }
        return fallback;
    }

    private static bool NeedsLocationFallback(ReaderLocation location) {
        return string.IsNullOrWhiteSpace(location.Path)
            || (!location.Page.HasValue && !location.Slide.HasValue && string.IsNullOrWhiteSpace(location.Sheet));
    }

    private static ReaderLocation MergeLocation(ReaderLocation? location, ReaderLocation fallback, int? fallbackTableIndex) {
        return new ReaderLocation {
            Path = Prefer(location?.Path, fallback.Path),
            BlockIndex = location?.BlockIndex ?? fallback.BlockIndex,
            SourceBlockIndex = location?.SourceBlockIndex ?? fallback.SourceBlockIndex,
            StartLine = location?.StartLine ?? fallback.StartLine,
            EndLine = location?.EndLine ?? fallback.EndLine,
            NormalizedStartLine = location?.NormalizedStartLine ?? fallback.NormalizedStartLine,
            NormalizedEndLine = location?.NormalizedEndLine ?? fallback.NormalizedEndLine,
            HeadingPath = Prefer(location?.HeadingPath, fallback.HeadingPath),
            HierarchyHeadingPath = Prefer(location?.HierarchyHeadingPath, fallback.HierarchyHeadingPath),
            HierarchyHeadingDisplayPath = string.IsNullOrWhiteSpace(location?.HierarchyHeadingPath)
                ? fallback.HierarchyHeadingDisplayPath : location!.HierarchyHeadingDisplayPath,
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
        AppendIdentity(builder, Prefer(location?.HierarchyHeadingPath, fallback?.HierarchyHeadingPath));
        AppendIdentity(builder, Prefer(location?.HeadingSlug, fallback?.HeadingSlug));
        AppendIdentity(builder, Prefer(location?.SourceBlockKind, fallback?.SourceBlockKind));
        AppendIdentity(builder, Prefer(location?.BlockAnchor, fallback?.BlockAnchor));
        AppendIdentity(builder, (location?.Page ?? fallback?.Page)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.Slide ?? fallback?.Slide)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.BlockIndex ?? fallback?.BlockIndex)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.SourceBlockIndex ?? fallback?.SourceBlockIndex)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.StartLine ?? fallback?.StartLine)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.EndLine ?? fallback?.EndLine)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.NormalizedStartLine ?? fallback?.NormalizedStartLine)?.ToString(CultureInfo.InvariantCulture));
        AppendIdentity(builder, (location?.NormalizedEndLine ?? fallback?.NormalizedEndLine)?.ToString(CultureInfo.InvariantCulture));
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

}

internal sealed class ReferenceIdentityComparer<T> : IEqualityComparer<T> where T : class {
    internal static ReferenceIdentityComparer<T> Instance { get; } = new ReferenceIdentityComparer<T>();

    public bool Equals(T? x, T? y) => ReferenceEquals(x, y);

    public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
}
