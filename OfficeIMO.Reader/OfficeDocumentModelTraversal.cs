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
}

internal sealed class ReferenceIdentityComparer<T> : IEqualityComparer<T> where T : class {
    internal static ReferenceIdentityComparer<T> Instance { get; } = new ReferenceIdentityComparer<T>();

    public bool Equals(T? x, T? y) => ReferenceEquals(x, y);

    public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
}
