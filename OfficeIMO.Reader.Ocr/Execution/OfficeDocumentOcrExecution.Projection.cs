namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrExecutionExtensions {
    private static void AppendNestedOcrProjection(OfficeDocumentReadResult parent, OfficeDocumentReadResult before,
        OfficeDocumentReadResult after, string documentId, string? documentPath, bool appendMarkdown,
        CancellationToken cancellationToken) {
        var blocks = new List<OfficeDocumentBlock>();
        var chunks = new List<ReaderChunk>();
        foreach (OfficeDocumentBlock block in after.Blocks.Skip(before.Blocks?.Count ?? 0)) {
            cancellationToken.ThrowIfCancellationRequested();
            string id = ProjectId(documentId, block.Id);
            blocks.Add(new OfficeDocumentBlock {
                Id = id, Kind = block.Kind, Text = block.Text, Level = block.Level, Marker = block.Marker,
                Region = block.Region, Location = ProjectLocation(block.Location, before.Source?.Path, documentPath, ProjectId(documentId, block.Location.BlockAnchor ?? block.Id))
            });
        }
        foreach (ReaderChunk chunk in after.Chunks.Skip(before.Chunks?.Count ?? 0)) {
            cancellationToken.ThrowIfCancellationRequested();
            string id = ProjectId(documentId, chunk.Id);
            ReaderChunk projected = chunk.CopyForContainer();
            projected.Id = id;
            projected.Location = ProjectLocation(chunk.Location, before.Source?.Path, documentPath,
                ProjectId(documentId, chunk.Location.BlockAnchor ?? chunk.Id));
            projected.ChunkHash = null; // A new container identity invalidates any chunk hash.
            chunks.Add(projected);
        }
        parent.Blocks = parent.Blocks.Concat(blocks).ToArray();
        parent.Chunks = parent.Chunks.Concat(chunks).ToArray();
        if (appendMarkdown && blocks.Count > 0) {
            string text = string.Join("\n\n", blocks.Select(block => block.Text));
            parent.Markdown = string.IsNullOrWhiteSpace(parent.Markdown) ? text : parent.Markdown!.TrimEnd() + "\n\n" + text;
        }
    }

    private static string ProjectId(string documentId, string id) =>
        id.StartsWith(documentId + "/", StringComparison.Ordinal) ? id : documentId + "/" + id;

    private static ReaderLocation ProjectLocation(ReaderLocation? input, string? originalPath, string? documentPath, string anchor) {
        input ??= new ReaderLocation();
        string? path = input.Path;
        if (string.IsNullOrWhiteSpace(path) || path == originalPath) path = documentPath;
        else if (!string.IsNullOrWhiteSpace(originalPath) && IsNestedPath(originalPath!, path!))
            path = documentPath + path.Substring(originalPath!.Length);
        else if (documentPath != null && path != documentPath && !IsNestedPath(documentPath, path!)
            && !System.IO.Path.IsPathRooted(path)) path = documentPath + "!/" + path;
        var result = new ReaderLocation {
            Path = path, BlockIndex = input.BlockIndex, SourceBlockIndex = input.SourceBlockIndex,
            StartLine = input.StartLine, EndLine = input.EndLine,
            NormalizedStartLine = input.NormalizedStartLine, NormalizedEndLine = input.NormalizedEndLine,
            HeadingPath = input.HeadingPath, HeadingSlug = input.HeadingSlug,
            SourceBlockKind = input.SourceBlockKind, BlockAnchor = anchor,
            Sheet = input.Sheet, A1Range = input.A1Range, Slide = input.Slide, Page = input.Page, TableIndex = input.TableIndex
        };
        ReaderHeadingPath.SetHierarchyPath(result, ReaderHeadingPath.GetValidatedHierarchyPath(input));
        return result;
    }
}
