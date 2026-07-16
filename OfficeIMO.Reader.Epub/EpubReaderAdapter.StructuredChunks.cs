using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;

namespace OfficeIMO.Reader.Epub;

internal static partial class EpubReaderAdapter {
    /// <summary>
    /// Projects chapter XHTML through the shared HTML adapter while retaining EPUB source identity.
    /// </summary>
    private static IReadOnlyList<ReaderChunk> ReadStructuredChapter(
        EpubChapter chapter,
        SourceMetadata source,
        ReaderOptions options,
        int firstBlockIndex,
        CancellationToken cancellationToken) {
        string virtualPath = BuildVirtualPath(source.Path, chapter.Path);
        string fileName = Path.GetFileName(source.Path);
        var chunks = new List<ReaderChunk>();
        foreach (ReaderChunk htmlChunk in HtmlReaderAdapter.ReadContent(
                     chapter.Html!,
                     virtualPath,
                     options,
                     cancellationToken: cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            int chunkPart = chunks.Count;
            string markdown = htmlChunk.Markdown ?? htmlChunk.Text;
            if (chunkPart == 0) {
                markdown = BuildMarkdown(chapter.Title, markdown);
            }

            var chunk = new ReaderChunk {
                Id = BuildId(fileName, chapter.Order, chunkPart),
                Kind = ReaderInputKind.Epub,
                Location = MapStructuredChapterLocation(
                    htmlChunk.Location,
                    virtualPath,
                    firstBlockIndex + chunkPart),
                Text = htmlChunk.Text,
                Markdown = markdown,
                Tables = htmlChunk.Tables,
                Visuals = htmlChunk.Visuals,
                FormFields = htmlChunk.FormFields,
                Actions = htmlChunk.Actions,
                Diagnostics = htmlChunk.Diagnostics,
                Warnings = htmlChunk.Warnings
            };
            EnrichChunk(chunk, source, options.ComputeHashes);
            ApplyVirtualSourceMetadata(chunk, virtualPath, options.ComputeHashes);
            chunks.Add(chunk);
        }
        return chunks;
    }

    private static ReaderLocation MapStructuredChapterLocation(
        ReaderLocation source,
        string virtualPath,
        int blockIndex) {
        var location = new ReaderLocation {
            Path = virtualPath,
            BlockIndex = blockIndex,
            SourceBlockIndex = source.SourceBlockIndex ?? source.BlockIndex,
            StartLine = source.StartLine,
            EndLine = source.EndLine,
            NormalizedStartLine = source.NormalizedStartLine,
            NormalizedEndLine = source.NormalizedEndLine,
            HeadingPath = source.HeadingPath,
            HeadingSlug = source.HeadingSlug,
            SourceBlockKind = source.SourceBlockKind,
            BlockAnchor = source.BlockAnchor,
            TableIndex = source.TableIndex
        };
        ReaderHeadingPath.SetHierarchyPath(location, source.HierarchyHeadingPath);
        return location;
    }
}
