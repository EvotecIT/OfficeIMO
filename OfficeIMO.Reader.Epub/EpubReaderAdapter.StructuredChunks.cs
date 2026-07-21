using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using System.Linq;

namespace OfficeIMO.Reader.Epub;

internal static partial class EpubReaderAdapter {
    private const string HtmlNoMarkdownWarning = "HTML content produced no markdown text.";

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
        var nonContentWarnings = new List<string>();
        ReaderHtmlOptions htmlOptions = ReaderHtmlOptions.CreateOfficeIMOProfile();
        ConfigureEpubMarkdownReferencePolicies(htmlOptions, source.Path, chapter);
        foreach (ReaderChunk htmlChunk in HtmlReaderAdapter.ReadContent(
                     chapter.Html!,
                     virtualPath,
                     options,
                     htmlOptions,
                     cancellationToken: cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsHtmlNoMarkdownWarningChunk(htmlChunk)) {
                if (htmlChunk.Warnings == null || htmlChunk.Warnings.Count == 0) {
                    nonContentWarnings.Add(HtmlNoMarkdownWarning);
                } else {
                    nonContentWarnings.AddRange(htmlChunk.Warnings);
                }
                continue;
            }
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
                    firstBlockIndex + chunkPart,
                    chapter.Order > 0 ? chapter.Order - 1 : null,
                    chapter.Title),
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
        if (chunks.Count == 0 && nonContentWarnings.Count > 0) {
            chunks.Add(BuildStructuredOnlyChapterChunk(
                chapter,
                source,
                options,
                firstBlockIndex,
                nonContentWarnings));
        }
        return chunks;
    }

    private static void ConfigureEpubMarkdownReferencePolicies(
        ReaderHtmlOptions htmlOptions,
        string sourcePath,
        EpubChapter chapter) {
        HtmlToMarkdownOptions? markdownOptions = htmlOptions.HtmlToMarkdownOptions;
        if (markdownOptions == null) return;

        Func<string, string?> transform = value => ResolveEpubMarkdownReference(sourcePath, chapter, value);
        markdownOptions.UrlPolicy.ResolvedUrlTransform = transform;
        HtmlUrlPolicy resourcePolicy = (markdownOptions.ResourceUrlPolicy ?? markdownOptions.UrlPolicy).Clone();
        resourcePolicy.ResolvedUrlTransform = transform;
        markdownOptions.ResourceUrlPolicy = resourcePolicy;

        HtmlConversionDocumentOptions conversionOptions = HtmlConversionDocumentOptions.CreateUntrustedProfile();
        AllowEpubVirtualFileLocations(conversionOptions.UrlPolicy);
        AllowEpubVirtualFileLocations(conversionOptions.ResourceUrlPolicy);
        htmlOptions.ConversionOptions = conversionOptions;
    }

    private static void AllowEpubVirtualFileLocations(HtmlUrlPolicy policy) {
        policy.DisallowFileUrls = false;
        if (policy.RestrictUrlSchemes) policy.AllowedUrlSchemes.Add(Uri.UriSchemeFile);
    }

    private static bool IsHtmlNoMarkdownWarningChunk(ReaderChunk chunk) {
        return string.Equals(chunk.Id, "html-warning-0000", StringComparison.Ordinal) &&
               chunk.Warnings?.Any(warning => string.Equals(warning, HtmlNoMarkdownWarning, StringComparison.Ordinal)) == true;
    }

    private static ReaderChunk BuildStructuredOnlyChapterChunk(
        EpubChapter chapter,
        SourceMetadata source,
        ReaderOptions options,
        int blockIndex,
        IEnumerable<string>? warnings) {
        string virtualPath = BuildVirtualPath(source.Path, chapter.Path);
        string? displayHeading = string.IsNullOrWhiteSpace(chapter.Title) ? null : chapter.Title!.Trim();
        var location = new ReaderLocation {
            Path = virtualPath,
            BlockIndex = blockIndex,
            SourceBlockIndex = chapter.Order > 0 ? chapter.Order - 1 : null,
            HeadingPath = displayHeading,
            SourceBlockKind = "chapter"
        };
        ReaderHeadingPath.SetHierarchyPath(location, ReaderHeadingPath.Combine(new[] { displayHeading }));
        var chunk = new ReaderChunk {
            Id = BuildId(Path.GetFileName(source.Path), chapter.Order, 0),
            Kind = ReaderInputKind.Epub,
            Location = location,
            Text = string.Empty,
            Markdown = BuildMarkdown(chapter.Title, string.Empty),
            Warnings = warnings?.Distinct(StringComparer.Ordinal).ToArray()
        };
        EnrichChunk(chunk, source, options.ComputeHashes);
        ApplyVirtualSourceMetadata(chunk, virtualPath, options.ComputeHashes);
        return chunk;
    }

    private static ReaderLocation MapStructuredChapterLocation(
        ReaderLocation source,
        string virtualPath,
        int blockIndex,
        int? sourceBlockIndex,
        string? chapterTitle) {
        string? displayHeading = string.IsNullOrWhiteSpace(chapterTitle)
            ? source.HeadingPath
            : chapterTitle!.Trim();
        var location = new ReaderLocation {
            Path = virtualPath,
            BlockIndex = blockIndex,
            SourceBlockIndex = sourceBlockIndex,
            StartLine = source.StartLine,
            EndLine = source.EndLine,
            NormalizedStartLine = source.NormalizedStartLine,
            NormalizedEndLine = source.NormalizedEndLine,
            HeadingPath = displayHeading,
            HeadingSlug = source.HeadingSlug,
            SourceBlockKind = source.SourceBlockKind,
            BlockAnchor = source.BlockAnchor,
            TableIndex = source.TableIndex
        };
        var hierarchy = new List<string?> { displayHeading };
        if (!string.IsNullOrWhiteSpace(source.HierarchyHeadingPath)) {
            hierarchy.AddRange(ReaderHeadingPath.Split(source.HierarchyHeadingPath));
        } else if (!string.Equals(source.HeadingPath, displayHeading, StringComparison.Ordinal)) {
            hierarchy.Add(source.HeadingPath);
        }
        ReaderHeadingPath.SetHierarchyPath(location, ReaderHeadingPath.Combine(hierarchy));
        return location;
    }
}
