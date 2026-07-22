using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using System.Linq;

namespace OfficeIMO.Reader.Epub;

internal static partial class EpubReaderAdapter {
    /// <summary>Reads an EPUB file into the shared rich document envelope.</summary>
    public static OfficeDocumentReadResult ReadDocument(string epubPath, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, CancellationToken cancellationToken = default) {
        if (epubPath == null) throw new ArgumentNullException(nameof(epubPath));
        if (epubPath.Length == 0) throw new ArgumentException("EPUB path cannot be empty.", nameof(epubPath));
        if (!File.Exists(epubPath)) throw new FileNotFoundException($"EPUB file '{epubPath}' doesn't exist.", epubPath);
        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(epubPath, effective.MaxInputBytes);
        SourceMetadata source = BuildSourceMetadataFromPath(epubPath, effective.ComputeHashes);
        EpubDocument document = EpubDocument.Load(epubPath, CreateRichOptions(epubOptions));
        return BuildEpubDocumentResult(document, source, effective, cancellationToken);
    }

    /// <summary>Reads an EPUB stream into the shared rich document envelope.</summary>
    public static OfficeDocumentReadResult ReadDocument(Stream epubStream, string? sourceName = null, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, CancellationToken cancellationToken = default) {
        if (epubStream == null) throw new ArgumentNullException(nameof(epubStream));
        if (!epubStream.CanRead) throw new ArgumentException("EPUB stream must be readable.", nameof(epubStream));
        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        string logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "document.epub" : sourceName!.Trim();
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(epubStream, effective.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            SourceMetadata source = BuildSourceMetadataFromStream(parseStream, logicalSourceName, effective.ComputeHashes);
            if (parseStream.CanSeek) parseStream.Position = 0;
            EpubDocument document = EpubDocument.Load(parseStream, CreateRichOptions(epubOptions));
            return BuildEpubDocumentResult(document, source, effective, cancellationToken);
        } finally {
            if (ownsParseStream) parseStream.Dispose();
        }
    }

    /// <summary>Reads an EPUB file into the shared rich document JSON envelope.</summary>
    public static string ReadDocumentJson(string epubPath, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(epubPath, readerOptions, epubOptions, cancellationToken), indented);
    }

    /// <summary>Reads an EPUB stream into the shared rich document JSON envelope.</summary>
    public static string ReadDocumentJson(Stream epubStream, string? sourceName = null, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(epubStream, sourceName, readerOptions, epubOptions, cancellationToken), indented);
    }

    private static OfficeDocumentReadResult BuildEpubDocumentResult(EpubDocument document, SourceMetadata source, ReaderOptions readerOptions, CancellationToken cancellationToken) {
        ReaderChunk[] chunks = ReadDocument(document, source, readerOptions, cancellationToken).ToArray();
        List<OfficeDocumentAsset> assets = BuildEpubAssets(document, source.Path).ToList();
        Dictionary<string, OfficeDocumentAsset> assetsByLocation = BuildEpubAssetIndex(assets);
        var blocks = new List<OfficeDocumentBlock>();
        var tables = new List<ReaderTable>();
        var links = new List<OfficeDocumentLink>();
        var forms = new List<OfficeDocumentFormField>();
        var visuals = new List<ReaderVisual>();
        var diagnostics = new List<OfficeDocumentDiagnostic>(BuildEpubDiagnostics(document, source.Path));
        var pages = new List<OfficeDocumentPage>();
        var capabilities = new List<string> {
            "officeimo.reader.epub.rich-v5",
            "officeimo.epub.package-model",
            "officeimo.html.logical-document"
        };
        var seenCapabilities = new HashSet<string>(capabilities, StringComparer.Ordinal);
        int tableIndex = 0;
        for (int chapterIndex = 0; chapterIndex < document.Chapters.Count; chapterIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            EpubChapter chapter = document.Chapters[chapterIndex];
            string virtualPath = BuildVirtualPath(source.Path, chapter.Path);
            var chapterBlocks = new List<OfficeDocumentBlock>();
            var chapterTables = new List<ReaderTable>();
            var chapterLinks = new List<OfficeDocumentLink>();
            var chapterForms = new List<OfficeDocumentFormField>();
            var chapterAssets = new List<OfficeDocumentAsset>();
            if (!string.IsNullOrWhiteSpace(chapter.Html)) {
                OfficeDocumentReadResult htmlResult = HtmlReaderAdapter.ReadContentDocument(
                    chapter.Html!,
                    virtualPath,
                    readerOptions,
                    cancellationToken: cancellationToken);
                foreach (string capability in htmlResult.CapabilitiesUsed) {
                    if (capability.StartsWith("officeimo.html.", StringComparison.Ordinal)
                        && seenCapabilities.Add(capability)) {
                        capabilities.Add(capability);
                    }
                }
                IReadOnlyList<OfficeDocumentLink> resolvedLinks = ResolveEpubChapterLinks(
                    source.Path,
                    chapter,
                    htmlResult.Links,
                    diagnostics);
                string prefix = "epub-chapter-" + (chapterIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-";
                PrefixHtmlProjection(prefix, htmlResult, ref tableIndex);
                chapterBlocks.AddRange(htmlResult.Blocks);
                chapterTables.AddRange(htmlResult.Tables);
                chapterLinks.AddRange(resolvedLinks);
                chapterForms.AddRange(htmlResult.Forms);
                AddEpubChapterAssets(source.Path, chapter, htmlResult.Assets, assets, assetsByLocation, chapterAssets, diagnostics);
                IReadOnlyList<ReaderVisual> resolvedVisuals = ResolveEpubChapterVisuals(
                    source.Path,
                    chapter,
                    htmlResult.Visuals,
                    assetsByLocation,
                    chapterAssets,
                    diagnostics);
                visuals.AddRange(resolvedVisuals);
                diagnostics.AddRange(htmlResult.Diagnostics.Where(static diagnostic =>
                    !string.Equals(diagnostic.Code, "ocr-needed", StringComparison.Ordinal)));
            }
            if (chapterBlocks.Count == 0) {
                string anchor = "epub-chapter-" + (chapterIndex + 1).ToString("D4", CultureInfo.InvariantCulture);
                chapterBlocks.Add(new OfficeDocumentBlock {
                    Id = anchor,
                    Kind = "chapter",
                    Text = chapter.Text,
                    Level = 1,
                    Location = new ReaderLocation { Path = virtualPath, SourceBlockIndex = chapterIndex, SourceBlockKind = "chapter", BlockAnchor = anchor }
                });
            }
            blocks.AddRange(chapterBlocks);
            tables.AddRange(chapterTables);
            links.AddRange(chapterLinks);
            forms.AddRange(chapterForms);
            pages.Add(new OfficeDocumentPage {
                Number = chapter.Order,
                Name = chapter.Title,
                Location = new ReaderLocation { Path = virtualPath, SourceBlockIndex = chapterIndex, SourceBlockKind = "chapter", BlockAnchor = "epub-chapter-" + (chapterIndex + 1).ToString("D4", CultureInfo.InvariantCulture) },
                Blocks = chapterBlocks,
                Tables = chapterTables,
                Assets = chapterAssets,
                Links = chapterLinks,
                Forms = chapterForms
            });
        }

        var documentSource = new OfficeDocumentSource {
            Path = source.Path,
            SourceId = source.SourceId,
            SourceHash = source.SourceHash,
            LastWriteUtc = source.LastWriteUtc,
            LengthBytes = source.LengthBytes,
            Title = document.Title,
            Author = document.Creator
        };
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Epub,
            documentSource,
            capabilities,
            assets);
        result.Blocks = blocks;
        result.Tables = tables;
        result.Links = links;
        result.Forms = forms;
        result.Visuals = visuals;
        result.Pages = AttachEpubOcrCandidates(pages, result.OcrCandidates);
        var packageWarningMessages = new HashSet<string>(
            document.Diagnostics
                .Where(static diagnostic => diagnostic.Severity != EpubDiagnosticSeverity.Info)
                .Select(static diagnostic => diagnostic.Message),
            StringComparer.Ordinal);
        result.Diagnostics = result.Diagnostics
            .Where(diagnostic =>
                !string.Equals(diagnostic.Code, "reader-warning", StringComparison.Ordinal) ||
                !packageWarningMessages.Contains(diagnostic.Message))
            .Concat(diagnostics)
            .ToArray();
        result.Metadata = BuildEpubMetadata(document, source.Path, blocks.Count, tables.Count, links.Count, assets.Count);
        return result;
    }

    private static IReadOnlyList<OfficeDocumentPage> AttachEpubOcrCandidates(
        IReadOnlyList<OfficeDocumentPage> pages,
        IReadOnlyList<OfficeDocumentOcrCandidate> candidates) {
        foreach (OfficeDocumentPage page in pages) {
            var pageAssetIds = new HashSet<string>(
                page.Assets.Where(static asset => !string.IsNullOrWhiteSpace(asset.Id)).Select(static asset => asset.Id),
                StringComparer.Ordinal);
            page.OcrCandidates = candidates
                .Where(candidate => !string.IsNullOrWhiteSpace(candidate.AssetId) && pageAssetIds.Contains(candidate.AssetId!))
                .ToArray();
        }
        return pages;
    }

    private static void PrefixHtmlProjection(string prefix, OfficeDocumentReadResult result, ref int tableIndex) {
        foreach (OfficeDocumentBlock block in result.Blocks) {
            block.Id = prefix + block.Id;
            if (!string.IsNullOrWhiteSpace(block.Location.BlockAnchor)) block.Location.BlockAnchor = prefix + block.Location.BlockAnchor;
        }
        foreach (ReaderTable table in result.Tables) {
            if (table.Location != null) {
                table.Location.TableIndex = tableIndex++;
                if (!string.IsNullOrWhiteSpace(table.Location.BlockAnchor)) table.Location.BlockAnchor = prefix + table.Location.BlockAnchor;
            }
        }
        foreach (OfficeDocumentLink link in result.Links) {
            link.Id = prefix + link.Id;
            if (!string.IsNullOrWhiteSpace(link.Location.BlockAnchor)) link.Location.BlockAnchor = prefix + link.Location.BlockAnchor;
        }
        foreach (OfficeDocumentFormField form in result.Forms) {
            form.Id = prefix + form.Id;
            if (!string.IsNullOrWhiteSpace(form.Location.BlockAnchor)) form.Location.BlockAnchor = prefix + form.Location.BlockAnchor;
        }
        foreach (OfficeDocumentAsset asset in result.Assets) {
            asset.Id = prefix + asset.Id;
            string? extension = string.IsNullOrWhiteSpace(asset.Extension)
                ? Path.GetExtension(asset.FileName)
                : asset.Extension;
            asset.FileName = OfficeDocumentAssetNaming.BuildFileName(asset.Id, extension);
            if (!string.IsNullOrWhiteSpace(asset.Location.BlockAnchor)) asset.Location.BlockAnchor = prefix + asset.Location.BlockAnchor;
        }
        foreach (ReaderVisual visual in result.Visuals) {
            if (visual.Location != null && !string.IsNullOrWhiteSpace(visual.Location.BlockAnchor)) visual.Location.BlockAnchor = prefix + visual.Location.BlockAnchor;
        }
    }

}
