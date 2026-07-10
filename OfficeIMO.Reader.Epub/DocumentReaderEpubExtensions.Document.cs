using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using System.Linq;

namespace OfficeIMO.Reader.Epub;

public static partial class DocumentReaderEpubExtensions {
    /// <summary>Reads an EPUB file into the shared rich document envelope.</summary>
    public static OfficeDocumentReadResult ReadEpubDocument(string epubPath, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, CancellationToken cancellationToken = default) {
        if (epubPath == null) throw new ArgumentNullException(nameof(epubPath));
        if (epubPath.Length == 0) throw new ArgumentException("EPUB path cannot be empty.", nameof(epubPath));
        if (!File.Exists(epubPath)) throw new FileNotFoundException($"EPUB file '{epubPath}' doesn't exist.", epubPath);
        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(epubPath, effective.MaxInputBytes);
        SourceMetadata source = BuildSourceMetadataFromPath(epubPath, effective.ComputeHashes);
        EpubDocument document = EpubReader.Read(epubPath, CreateRichOptions(epubOptions));
        return BuildEpubDocumentResult(document, source, effective, cancellationToken);
    }

    /// <summary>Reads an EPUB stream into the shared rich document envelope.</summary>
    public static OfficeDocumentReadResult ReadEpubDocument(Stream epubStream, string? sourceName = null, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, CancellationToken cancellationToken = default) {
        if (epubStream == null) throw new ArgumentNullException(nameof(epubStream));
        if (!epubStream.CanRead) throw new ArgumentException("EPUB stream must be readable.", nameof(epubStream));
        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        string logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "document.epub" : sourceName!.Trim();
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(epubStream, effective.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            SourceMetadata source = BuildSourceMetadataFromStream(parseStream, logicalSourceName, effective.ComputeHashes);
            if (parseStream.CanSeek) parseStream.Position = 0;
            EpubDocument document = EpubReader.Read(parseStream, CreateRichOptions(epubOptions));
            return BuildEpubDocumentResult(document, source, effective, cancellationToken);
        } finally {
            if (ownsParseStream) parseStream.Dispose();
        }
    }

    /// <summary>Reads an EPUB file into the shared rich document JSON envelope.</summary>
    public static string ReadEpubDocumentJson(string epubPath, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadEpubDocument(epubPath, readerOptions, epubOptions, cancellationToken), indented);
    }

    /// <summary>Reads an EPUB stream into the shared rich document JSON envelope.</summary>
    public static string ReadEpubDocumentJson(Stream epubStream, string? sourceName = null, ReaderOptions? readerOptions = null, EpubReadOptions? epubOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadEpubDocument(epubStream, sourceName, readerOptions, epubOptions, cancellationToken), indented);
    }

    private static OfficeDocumentReadResult BuildEpubDocumentResult(EpubDocument document, SourceMetadata source, ReaderOptions readerOptions, CancellationToken cancellationToken) {
        ReaderChunk[] chunks = ReadEpubDocument(document, source, readerOptions, cancellationToken).ToArray();
        List<OfficeDocumentAsset> assets = BuildEpubAssets(document, source.Path).ToList();
        var blocks = new List<OfficeDocumentBlock>();
        var tables = new List<ReaderTable>();
        var links = new List<OfficeDocumentLink>();
        var forms = new List<OfficeDocumentFormField>();
        var visuals = new List<ReaderVisual>();
        var diagnostics = new List<OfficeDocumentDiagnostic>();
        var pages = new List<OfficeDocumentPage>();
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
                OfficeDocumentReadResult htmlResult = DocumentReaderHtmlExtensions.ReadHtmlStringDocument(
                    chapter.Html!,
                    virtualPath,
                    readerOptions,
                    cancellationToken: cancellationToken);
                string prefix = "epub-chapter-" + (chapterIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-";
                PrefixHtmlProjection(prefix, htmlResult, ref tableIndex);
                chapterBlocks.AddRange(htmlResult.Blocks);
                chapterTables.AddRange(htmlResult.Tables);
                chapterLinks.AddRange(htmlResult.Links);
                chapterForms.AddRange(htmlResult.Forms);
                visuals.AddRange(htmlResult.Visuals);
                AddEpubChapterAssets(source.Path, chapter.Path, htmlResult.Assets, assets, chapterAssets);
                diagnostics.AddRange(htmlResult.Diagnostics);
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
        OfficeDocumentReadResult result = DocumentReader.CreateDocumentResult(
            chunks,
            ReaderInputKind.Epub,
            documentSource,
            new[] { "officeimo.reader.epub.rich-v5", "officeimo.epub.package-model", "officeimo.html.logical-document" },
            assets);
        result.Blocks = blocks;
        result.Tables = tables;
        result.Links = links;
        result.Forms = forms;
        result.Visuals = visuals;
        result.Pages = pages;
        result.Diagnostics = result.Diagnostics.Concat(diagnostics).ToArray();
        result.Metadata = BuildEpubMetadata(document, blocks.Count, tables.Count, links.Count, assets.Count);
        return result;
    }

    private static EpubReadOptions CreateRichOptions(EpubReadOptions? options) {
        EpubReadOptions source = options ?? new EpubReadOptions();
        return new EpubReadOptions {
            MaxChapters = source.MaxChapters,
            MaxChapterBytes = source.MaxChapterBytes,
            IncludeRawHtml = true,
            IncludeResourceData = options == null || source.IncludeResourceData,
            MaxResources = source.MaxResources,
            MaxResourceBytes = source.MaxResourceBytes,
            MaxTotalResourceBytes = source.MaxTotalResourceBytes,
            DeterministicOrder = source.DeterministicOrder,
            PreferSpineOrder = source.PreferSpineOrder,
            IncludeNonLinearSpineItems = source.IncludeNonLinearSpineItems,
            FallbackToHtmlScan = source.FallbackToHtmlScan
        };
    }

    private static IEnumerable<OfficeDocumentAsset> BuildEpubAssets(EpubDocument document, string sourcePath) {
        int assetIndex = 0;
        foreach (EpubResource resource in document.Resources) {
            if (string.IsNullOrWhiteSpace(resource.MediaType) || !resource.MediaType!.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) continue;
            string id = "epub-image-" + assetIndex.ToString("D4", CultureInfo.InvariantCulture);
            string extension = Path.GetExtension(resource.Path);
            yield return new OfficeDocumentAsset {
                Id = id,
                Kind = "image",
                MediaType = resource.MediaType,
                Extension = string.IsNullOrWhiteSpace(extension) ? null : extension,
                FileName = OfficeDocumentAssetNaming.BuildFileName(id, extension),
                LengthBytes = resource.LengthBytes,
                PayloadHash = resource.Data == null ? null : ComputeEpubPayloadHash(resource.Data),
                PayloadBytes = resource.Data,
                SourceObjectId = resource.Id,
                Location = new ReaderLocation {
                    Path = BuildVirtualPath(sourcePath, resource.Path),
                    SourceBlockKind = "image",
                    BlockAnchor = id
                }
            };
            assetIndex++;
        }
    }

    private static void AddEpubChapterAssets(
        string sourcePath,
        string chapterPath,
        IReadOnlyList<OfficeDocumentAsset> htmlAssets,
        List<OfficeDocumentAsset> documentAssets,
        List<OfficeDocumentAsset> chapterAssets) {
        foreach (OfficeDocumentAsset htmlAsset in htmlAssets) {
            OfficeDocumentAsset? mappedAsset = null;
            if (htmlAsset.PayloadBytes == null) {
                string? resourcePath = ResolveEpubResourcePath(chapterPath, htmlAsset.SourceObjectId);
                if (!string.IsNullOrWhiteSpace(resourcePath)) {
                    string virtualResourcePath = BuildVirtualPath(sourcePath, resourcePath!);
                    mappedAsset = documentAssets.FirstOrDefault(asset => string.Equals(asset.Location.Path, virtualResourcePath, StringComparison.Ordinal));
                }
            }

            if (mappedAsset == null) {
                mappedAsset = htmlAsset;
                documentAssets.Add(mappedAsset);
            }
            if (!chapterAssets.Contains(mappedAsset)) chapterAssets.Add(mappedAsset);
        }
    }

    private static string? ResolveEpubResourcePath(string chapterPath, string? sourceObjectId) {
        if (string.IsNullOrWhiteSpace(sourceObjectId)) return null;
        string candidate = sourceObjectId!.Trim();
        int fragmentIndex = candidate.IndexOfAny(new[] { '#', '?' });
        if (fragmentIndex >= 0) candidate = candidate.Substring(0, fragmentIndex);
        if (candidate.Length == 0
            || candidate.StartsWith("data:", StringComparison.OrdinalIgnoreCase)
            || Uri.TryCreate(candidate, UriKind.Absolute, out _)) {
            return null;
        }

        try {
            candidate = Uri.UnescapeDataString(candidate).Replace('\\', '/');
        } catch {
            candidate = candidate.Replace('\\', '/');
        }

        string combined;
        if (candidate.StartsWith("/", StringComparison.Ordinal)) {
            combined = candidate.TrimStart('/');
        } else {
            int lastSlash = chapterPath.LastIndexOf('/');
            string chapterDirectory = lastSlash < 0 ? string.Empty : chapterPath.Substring(0, lastSlash + 1);
            combined = chapterDirectory + candidate;
        }

        var segments = new List<string>();
        foreach (string segment in combined.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (segment == ".") continue;
            if (segment == "..") {
                if (segments.Count == 0) return null;
                segments.RemoveAt(segments.Count - 1);
            } else {
                segments.Add(segment);
            }
        }
        return segments.Count == 0 ? null : string.Join("/", segments);
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

    private static IReadOnlyList<OfficeDocumentMetadataEntry> BuildEpubMetadata(EpubDocument document, int blockCount, int tableCount, int linkCount, int assetCount) {
        var metadata = new List<OfficeDocumentMetadataEntry> {
            EpubMetadata("epub-chapter-count", "ChapterCount", document.Chapters.Count, "count"),
            EpubMetadata("epub-resource-count", "ResourceCount", document.Resources.Count, "count"),
            EpubMetadata("epub-block-count", "BlockCount", blockCount, "count"),
            EpubMetadata("epub-table-count", "TableCount", tableCount, "count"),
            EpubMetadata("epub-link-count", "LinkCount", linkCount, "count"),
            EpubMetadata("epub-asset-count", "AssetCount", assetCount, "count")
        };
        if (!string.IsNullOrWhiteSpace(document.Identifier)) metadata.Add(EpubMetadata("epub-identifier", "Identifier", document.Identifier!, "string"));
        if (!string.IsNullOrWhiteSpace(document.Language)) metadata.Add(EpubMetadata("epub-language", "Language", document.Language!, "string"));
        if (!string.IsNullOrWhiteSpace(document.OpfPath)) metadata.Add(EpubMetadata("epub-package-path", "PackagePath", document.OpfPath!, "string"));
        return metadata;
    }

    private static OfficeDocumentMetadataEntry EpubMetadata(string id, string name, object value, string valueType) => new OfficeDocumentMetadataEntry {
        Id = id, Category = "epub.package", Name = name, Value = Convert.ToString(value, CultureInfo.InvariantCulture), ValueType = valueType
    };

    private static string ComputeEpubPayloadHash(byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        return ComputeSha256Hex(stream);
    }
}
