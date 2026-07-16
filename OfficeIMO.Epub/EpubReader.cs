namespace OfficeIMO.Epub;

using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;

/// <summary>
/// Standards-based EPUB extractor using container/OPF/spine/navigation metadata.
/// </summary>
internal static partial class EpubReader {
    /// <summary>
    /// Reads an EPUB document from disk.
    /// </summary>
    public static EpubDocument Read(string epubPath, EpubReadOptions? options = null) {
        if (epubPath == null) throw new ArgumentNullException(nameof(epubPath));
        if (epubPath.Length == 0) throw new ArgumentException("EPUB path cannot be empty.", nameof(epubPath));
        if (!File.Exists(epubPath)) throw new FileNotFoundException($"EPUB file '{epubPath}' doesn't exist.", epubPath);

        using var fs = new FileStream(epubPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Read(fs, options);
    }

    /// <summary>
    /// Reads an EPUB document from a stream.
    /// </summary>
    public static EpubDocument Read(Stream epubStream, EpubReadOptions? options = null) {
        if (epubStream == null) throw new ArgumentNullException(nameof(epubStream));
        if (!epubStream.CanRead) throw new IOException("EPUB stream must be readable.");

        EpubReadOptions effective = Normalize(options);
        try {
            return ReadBytes(OfficeStreamReader.ReadAllBytes(epubStream, effective.MaxPackageBytes), effective);
        } catch (EpubReadException) {
            throw;
        } catch (InvalidDataException exception) {
            throw CreateFatalReadException(
                "epub.package.size-limit",
                exception.Message,
                null,
                exception);
        }
    }

    internal static async Task<EpubDocument> ReadAsync(
        Stream epubStream,
        EpubReadOptions? options,
        CancellationToken cancellationToken) {
        if (epubStream == null) throw new ArgumentNullException(nameof(epubStream));
        if (!epubStream.CanRead) throw new IOException("EPUB stream must be readable.");

        EpubReadOptions effective = Normalize(options);
        try {
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(
                epubStream,
                cancellationToken,
                effective.MaxPackageBytes).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return ReadBytes(bytes, effective);
        } catch (EpubReadException) {
            throw;
        } catch (InvalidDataException exception) {
            throw CreateFatalReadException(
                "epub.package.size-limit",
                exception.Message,
                null,
                exception);
        }
    }

    internal static EpubDocument ReadBytes(byte[] bytes, EpubReadOptions? options = null) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        EpubReadOptions effective = Normalize(options);
        if (bytes.LongLength > effective.MaxPackageBytes) {
            throw CreateFatalReadException(
                "epub.package.size-limit",
                $"EPUB package size {bytes.LongLength} exceeds MaxPackageBytes ({effective.MaxPackageBytes}).");
        }

        try {
            return ReadArchive(bytes, effective);
        } catch (EpubReadException) {
            throw;
        } catch (InvalidDataException exception) {
            throw CreateFatalReadException(
                "epub.archive.invalid",
                "EPUB container is not a readable ZIP archive.",
                null,
                exception);
        }
    }

    private static EpubDocument ReadArchive(byte[] bytes, EpubReadOptions effective) {
        using var epubStream = new MemoryStream(bytes, writable: false);
        var diagnostics = new EpubDiagnosticCollector();

        using var archive = new ZipArchive(epubStream, ZipArchiveMode.Read, leaveOpen: true);
        Dictionary<string, ZipArchiveEntry> entryIndex = BuildEntryIndex(archive, effective, diagnostics);
        EpubPackage? package = TryReadPackage(entryIndex, effective, diagnostics, out IReadOnlyList<EpubRootfile> rootfiles);
        IReadOnlyList<EpubEncryptionInfo> encryption = ReadEncryption(entryIndex, effective, diagnostics);
        Dictionary<string, EpubEncryptionInfo> encryptionByPath = encryption
            .GroupBy(static item => item.Path, StringComparer.Ordinal)
            .ToDictionary(static group => group.Key, static group => group.First(), StringComparer.Ordinal);
        EpubNavigationResult navigation = ReadNavigation(entryIndex, package, effective, diagnostics);
        List<ChapterCandidate> candidates = BuildChapterCandidates(entryIndex, package, effective, diagnostics);

        var chapters = new List<EpubChapter>();
        int emitted = 0;
        long totalRawHtmlBytes = 0;

        foreach (var candidate in candidates) {
            if (emitted >= effective.MaxChapters) break;
            if (effective.MaxChapterBytes.HasValue && candidate.Entry.Length > effective.MaxChapterBytes.Value) {
                string path = candidate.Path;
                diagnostics.Warning(
                    "epub.chapter.size-limit",
                    $"Skipped chapter '{path}' because size {candidate.Entry.Length} exceeds MaxChapterBytes ({effective.MaxChapterBytes.Value}).",
                    path);
                continue;
            }

            string normalizedPath = candidate.Path;
            encryptionByPath.TryGetValue(normalizedPath, out EpubEncryptionInfo? chapterEncryption);
            if (chapterEncryption?.RequiresDecryption == true) {
                diagnostics.Warning(
                    "epub.chapter.encrypted",
                    $"Skipped encrypted chapter '{normalizedPath}' because its algorithm is not supported.",
                    normalizedPath);
                continue;
            }

            string markup = ReadEntryText(candidate.Entry, effective.MaxChapterBytes);
            if (!TryParseXml(markup, out var chapterDocument) || chapterDocument == null) {
                diagnostics.Warning(
                    "epub.chapter.invalid-xhtml",
                    $"Skipped chapter '{normalizedPath}' because chapter markup is not valid XML/XHTML.",
                    normalizedPath);
                continue;
            }

            var text = ExtractVisibleText(chapterDocument);
            bool hasStructuredContent = HasStructuredChapterContent(chapterDocument);
            string? retainedHtml = null;
            if (effective.IncludeRawHtml) {
                if (candidate.Entry.Length > effective.MaxTotalRawHtmlBytes - totalRawHtmlBytes) {
                    diagnostics.Warning(
                        "epub.chapter.raw-html-total-limit",
                        $"Did not retain raw HTML for chapter '{normalizedPath}' because MaxTotalRawHtmlBytes ({effective.MaxTotalRawHtmlBytes}) was reached.",
                        normalizedPath);
                } else {
                    retainedHtml = markup;
                    totalRawHtmlBytes += candidate.Entry.Length;
                }
            }
            if (text.Length == 0 && !hasStructuredContent) {
                continue;
            }

            emitted++;
            var title = ResolveChapterTitle(chapterDocument, navigation.TitleMap, normalizedPath);

            chapters.Add(new EpubChapter {
                Order = emitted,
                Path = normalizedPath,
                ManifestId = candidate.ManifestId,
                MediaType = candidate.MediaType,
                SpineIndex = candidate.SpineIndex,
                IsLinear = candidate.IsLinear,
                RenditionLayout = candidate.RenditionLayout,
                Encryption = chapterEncryption,
                Title = title,
                Text = text,
                HasStructuredContent = hasStructuredContent,
                Html = retainedHtml
            });
        }

        IReadOnlyList<EpubResource> resources = BuildResources(
            entryIndex,
            package,
            encryptionByPath,
            effective,
            diagnostics);

        if (package?.RenditionLayout == EpubRenditionLayout.PrePaginated || chapters.Any(static chapter => chapter.IsFixedLayout)) {
            diagnostics.Warning(
                "epub.layout.fixed",
                "Fixed-layout EPUB content was detected. Text and structure are extracted, but page geometry is not reproduced.",
                package?.OpfPath);
        }

        return new EpubDocument {
            Title = ResolveDocumentTitle(package, chapters),
            Identifier = package?.Identifier,
            Language = package?.Language,
            Creator = package?.Creator,
            OpfPath = package?.OpfPath,
            PackageVersion = package?.PackageVersion,
            UniqueIdentifierId = package?.UniqueIdentifierId,
            RenditionLayout = package?.RenditionLayout,
            Rootfiles = rootfiles,
            Metadata = package?.Metadata.ToArray() ?? Array.Empty<EpubMetadataEntry>(),
            TableOfContents = navigation.TableOfContents.ToArray(),
            PageList = navigation.PageList.ToArray(),
            Landmarks = navigation.Landmarks.ToArray(),
            Chapters = chapters.ToArray(),
            Resources = resources.ToArray(),
            Encryption = encryption.ToArray(),
            Diagnostics = diagnostics.Items,
            Warnings = diagnostics.WarningMessages
        };
    }

    private static bool HasStructuredChapterContent(XDocument document) {
        return document.Descendants().Any(static element => {
            string name = element.Name.LocalName;
            return name.Equals("img", StringComparison.OrdinalIgnoreCase)
                || name.Equals("picture", StringComparison.OrdinalIgnoreCase)
                || name.Equals("svg", StringComparison.OrdinalIgnoreCase)
                || name.Equals("table", StringComparison.OrdinalIgnoreCase)
                || name.Equals("form", StringComparison.OrdinalIgnoreCase)
                || name.Equals("input", StringComparison.OrdinalIgnoreCase)
                || name.Equals("select", StringComparison.OrdinalIgnoreCase)
                || name.Equals("textarea", StringComparison.OrdinalIgnoreCase)
                || name.Equals("audio", StringComparison.OrdinalIgnoreCase)
                || name.Equals("video", StringComparison.OrdinalIgnoreCase)
                || name.Equals("object", StringComparison.OrdinalIgnoreCase)
                || name.Equals("canvas", StringComparison.OrdinalIgnoreCase);
        });
    }

    private static IReadOnlyList<EpubResource> BuildResources(
        IReadOnlyDictionary<string, ZipArchiveEntry> entryIndex,
        EpubPackage? package,
        IReadOnlyDictionary<string, EpubEncryptionInfo> encryptionByPath,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        if (package == null || package.Manifest.Count == 0) return Array.Empty<EpubResource>();
        var resources = new List<EpubResource>(Math.Min(package.Manifest.Count, options.MaxResources));
        long totalPayloadBytes = 0;
        IEnumerable<ManifestItem> items = package.Manifest.Values;
        if (options.DeterministicOrder) items = items.OrderBy(static item => item.FullPath, StringComparer.Ordinal);
        foreach (ManifestItem item in items) {
            if (resources.Count >= options.MaxResources) {
                diagnostics.Warning(
                    "epub.resource.count-limit",
                    $"EPUB manifest resources were truncated at MaxResources ({options.MaxResources}).",
                    package.OpfPath);
                break;
            }
            if (item.IsRemote) {
                diagnostics.Info(
                    "epub.resource.remote",
                    $"Remote EPUB resource '{item.RemoteUri}' is retained as metadata and is not fetched.",
                    item.RemoteUri);
                resources.Add(new EpubResource {
                    Id = item.Id,
                    Path = item.RemoteUri ?? item.Href,
                    Href = item.Href,
                    MediaType = item.MediaType,
                    Properties = item.Properties,
                    IsRemote = true,
                    RemoteUri = item.RemoteUri,
                    LengthBytes = 0
                });
                continue;
            }
            if (!entryIndex.TryGetValue(item.FullPath, out ZipArchiveEntry? entry)) {
                diagnostics.Warning(
                    "epub.resource.missing",
                    $"EPUB manifest resource '{item.FullPath}' was not found in archive.",
                    item.FullPath);
                continue;
            }

            encryptionByPath.TryGetValue(item.FullPath, out EpubEncryptionInfo? resourceEncryption);
            byte[]? data = null;
            if (options.IncludeResourceData) {
                if (resourceEncryption?.RequiresDecryption == true) {
                    diagnostics.Warning(
                        "epub.resource.encrypted",
                        $"Skipped encrypted payload for EPUB resource '{item.FullPath}' because its algorithm is not supported.",
                        item.FullPath);
                } else if (entry.Length > options.MaxResourceBytes) {
                    diagnostics.Warning(
                        "epub.resource.size-limit",
                        $"Skipped payload for EPUB resource '{item.FullPath}' because size {entry.Length} exceeds MaxResourceBytes ({options.MaxResourceBytes}).",
                        item.FullPath);
                } else if (entry.Length > options.MaxTotalResourceBytes - totalPayloadBytes) {
                    diagnostics.Warning(
                        "epub.resource.total-size-limit",
                        $"Skipped payload for EPUB resource '{item.FullPath}' because MaxTotalResourceBytes ({options.MaxTotalResourceBytes}) was reached.",
                        item.FullPath);
                } else {
                    data = ReadEntryBytes(entry, options.MaxResourceBytes);
                    totalPayloadBytes += data.LongLength;
                }
            }
            resources.Add(new EpubResource {
                Id = item.Id,
                Path = item.FullPath,
                Href = item.Href,
                MediaType = item.MediaType,
                Properties = item.Properties,
                LengthBytes = entry.Length,
                Encryption = resourceEncryption,
                Data = data
            });
        }
        return resources;
    }

}
