using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.Reader.OneNote;

internal static partial class OneNoteReaderAdapter {
    /// <summary>Reads a OneNote section into the shared rich document result.</summary>
    public static OfficeDocumentReadResult ReadDocument(
        string oneNotePath,
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        ReadContext context = ReadCore(oneNotePath, readerOptions, oneNoteOptions, cancellationToken);
        return BuildDocumentResult(context, cancellationToken);
    }

    /// <summary>Reads a OneNote stream into the shared rich document result.</summary>
    public static OfficeDocumentReadResult ReadDocument(
        Stream oneNoteStream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        ReadContext context = ReadCore(oneNoteStream, sourceName, readerOptions, oneNoteOptions, cancellationToken);
        return BuildDocumentResult(context, cancellationToken);
    }

    /// <summary>Projects an already loaded OneNote section into the shared rich document result.</summary>
    public static OfficeDocumentReadResult ReadDocument(
        OneNoteSection section,
        string sourceName = "section.one",
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderOneNoteOptions native = ReaderOneNoteOptionsCloner.CloneOrDefault(oneNoteOptions);
        SourceInfo source = SourceInfo.ForLogicalName(NormalizeSourceName(sourceName));
        return BuildDocumentResult(CreateSectionContext(section, source, reader, native, cancellationToken), cancellationToken);
    }

    /// <summary>Projects an already loaded notebook into the shared rich document result.</summary>
    public static OfficeDocumentReadResult ReadDocument(
        OneNoteNotebook notebook,
        string sourceName = "notebook.onepkg",
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderOneNoteOptions native = ReaderOneNoteOptionsCloner.CloneOrDefault(oneNoteOptions);
        SourceInfo source = SourceInfo.ForLogicalName(NormalizeSourceName(sourceName));
        return BuildDocumentResult(CreateNotebookContext(notebook, source, reader, native, cancellationToken), cancellationToken);
    }

    private static OfficeDocumentReadResult BuildDocumentResult(ReadContext context, CancellationToken cancellationToken) {
        OfficeDocumentAsset[] assets = BuildAssets(context, cancellationToken).ToArray();
        OfficeDocumentLink[] links = BuildLinks(context.Section, context.Source).ToArray();
        var source = new OfficeDocumentSource {
            Path = context.Source.Path,
            SourceId = context.Source.SourceId,
            SourceHash = context.Source.SourceHash,
            LastWriteUtc = context.Source.LastWriteUtc,
            LengthBytes = context.Source.LengthBytes,
            Title = OneNoteTextProjection.Normalize(context.Section.Name),
            Author = NormalizeOptional(context.Section.Pages.Select(static page => page.MostRecentAuthor ?? page.OriginalAuthor).FirstOrDefault(static author => !string.IsNullOrWhiteSpace(author)))
        };
        var capabilities = new List<string> { "officeimo.onenote.native", "officeimo.reader.onenote.offline" };
        if (context.Notebook != null) capabilities.Add("officeimo.onenote.notebook");
        if (context.OneNoteOptions.IncludeConflictPages) capabilities.Add("officeimo.onenote.conflict-pages");
        if (context.OneNoteOptions.IncludeVersionHistory) capabilities.Add("officeimo.onenote.version-history");
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            context.Chunks,
            ReaderInputKind.OneNote,
            source,
            capabilities,
            assets);
        result.Links = links;
        result.Metadata = result.Metadata.Concat(BuildMetadata(context.MetadataSection, assets, links)).ToArray();
        if (context.Notebook != null) result.Metadata = result.Metadata.Concat(BuildNotebookMetadata(context.Notebook)).ToArray();
        EnrichPages(result, context.Section, links);
        return result;
    }

    private static IEnumerable<OfficeDocumentAsset> BuildAssets(ReadContext context, CancellationToken cancellationToken) {
        long totalMaterialized = 0;
        for (int pageIndex = 0; pageIndex < context.Section.Pages.Count; pageIndex++) {
            int assetIndex = 0;
            foreach (OneNoteElement element in EnumerateAllElements(context.Section.Pages[pageIndex])) {
                if (!(element is OneNoteBinaryElement binary)) continue;
                cancellationToken.ThrowIfCancellationRequested();
                string kind = ResolveAssetKind(binary);
                string extension = ResolveExtension(binary.FileName, binary.MediaType, kind);
                string assetId = BuildAssetId(pageIndex, assetIndex);
                byte[]? bytes = null;
                OneNoteBinaryPayload? payload = binary.Payload;
                long? length = payload?.Length;
                bool canMaterialize = payload != null && (!length.HasValue || length.Value <= context.OneNoteOptions.OneNoteOptions.MaxAssetBytes);
                long remainingBudget = context.OneNoteOptions.OneNoteOptions.MaxTotalAssetBytes - totalMaterialized;
                long materializationLimit = Math.Min(context.OneNoteOptions.OneNoteOptions.MaxAssetBytes, remainingBudget);
                if (payload != null && context.OneNoteOptions.IncludeAssetPayloads && canMaterialize && materializationLimit > 0 &&
                    (!length.HasValue || length.Value <= materializationLimit)) {
                    bytes = length.HasValue
                        ? payload.ToArray(materializationLimit)
                        : TryMaterializeUnknownLengthPayload(payload, materializationLimit, cancellationToken);
                    if (bytes != null) {
                        totalMaterialized += bytes.LongLength;
                        length = bytes.LongLength;
                    }
                }
                string? payloadHash = null;
                if (payload != null && context.ReaderOptions.ComputeHashes) {
                    if (bytes != null) payloadHash = ComputeHash(bytes);
                    else if (canMaterialize) payloadHash = ComputePayloadHash(payload, context.OneNoteOptions.OneNoteOptions.MaxAssetBytes);
                }
                OneNoteImage? image = binary as OneNoteImage;
                yield return new OfficeDocumentAsset {
                    Id = assetId,
                    Kind = kind,
                    MediaType = binary.MediaType,
                    Extension = extension,
                    FileName = string.IsNullOrWhiteSpace(binary.FileName) ? OfficeDocumentAssetNaming.BuildFileName(assetId, extension) : OneNoteTextProjection.Normalize(binary.FileName),
                    AltText = NormalizeOptional(image?.AltText),
                    Width = ToPixels(image?.WidthHalfInches),
                    Height = ToPixels(image?.HeightHalfInches),
                    LengthBytes = length,
                    PayloadHash = payloadHash,
                    PayloadBytes = bytes,
                    SourceObjectId = binary.Id?.ToString(),
                    Region = BuildRegion(binary.Layout),
                    Location = BuildLocation(context.Source, pageIndex, kind, assetId)
                };
                assetIndex++;
            }
        }
    }

    private static byte[]? TryMaterializeUnknownLengthPayload(
        OneNoteBinaryPayload payload,
        long maxBytes,
        CancellationToken cancellationToken) {
        using (Stream stream = payload.OpenRead())
        using (var output = new MemoryStream()) {
            int bufferSize = (int)Math.Min(64L * 1024L, maxBytes);
            var buffer = new byte[bufferSize];
            long total = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                long remaining = maxBytes - total;
                int requested = remaining >= buffer.Length ? buffer.Length : (int)remaining + 1;
                int read = stream.Read(buffer, 0, requested);
                if (read <= 0) return output.ToArray();
                if (read > remaining) return null;
                output.Write(buffer, 0, read);
                total += read;
            }
        }
    }

    private static IEnumerable<OfficeDocumentLink> BuildLinks(OneNoteSection section, SourceInfo source) {
        for (int pageIndex = 0; pageIndex < section.Pages.Count; pageIndex++) {
            int linkIndex = 0;
            foreach (OneNoteElement element in EnumerateAllElements(section.Pages[pageIndex])) {
                if (element is OneNoteImage image && !string.IsNullOrWhiteSpace(image.Hyperlink)) {
                    string imageId = "onenote-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-link-" + (linkIndex + 1).ToString("D4", CultureInfo.InvariantCulture);
                    yield return new OfficeDocumentLink {
                        Id = imageId,
                        Kind = "uri",
                        Uri = OneNoteTextProjection.Normalize(image.Hyperlink),
                        Text = NormalizeOptional(image.AltText ?? image.FileName),
                        Location = BuildLocation(source, pageIndex, "image-hyperlink", imageId)
                    };
                    linkIndex++;
                }
                if (!(element is OneNoteParagraph paragraph)) continue;
                foreach (OneNoteTextRun run in paragraph.Runs) {
                    if (string.IsNullOrWhiteSpace(run.Hyperlink)) continue;
                    string id = "onenote-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-link-" + (linkIndex + 1).ToString("D4", CultureInfo.InvariantCulture);
                    yield return new OfficeDocumentLink {
                        Id = id,
                        Kind = "uri",
                        Uri = OneNoteTextProjection.Normalize(run.Hyperlink),
                        Text = OneNoteTextProjection.Normalize(run.Text),
                        Location = BuildLocation(source, pageIndex, "hyperlink", id)
                    };
                    linkIndex++;
                }
            }
        }
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> BuildMetadata(OneNoteSection section, IReadOnlyCollection<OfficeDocumentAsset> assets, IReadOnlyCollection<OfficeDocumentLink> links) {
        OneNotePage[] pages = EnumeratePages(section.Pages).ToArray();
        yield return CountMetadata("onenote-page-count", "PageCount", section.Pages.Count);
        yield return CountMetadata("onenote-revision-count", "RevisionCount", section.Revisions.Count + pages.Sum(static page => page.Revisions.Count));
        yield return CountMetadata("onenote-asset-count", "AssetCount", assets.Count);
        yield return CountMetadata("onenote-link-count", "LinkCount", links.Count);
        yield return CountMetadata("onenote-conflict-page-count", "ConflictPageCount", pages.Count(static page => page.IsConflictPage));
        yield return CountMetadata("onenote-version-page-count", "VersionPageCount", pages.Count(static page => page.IsVersionHistoryPage));
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> BuildNotebookMetadata(OneNoteNotebook notebook) {
        int groupCount = 0;
        int sectionCount = notebook.Sections.Count;
        var groups = new Stack<OneNoteSectionGroup>(notebook.SectionGroups.Reverse());
        while (groups.Count > 0) {
            OneNoteSectionGroup group = groups.Pop();
            groupCount++;
            sectionCount += group.Sections.Count;
            for (int index = group.SectionGroups.Count - 1; index >= 0; index--) groups.Push(group.SectionGroups[index]);
        }
        yield return CountMetadata("onenote-notebook-section-count", "NotebookSectionCount", sectionCount);
        yield return CountMetadata("onenote-section-group-count", "SectionGroupCount", groupCount);
        yield return CountMetadata("onenote-notebook-diagnostic-count", "NotebookDiagnosticCount", notebook.Diagnostics.Count);
    }

    private static IEnumerable<OneNotePage> EnumeratePages(IEnumerable<OneNotePage> pages) {
        foreach (OneNotePage page in pages) {
            yield return page;
            foreach (OneNotePage conflict in EnumeratePages(page.ConflictPages)) yield return conflict;
            foreach (OneNotePage version in EnumeratePages(page.VersionHistory)) yield return version;
        }
    }

    private static OfficeDocumentMetadataEntry CountMetadata(string id, string name, int value) {
        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "onenote.summary",
            Name = name,
            Value = value.ToString(CultureInfo.InvariantCulture),
            ValueType = "count"
        };
    }

    private static void EnrichPages(OfficeDocumentReadResult result, OneNoteSection section, IReadOnlyList<OfficeDocumentLink> links) {
        for (int index = 0; index < result.Pages.Count && index < section.Pages.Count; index++) {
            OneNotePage nativePage = section.Pages[index];
            OfficeDocumentPage page = result.Pages[index];
            page.Name = NormalizeOptional(nativePage.Title);
            page.Width = nativePage.Width;
            page.Height = nativePage.Height;
            int pageNumber = index + 1;
            page.Links = links.Where(link => link.Location.Page == pageNumber).ToArray();
        }
    }

    private static string ResolveAssetKind(OneNoteBinaryElement element) {
        if (element is OneNoteImage) return "image";
        if (element is OneNoteEmbeddedFile) return "embedded-file";
        if (element is OneNoteInk) return "ink";
        if (element is OneNoteMedia) return "media";
        return "binary";
    }

    private static string BuildAssetId(int pageIndex, int assetIndex) {
        return "onenote-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) +
            "-asset-" + (assetIndex + 1).ToString("D4", CultureInfo.InvariantCulture);
    }

    private static string ResolveExtension(string? fileName, string? mediaType, string kind) {
        string extension = string.IsNullOrWhiteSpace(fileName) ? string.Empty : Path.GetExtension(fileName);
        if (!string.IsNullOrWhiteSpace(extension)) return extension;
        if (string.Equals(mediaType, "image/png", StringComparison.OrdinalIgnoreCase)) return ".png";
        if (string.Equals(mediaType, "image/jpeg", StringComparison.OrdinalIgnoreCase)) return ".jpg";
        if (string.Equals(mediaType, "image/gif", StringComparison.OrdinalIgnoreCase)) return ".gif";
        if (string.Equals(mediaType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)) return ".svg";
        return kind == "image" ? ".bin" : ".dat";
    }

    private static OfficeDocumentRegion? BuildRegion(OneNoteLayout? layout) {
        if (layout == null || !layout.X.HasValue || !layout.Y.HasValue || !layout.Width.HasValue || !layout.Height.HasValue) return null;
        return new OfficeDocumentRegion { X = layout.X.Value, Y = layout.Y.Value, Width = layout.Width.Value, Height = layout.Height.Value };
    }

    private static int? ToNullableInt(double? value) {
        if (!value.HasValue || double.IsNaN(value.Value) || double.IsInfinity(value.Value) || value.Value < 0 || value.Value > int.MaxValue) return null;
        return (int)Math.Round(value.Value, MidpointRounding.AwayFromZero);
    }

    private static int? ToPixels(double? halfInches) {
        return ToNullableInt(halfInches * 48D);
    }

    private static string? NormalizeOptional(string? value) {
        return value == null ? null : OneNoteTextProjection.Normalize(value);
    }
}
