using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.Reader.OneNote;

/// <summary>Projects native offline OneNote content into OfficeIMO.Reader contracts.</summary>
internal static partial class OneNoteReaderAdapter {
    /// <summary>Reads an offline OneNote section, notebook index, or notebook package from a file path.</summary>
    public static IEnumerable<ReaderChunk> Read(
        string oneNotePath,
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        return ReadCore(oneNotePath, readerOptions, oneNoteOptions, cancellationToken).Chunks;
    }

    /// <summary>Reads an offline OneNote section or notebook package from a caller-owned stream.</summary>
    public static IEnumerable<ReaderChunk> Read(
        Stream oneNoteStream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        return ReadCore(oneNoteStream, sourceName, readerOptions, oneNoteOptions, cancellationToken).Chunks;
    }

    /// <summary>Projects an already loaded section into normalized chunks.</summary>
    public static IEnumerable<ReaderChunk> Read(
        OneNoteSection section,
        string sourceName = "section.one",
        ReaderOptions? readerOptions = null,
        CancellationToken cancellationToken = default) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderOneNoteOptions native = ReaderOneNoteOptionsCloner.CloneOrDefault(null);
        var source = SourceInfo.ForLogicalName(NormalizeSourceName(sourceName));
        return CreateSectionContext(section, source, reader, native, cancellationToken).Chunks;
    }

    /// <summary>Projects an already loaded section with OneNote-specific page-selection options.</summary>
    public static IEnumerable<ReaderChunk> Read(
        OneNoteSection section,
        ReaderOneNoteOptions oneNoteOptions,
        string sourceName = "section.one",
        ReaderOptions? readerOptions = null,
        CancellationToken cancellationToken = default) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        if (oneNoteOptions == null) throw new ArgumentNullException(nameof(oneNoteOptions));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderOneNoteOptions native = ReaderOneNoteOptionsCloner.CloneOrDefault(oneNoteOptions);
        var source = SourceInfo.ForLogicalName(NormalizeSourceName(sourceName));
        return CreateSectionContext(section, source, reader, native, cancellationToken).Chunks;
    }

    /// <summary>Projects an already loaded notebook into normalized chunks.</summary>
    public static IEnumerable<ReaderChunk> Read(
        OneNoteNotebook notebook,
        string sourceName = "notebook.onepkg",
        ReaderOptions? readerOptions = null,
        CancellationToken cancellationToken = default) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderOneNoteOptions native = ReaderOneNoteOptionsCloner.CloneOrDefault(null);
        SourceInfo source = SourceInfo.ForLogicalName(NormalizeSourceName(sourceName));
        return CreateNotebookContext(notebook, source, reader, native, cancellationToken).Chunks;
    }

    /// <summary>Projects an already loaded notebook with OneNote-specific page-selection options.</summary>
    public static IEnumerable<ReaderChunk> Read(
        OneNoteNotebook notebook,
        ReaderOneNoteOptions oneNoteOptions,
        string sourceName = "notebook.onepkg",
        ReaderOptions? readerOptions = null,
        CancellationToken cancellationToken = default) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        if (oneNoteOptions == null) throw new ArgumentNullException(nameof(oneNoteOptions));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderOneNoteOptions native = ReaderOneNoteOptionsCloner.CloneOrDefault(oneNoteOptions);
        SourceInfo source = SourceInfo.ForLogicalName(NormalizeSourceName(sourceName));
        return CreateNotebookContext(notebook, source, reader, native, cancellationToken).Chunks;
    }

    private static ReadContext ReadCore(
        string path,
        ReaderOptions? readerOptions,
        ReaderOneNoteOptions? oneNoteOptions,
        CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("OneNote path cannot be empty.", nameof(path));
        if (!File.Exists(path)) throw new FileNotFoundException("OneNote file does not exist.", path);

        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderOneNoteOptions native = PrepareOptions(reader, oneNoteOptions);
        ReaderInputLimits.EnforceFileSize(path, EffectiveLimit(reader.MaxInputBytes, native.OneNoteOptions.MaxInputBytes));
        SourceInfo source = SourceInfo.ForPath(path, reader.ComputeHashes);
        cancellationToken.ThrowIfCancellationRequested();
        OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(path, native.OneNoteOptions);
        if (header.FileKind == OneNoteFileKind.NotebookPackage) {
            OneNoteNotebook package = OneNotePackageReader.Read(path, native.NotebookOptions);
            return CreateNotebookContext(package, source, reader, native, cancellationToken);
        }
        if (header.FileKind == OneNoteFileKind.TableOfContents) {
            OneNoteNotebook notebook = OneNoteNotebookReader.Read(path, native.NotebookOptions);
            return CreateNotebookContext(notebook, source, reader, native, cancellationToken);
        }
        OneNoteSection section = OneNoteSectionReader.Read(path, native.OneNoteOptions);
        cancellationToken.ThrowIfCancellationRequested();
        return CreateSectionContext(section, source, reader, native, cancellationToken);
    }

    private static ReadContext ReadCore(
        Stream stream,
        string? sourceName,
        ReaderOptions? readerOptions,
        ReaderOneNoteOptions? oneNoteOptions,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("OneNote stream must be readable.", nameof(stream));

        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderOneNoteOptions native = PrepareOptions(reader, oneNoteOptions);
        long? limit = EffectiveLimit(reader.MaxInputBytes, native.OneNoteOptions.MaxInputBytes);
        string logicalName = NormalizeSourceName(sourceName);
        Stream parseStream = EnsureSeekable(stream, limit, cancellationToken, out bool ownsStream);
        long originalPosition = parseStream.Position;
        try {
            ReaderInputLimits.EnforceSeekableStreamSize(parseStream, limit);
            SourceInfo source = SourceInfo.ForStream(logicalName, parseStream, reader.ComputeHashes);
            parseStream.Position = 0;
            cancellationToken.ThrowIfCancellationRequested();
            OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(parseStream, native.OneNoteOptions);
            parseStream.Position = 0;
            if (header.FileKind == OneNoteFileKind.NotebookPackage) {
                OneNoteNotebook package = OneNotePackageReader.Read(parseStream, logicalName, native.NotebookOptions);
                return CreateNotebookContext(package, source, reader, native, cancellationToken);
            }
            if (header.FileKind == OneNoteFileKind.TableOfContents) {
                OneNoteNotebook notebook = OneNoteNotebookReader.Read(parseStream, logicalName, native.NotebookOptions);
                return CreateNotebookContext(notebook, source, reader, native, cancellationToken);
            }
            OneNoteSection section = OneNoteSectionReader.Read(parseStream, native.OneNoteOptions);
            cancellationToken.ThrowIfCancellationRequested();
            return CreateSectionContext(section, source, reader, native, cancellationToken);
        } finally {
            if (ownsStream) {
                parseStream.Dispose();
            } else if (parseStream.CanSeek) {
                parseStream.Position = originalPosition;
            }
        }
    }

    private static ReaderOneNoteOptions PrepareOptions(ReaderOptions reader, ReaderOneNoteOptions? options) {
        ReaderOneNoteOptions clone = ReaderOneNoteOptionsCloner.CloneOrDefault(options);
        clone.OneNoteOptions.MaxInputBytes = EffectiveLimit(reader.MaxInputBytes, clone.OneNoteOptions.MaxInputBytes);
        clone.NotebookOptions.OneNoteOptions = clone.OneNoteOptions;
        return clone;
    }

    private static ReadContext CreateSectionContext(
        OneNoteSection section,
        SourceInfo source,
        ReaderOptions reader,
        ReaderOneNoteOptions native,
        CancellationToken cancellationToken) {
        OneNoteMarkdownModelValidator.ValidateSection(section, CreateProjectionValidationOptions(native));
        PageSelection selection = SelectPages(section, native);
        ReaderChunk[] chunks = BuildChunks(selection.Section, source, reader, cancellationToken, selection.Hierarchy).ToArray();
        return new ReadContext(selection.Section, section, null, selection.Hierarchy, source, reader, native, chunks);
    }

    private static ReadContext CreateNotebookContext(
        OneNoteNotebook notebook,
        SourceInfo source,
        ReaderOptions reader,
        ReaderOneNoteOptions native,
        CancellationToken cancellationToken) {
        OneNoteMarkdownModelValidator.ValidateNotebook(notebook, CreateProjectionValidationOptions(native));
        string notebookName = string.IsNullOrWhiteSpace(notebook.Name) ? "OneNote notebook" : OneNoteTextProjection.Normalize(notebook.Name);
        var aggregate = new OneNoteSection { Name = notebookName, SourcePath = notebook.SourcePath };
        var metadataAggregate = new OneNoteSection { Name = notebookName, SourcePath = notebook.SourcePath };
        foreach (OneNoteDiagnostic diagnostic in notebook.Diagnostics) aggregate.Diagnostics.Add(diagnostic);
        var hierarchy = new List<string>();
        NotebookSectionScope[] scopes = EnumerateNotebookSections(notebook, notebookName).ToArray();
        foreach (NotebookSectionScope scope in scopes) {
            cancellationToken.ThrowIfCancellationRequested();
            OneNoteSection section = scope.Section;
            PageSelection selection = SelectPages(section, native);
            for (int index = 0; index < selection.Section.Pages.Count; index++) {
                aggregate.Pages.Add(selection.Section.Pages[index]);
                hierarchy.Add(scope.ParentHierarchy + " > " + selection.Hierarchy[index]);
            }
            foreach (OneNotePage page in section.Pages) metadataAggregate.Pages.Add(page);
            foreach (OneNoteRevision revision in section.Revisions) {
                aggregate.Revisions.Add(revision);
                metadataAggregate.Revisions.Add(revision);
            }
            foreach (OneNoteOpaqueObject item in section.UnknownObjects) aggregate.UnknownObjects.Add(item);
            foreach (OneNoteDiagnostic diagnostic in section.Diagnostics) aggregate.Diagnostics.Add(diagnostic);
        }
        string[] pageHierarchy = hierarchy.ToArray();
        var chunks = new List<ReaderChunk>(BuildChunks(aggregate, source, reader, cancellationToken, pageHierarchy));
        int emptySectionIndex = 0;
        foreach (NotebookSectionScope scope in scopes.Where(scope => scope.Section.Pages.Count == 0)) {
            cancellationToken.ThrowIfCancellationRequested();
            string sectionName = OneNoteTextProjection.Normalize(scope.Section.Name);
            string heading = scope.ParentHierarchy + " > " + sectionName;
            string id = "onenote-section-" + (++emptySectionIndex).ToString("D4", CultureInfo.InvariantCulture);
            string[] warnings = notebook.Diagnostics.Concat(scope.Section.Diagnostics)
                .Where(static diagnostic => diagnostic.Severity != OneNoteDiagnosticSeverity.Information)
                .Select(static diagnostic => OneNoteTextProjection.Normalize(diagnostic.Code + ": " + diagnostic.Message))
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            var chunk = new ReaderChunk {
                Id = id,
                Kind = ReaderInputKind.OneNote,
                Location = new ReaderLocation {
                    Path = source.Path,
                    SourceBlockIndex = emptySectionIndex - 1,
                    SourceBlockKind = "section",
                    BlockAnchor = id,
                    HeadingPath = heading,
                    HierarchyHeadingPath = heading
                },
                SourceId = source.SourceId,
                SourceHash = source.SourceHash,
                SourceLastWriteUtc = source.LastWriteUtc,
                SourceLengthBytes = source.LengthBytes,
                Text = sectionName,
                Markdown = "## " + EscapeMarkdown(sectionName),
                Warnings = warnings.Length == 0 ? null : warnings
            };
            chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown);
            if (reader.ComputeHashes) chunk.ChunkHash = ComputeHash(BuildChunkHashInput(chunk));
            chunks.Add(chunk);
        }
        return new ReadContext(aggregate, metadataAggregate, notebook, pageHierarchy, source, reader, native, chunks.ToArray());
    }

    private static OneNoteMarkdownOptions CreateProjectionValidationOptions(ReaderOneNoteOptions options) {
        return new OneNoteMarkdownOptions {
            // Reader metadata counts all related pages even when their chunk projection is disabled.
            IncludeConflictPages = true,
            IncludeVersionHistory = true,
            MaxSectionGroupDepth = Math.Min(OneNoteWriterOptions.MaximumTraversalDepth, options.NotebookOptions.MaxSectionGroupDepth),
            MaxPageRelationshipDepth = Math.Min(OneNoteWriterOptions.MaximumTraversalDepth, options.OneNoteOptions.MaxPageRelationshipDepth),
            MaxContentDepth = Math.Min(OneNoteWriterOptions.MaximumTraversalDepth, options.OneNoteOptions.MaxPropertySetDepth)
        }.CloneValidated();
    }

    private static IEnumerable<NotebookSectionScope> EnumerateNotebookSections(OneNoteNotebook notebook, string notebookName) {
        foreach (OneNoteNotebookHierarchyItem item in OneNoteNotebookHierarchy.Order(notebook.Sections, notebook.SectionGroups)) {
            if (item.Section != null) {
                yield return new NotebookSectionScope(item.Section, notebookName);
            } else {
                OneNoteSectionGroup group = item.Group!;
                foreach (NotebookSectionScope scope in EnumerateNotebookGroup(group, notebookName + " > " + OneNoteTextProjection.Normalize(group.Name))) yield return scope;
            }
        }
    }

    private static IEnumerable<NotebookSectionScope> EnumerateNotebookGroup(OneNoteSectionGroup group, string parentHierarchy) {
        foreach (OneNoteNotebookHierarchyItem item in OneNoteNotebookHierarchy.Order(group.Sections, group.SectionGroups)) {
            if (item.Section != null) {
                yield return new NotebookSectionScope(item.Section, parentHierarchy);
            } else {
                OneNoteSectionGroup child = item.Group!;
                foreach (NotebookSectionScope scope in EnumerateNotebookGroup(child, parentHierarchy + " > " + OneNoteTextProjection.Normalize(child.Name))) yield return scope;
            }
        }
    }

    private static long? EffectiveLimit(long? first, long? second) {
        if (!first.HasValue) return second;
        if (!second.HasValue) return first;
        return Math.Min(first.Value, second.Value);
    }

    private static Stream EnsureSeekable(Stream input, long? maxBytes, CancellationToken cancellationToken, out bool ownsStream) {
        if (input.CanSeek) {
            ownsStream = false;
            return input;
        }

        var buffer = new MemoryStream();
        var bytes = new byte[64 * 1024];
        long total = 0;
        try {
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = input.Read(bytes, 0, bytes.Length);
                if (read == 0) break;
                total += read;
                if (maxBytes.HasValue && total > maxBytes.Value) {
                    throw new IOException("OneNote input exceeds the configured maximum size.");
                }
                buffer.Write(bytes, 0, read);
            }
            buffer.Position = 0;
            ownsStream = true;
            return buffer;
        } catch {
            buffer.Dispose();
            throw;
        }
    }

    private static string NormalizeSourceName(string? sourceName) {
        return string.IsNullOrWhiteSpace(sourceName) ? "section.one" : sourceName!.Trim();
    }

    private sealed class ReadContext {
        internal ReadContext(OneNoteSection section, OneNoteSection metadataSection, OneNoteNotebook? notebook, IReadOnlyList<string>? pageHierarchy, SourceInfo source, ReaderOptions readerOptions, ReaderOneNoteOptions oneNoteOptions, ReaderChunk[] chunks) {
            Section = section;
            MetadataSection = metadataSection;
            Notebook = notebook;
            PageHierarchy = pageHierarchy;
            Source = source;
            ReaderOptions = readerOptions;
            OneNoteOptions = oneNoteOptions;
            Chunks = chunks;
        }

        internal OneNoteSection Section { get; }
        internal OneNoteSection MetadataSection { get; }
        internal OneNoteNotebook? Notebook { get; }
        internal IReadOnlyList<string>? PageHierarchy { get; }
        internal SourceInfo Source { get; }
        internal ReaderOptions ReaderOptions { get; }
        internal ReaderOneNoteOptions OneNoteOptions { get; }
        internal ReaderChunk[] Chunks { get; }
    }

    private sealed class NotebookSectionScope {
        internal NotebookSectionScope(OneNoteSection section, string parentHierarchy) {
            Section = section;
            ParentHierarchy = parentHierarchy;
        }

        internal OneNoteSection Section { get; }
        internal string ParentHierarchy { get; }
    }

    private sealed class SourceInfo {
        internal string Path { get; private set; } = string.Empty;
        internal string SourceId { get; private set; } = string.Empty;
        internal string? SourceHash { get; private set; }
        internal DateTime? LastWriteUtc { get; private set; }
        internal long? LengthBytes { get; private set; }

        internal static SourceInfo ForLogicalName(string sourceName) {
            return new SourceInfo { Path = sourceName, SourceId = BuildSourceId(sourceName) };
        }

        internal static SourceInfo ForPath(string path, bool computeHash) {
            var info = new FileInfo(path);
            return new SourceInfo {
                Path = path,
                SourceId = BuildSourceId(NormalizePath(path)),
                SourceHash = computeHash ? ComputeFileHash(path) : null,
                LastWriteUtc = info.LastWriteTimeUtc,
                LengthBytes = info.Length
            };
        }

        internal static SourceInfo ForStream(string sourceName, Stream stream, bool computeHash) {
            return new SourceInfo {
                Path = sourceName,
                SourceId = BuildSourceId(sourceName),
                SourceHash = computeHash ? ComputeStreamHash(stream) : null,
                LengthBytes = stream.CanSeek ? stream.Length : (long?)null
            };
        }
    }
}
