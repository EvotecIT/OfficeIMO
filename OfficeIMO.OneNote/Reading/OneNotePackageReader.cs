namespace OfficeIMO.OneNote;

/// <summary>Reads a complete offline OneNote notebook from a Cabinet-based <c>.onepkg</c> archive.</summary>
public static class OneNotePackageReader {
    /// <summary>Reads a <c>.onepkg</c> archive from a file path.</summary>
    public static OneNoteNotebook Read(string packagePath, OneNoteNotebookReaderOptions? options = null) {
        if (packagePath == null) throw new ArgumentNullException(nameof(packagePath));
        string fullPath = Path.GetFullPath(packagePath);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("OneNote package does not exist.", fullPath);
        var effective = options ?? new OneNoteNotebookReaderOptions();
        ValidateOptions(effective);
        using (var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
            return ReadCore(stream, Path.GetFileNameWithoutExtension(fullPath), fullPath, effective);
        }
    }

    /// <summary>Reads a <c>.onepkg</c> archive from a caller-owned seekable stream.</summary>
    public static OneNoteNotebook Read(
        Stream packageStream,
        string sourceName = "notebook.onepkg",
        OneNoteNotebookReaderOptions? options = null) {
        if (packageStream == null) throw new ArgumentNullException(nameof(packageStream));
        if (!packageStream.CanRead || !packageStream.CanSeek) throw new ArgumentException("The OneNote package stream must be readable and seekable.", nameof(packageStream));
        var effective = options ?? new OneNoteNotebookReaderOptions();
        ValidateOptions(effective);
        return ReadCore(packageStream, Path.GetFileNameWithoutExtension(sourceName), sourceName, effective);
    }

    private static OneNoteNotebook ReadCore(
        Stream stream,
        string notebookName,
        string sourcePath,
        OneNoteNotebookReaderOptions options) {
        IReadOnlyList<OneNoteCabinetEntry> rawEntries = OneNoteCabinetArchiveReader.Read(
            stream,
            options.OneNoteOptions.MaxInputBytes,
            options.MaxPackageExpandedBytes,
            options.MaxPackageEntryBytes,
            options.MaxPackageEntries);
        var entries = new Dictionary<string, OneNoteCabinetEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (OneNoteCabinetEntry entry in rawEntries) {
            string name = NormalizeEntryName(entry.Name);
            if (entries.ContainsKey(name)) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_DUPLICATE_ENTRY", "The .onepkg archive contains duplicate entry paths.");
            }
            entries.Add(name, entry);
        }
        string rootToc = entries.Keys
            .Where(path => path.IndexOf('/') < 0 && path.EndsWith(".onetoc2", StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault()
            ?? throw new OneNoteFormatException("ONENOTE_PACKAGE_TOC", "The .onepkg archive contains no top-level .onetoc2 file.");

        var notebook = new OneNoteNotebook {
            Name = string.IsNullOrWhiteSpace(notebookName) ? "OneNote notebook" : notebookName,
            SourcePath = sourcePath
        };
        var state = new PackageReadState(entries, options, sourcePath);
        ReadToc(rootToc, notebook, null, state, 0);
        return notebook;
    }

    private static void ReadToc(
        string tocPath,
        OneNoteNotebook notebook,
        OneNoteSectionGroup? group,
        PackageReadState state,
        int depth) {
        if (!state.VisitedTocs.Add(tocPath)) {
            AddDiagnostic(notebook, "ONENOTE_TOC_CYCLE", "A packaged section-group TOC was already visited.", state.SourcePath + "::" + tocPath);
            return;
        }
        if (depth > state.Options.MaxSectionGroupDepth) throw new OneNoteFormatException("ONENOTE_TOC_DEPTH", "The notebook section-group depth limit was exceeded.");
        OneNoteCabinetEntry tocEntry = state.Entries[tocPath];
        using (var stream = new MemoryStream(tocEntry.Data, false)) {
            OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(stream, state.Options.OneNoteOptions);
            if (store.Header.FileKind != OneNoteFileKind.TableOfContents) throw new OneNoteFormatException("ONENOTE_TOC_KIND", "A packaged notebook index is not a .onetoc2 revision store.");
            OneNoteTocData toc = OneNoteTocMapper.Map(store);
            if (group == null) {
                OneNoteNotebookReader.ApplyRootTocMetadata(notebook, store, toc);
            } else {
                group.TableOfContentsRootObjectId = toc.RootObjectId;
                group.TableOfContentsStorageFormat = toc.StorageFormat;
                foreach (OneNoteOpaqueObject item in toc.PreservedObjects) group.UnknownObjects.Add(item);
            }
            MapEntries(toc, tocPath, notebook, group, state, depth);
        }
    }

    private static void MapEntries(
        OneNoteTocData toc,
        string tocPath,
        OneNoteNotebook notebook,
        OneNoteSectionGroup? group,
        PackageReadState state,
        int depth) {
        string directory = GetDirectory(tocPath);
        uint tableOfContentsOrder = 0;
        foreach (OneNoteTocEntry entry in toc.Entries.OrderBy(item => item.Order).ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)) {
            state.EntryCount++;
            if (state.EntryCount > state.Options.MaxNotebookEntries) throw new OneNoteFormatException("ONENOTE_TOC_ENTRY_LIMIT", "The notebook TOC entry limit was exceeded.");
            string childPath = CombineEntryPath(directory, entry.Name);
            if (entry.IsSection) {
                OneNoteSection section = LoadSection(entry, childPath, notebook, state);
                section.TableOfContentsOrder = tableOfContentsOrder++;
                if (group == null) notebook.Sections.Add(section); else group.Sections.Add(section);
                continue;
            }

            bool recycleBin = string.Equals(entry.Name, "OneNote_RecycleBin", StringComparison.OrdinalIgnoreCase);
            if (recycleBin && !state.Options.IncludeRecycleBin) continue;
            var childGroup = new OneNoteSectionGroup {
                Id = entry.Id,
                Name = entry.Name,
                RelativePath = childPath.Replace('/', Path.DirectorySeparatorChar),
                IsRecycleBin = recycleBin,
                TableOfContentsOrder = tableOfContentsOrder++
            };
            if (group == null) notebook.SectionGroups.Add(childGroup); else group.SectionGroups.Add(childGroup);
            if (!state.Options.RecurseSectionGroups) continue;
            string? childToc = FindChildToc(state.Entries.Keys, childPath);
            if (childToc == null) {
                AddDiagnostic(notebook, "ONENOTE_TOC_GROUP_MISSING", "A packaged section group has no .onetoc2 entry.", state.SourcePath + "::" + childPath);
                continue;
            }
            ReadToc(childToc, notebook, childGroup, state, depth + 1);
        }
    }

    private static OneNoteSection LoadSection(OneNoteTocEntry entry, string path, OneNoteNotebook notebook, PackageReadState state) {
        OneNoteSection section;
        if (state.Options.LoadSectionContent && state.Entries.TryGetValue(path, out OneNoteCabinetEntry? source)) {
            try {
                using (var stream = new MemoryStream(source.Data, false)) section = OneNoteSectionReader.Read(stream, state.Options.OneNoteOptions);
            } catch (Exception exception) when (state.Options.ContinueOnSectionError && IsRecoverableSectionError(exception)) {
                section = new OneNoteSection();
                AddDiagnostic(
                    notebook,
                    exception is OneNoteFormatException format ? format.Code : "ONENOTE_PACKAGE_SECTION_READ",
                    "A packaged section could not be read: " + exception.Message,
                    state.SourcePath + "::" + path);
            }
        } else {
            section = new OneNoteSection();
            if (state.Options.LoadSectionContent) AddDiagnostic(notebook, "ONENOTE_TOC_SECTION_MISSING", "A section referenced by the packaged TOC is missing.", state.SourcePath + "::" + path);
        }
        section.Id = entry.Id;
        if (string.IsNullOrWhiteSpace(section.Name)) section.Name = Path.GetFileNameWithoutExtension(entry.Name);
        section.ColorArgb = entry.ColorArgb ?? section.ColorArgb;
        section.SourcePath = state.SourcePath + "::" + path;
        return section;
    }

    private static bool IsRecoverableSectionError(Exception exception) {
        return exception is OneNoteFormatException || exception is EndOfStreamException || exception is InvalidDataException;
    }

    internal static string NormalizeEntryName(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new OneNoteFormatException("ONENOTE_PACKAGE_ENTRY_NAME", "The .onepkg archive contains an empty entry name.");
        string normalized = name.Replace('\\', '/');
        if (normalized.StartsWith("/", StringComparison.Ordinal)) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_ENTRY_PATH", "The .onepkg archive contains an unsafe entry path.");
        }
        string[] parts = normalized.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0 || parts.Any(part => part == "." || part == ".." || part.IndexOf(':') >= 0)) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_ENTRY_PATH", "The .onepkg archive contains an unsafe entry path.");
        }
        return string.Join("/", parts);
    }

    private static string CombineEntryPath(string directory, string childName) {
        string child = NormalizeEntryName(childName);
        if (child.IndexOf('/') >= 0) throw new OneNoteFormatException("ONENOTE_PACKAGE_TOC_PATH", "A packaged TOC child name contains a directory separator.");
        return string.IsNullOrEmpty(directory) ? child : directory + "/" + child;
    }

    private static string GetDirectory(string path) {
        int separator = path.LastIndexOf('/');
        return separator < 0 ? string.Empty : path.Substring(0, separator);
    }

    private static string? FindChildToc(IEnumerable<string> paths, string directory) {
        string conventional = directory + "/Open Notebook.onetoc2";
        return paths.FirstOrDefault(path => string.Equals(path, conventional, StringComparison.OrdinalIgnoreCase))
            ?? paths.Where(path => string.Equals(GetDirectory(path), directory, StringComparison.OrdinalIgnoreCase) && path.EndsWith(".onetoc2", StringComparison.OrdinalIgnoreCase))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
    }

    private static void AddDiagnostic(OneNoteNotebook notebook, string code, string message, string sourcePath) {
        notebook.Diagnostics.Add(new OneNoteDiagnostic { Code = code, Message = message, Severity = OneNoteDiagnosticSeverity.Warning, SourcePath = sourcePath });
    }

    private static void ValidateOptions(OneNoteNotebookReaderOptions options) {
        if (options.OneNoteOptions == null) throw new ArgumentException("OneNoteOptions cannot be null.", nameof(options));
        options.OneNoteOptions.Validate();
        if (options.MaxSectionGroupDepth < 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSectionGroupDepth));
        if (options.MaxNotebookEntries < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxNotebookEntries));
        if (options.MaxPackageEntries < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxPackageEntries));
        if (options.MaxPackageExpandedBytes < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxPackageExpandedBytes));
        if (options.MaxPackageEntryBytes < 1 || options.MaxPackageEntryBytes > options.MaxPackageExpandedBytes) throw new ArgumentOutOfRangeException(nameof(options.MaxPackageEntryBytes));
    }

    private sealed class PackageReadState {
        internal PackageReadState(Dictionary<string, OneNoteCabinetEntry> entries, OneNoteNotebookReaderOptions options, string sourcePath) {
            Entries = entries;
            Options = options;
            SourcePath = sourcePath;
        }
        internal Dictionary<string, OneNoteCabinetEntry> Entries { get; }
        internal OneNoteNotebookReaderOptions Options { get; }
        internal string SourcePath { get; }
        internal HashSet<string> VisitedTocs { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal int EntryCount { get; set; }
    }
}
