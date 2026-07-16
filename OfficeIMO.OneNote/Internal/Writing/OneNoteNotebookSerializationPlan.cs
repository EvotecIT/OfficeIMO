namespace OfficeIMO.OneNote;

internal sealed class OneNoteNotebookSerializationPlan {
    private const string TocFileName = "Open Notebook.onetoc2";
    private readonly OneNoteWriterOptions _options;
    private readonly List<OneNoteCabinetEntry> _entries = new List<OneNoteCabinetEntry>();
    private long _expandedBytes;

    private OneNoteNotebookSerializationPlan(OneNoteWriterOptions options) { _options = options; }

    internal IReadOnlyList<OneNoteCabinetEntry> Entries => _entries.AsReadOnly();

    internal static OneNoteNotebookSerializationPlan Create(OneNoteNotebook notebook, OneNoteWriterOptions options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        ValidateOptions(options);
        var plan = new OneNoteNotebookSerializationPlan(options);
        Guid rootId = notebook.Id.HasValue && notebook.Id.Value != Guid.Empty ? notebook.Id.Value : Guid.NewGuid();
        plan.BuildScope(
            string.Empty,
            rootId,
            Guid.Empty,
            notebook.Sections,
            notebook.SectionGroups,
            notebook.ColorArgb,
            notebook.HistoryEnabled,
            notebook.TableOfContentsStorageFormat,
            notebook.TableOfContentsRootObjectId,
            notebook.UnknownObjects);
        return plan;
    }

    internal static byte[] CreateRootTableOfContents(OneNoteNotebook notebook, OneNoteWriterOptions options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        ValidateOptions(options);
        Guid rootId = notebook.Id.HasValue && notebook.Id.Value != Guid.Empty ? notebook.Id.Value : Guid.NewGuid();
        IReadOnlyList<OneNoteTocWriteEntry> entries = CreateTocEntries(notebook.Sections, notebook.SectionGroups);
        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder(options.MaxOutputBytes, options.PreserveUnknownData).BuildTableOfContents(
            rootId,
            Guid.Empty,
            TocFileName,
            entries,
            notebook.ColorArgb,
            notebook.HistoryEnabled,
            notebook.UnknownObjects,
            notebook.TableOfContentsRootObjectId);
        return SerializeGraph(graph, options, true, notebook.TableOfContentsStorageFormat);
    }

    private void BuildScope(
        string prefix,
        Guid tocId,
        Guid ancestorId,
        IList<OneNoteSection> sections,
        IList<OneNoteSectionGroup> groups,
        uint? colorArgb,
        bool? historyEnabled,
        OneNoteStorageFormat sourceStorageFormat,
        OneNoteExtendedGuid? rootObjectId,
        IList<OneNoteOpaqueObject> preservedObjects) {
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var tocEntries = new List<OneNoteTocWriteEntry>();
        uint order = 0;
        foreach (OneNoteSection section in sections) {
            string fileName = UniqueName(GetSectionFileName(section), usedNames);
            Guid sectionId = section.Id.HasValue && section.Id.Value != Guid.Empty ? section.Id.Value : Guid.NewGuid();
            OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder(_options.MaxOutputBytes, _options.PreserveUnknownData).BuildSection(section, tocId, fileName, sectionId);
            AddEntry(Combine(prefix, fileName), SerializeGraph(graph, _options, false, section.StorageFormat));
            tocEntries.Add(new OneNoteTocWriteEntry(sectionId, fileName, order++, section.ColorArgb));
        }
        foreach (OneNoteSectionGroup group in groups) {
            string directoryName = UniqueName(SanitizeName(group.Name, "Section Group"), usedNames);
            Guid groupId = group.Id.HasValue && group.Id.Value != Guid.Empty ? group.Id.Value : Guid.NewGuid();
            tocEntries.Add(new OneNoteTocWriteEntry(groupId, directoryName, order++, null));
            BuildScope(
                Combine(prefix, directoryName),
                groupId,
                tocId,
                group.Sections,
                group.SectionGroups,
                null,
                historyEnabled,
                group.TableOfContentsStorageFormat,
                group.TableOfContentsRootObjectId,
                group.UnknownObjects);
        }
        OneNoteWriteGraph tocGraph = new OneNoteWriteGraphBuilder(_options.MaxOutputBytes, _options.PreserveUnknownData).BuildTableOfContents(
            tocId,
            ancestorId,
            TocFileName,
            tocEntries,
            colorArgb,
            historyEnabled,
            preservedObjects,
            rootObjectId);
        AddEntry(Combine(prefix, TocFileName), SerializeGraph(tocGraph, _options, true, sourceStorageFormat));
    }

    private void AddEntry(string name, byte[] data) {
        if (_entries.Count >= _options.MaxPackageEntries) throw new OneNoteFormatException("ONENOTE_WRITE_ENTRY_LIMIT", "The notebook exceeds MaxPackageEntries.");
        _expandedBytes = checked(_expandedBytes + data.LongLength);
        if (_expandedBytes > _options.MaxOutputBytes) throw new IOException("OneNote notebook output exceeds MaxOutputBytes.");
        _entries.Add(new OneNoteCabinetEntry(name, data));
    }

    private static byte[] SerializeGraph(
        OneNoteWriteGraph graph,
        OneNoteWriterOptions options,
        bool toc,
        OneNoteStorageFormat sourceStorageFormat) {
        byte[] data = OneNoteGraphSerializer.Write(graph, options, sourceStorageFormat);
        if (options.ValidateRoundTrip) {
            using (var stream = new MemoryStream(data, false)) {
                if (toc) OneNoteNotebookReader.Read(stream, TocFileName, new OneNoteNotebookReaderOptions { LoadSectionContent = false });
                else OneNoteSectionReader.Read(stream, new OneNoteReaderOptions { MaxInputBytes = options.MaxOutputBytes });
            }
        }
        return data;
    }

    private static IReadOnlyList<OneNoteTocWriteEntry> CreateTocEntries(IList<OneNoteSection> sections, IList<OneNoteSectionGroup> groups) {
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<OneNoteTocWriteEntry>();
        uint order = 0;
        foreach (OneNoteSection section in sections) {
            string name = UniqueName(GetSectionFileName(section), usedNames);
            result.Add(new OneNoteTocWriteEntry(section.Id ?? Guid.NewGuid(), name, order++, section.ColorArgb));
        }
        foreach (OneNoteSectionGroup group in groups) {
            string name = UniqueName(SanitizeName(group.Name, "Section Group"), usedNames);
            result.Add(new OneNoteTocWriteEntry(group.Id ?? Guid.NewGuid(), name, order++, null));
        }
        return result.AsReadOnly();
    }

    private static string GetSectionFileName(OneNoteSection section) {
        string? source = section.SourcePath;
        if (source != null && !string.IsNullOrWhiteSpace(source)) {
            int packageSeparator = source.LastIndexOf("::", StringComparison.Ordinal);
            if (packageSeparator >= 0) source = source.Substring(packageSeparator + 2);
            source = source.Replace('\\', '/');
            string candidate = source.Substring(source.LastIndexOf('/') + 1);
            if (candidate.EndsWith(".one", StringComparison.OrdinalIgnoreCase) && IsSafeName(candidate)) return candidate;
        }
        return SanitizeName(section.Name, "Section") + ".one";
    }

    private static string SanitizeName(string? value, string fallback) {
        string source = string.IsNullOrWhiteSpace(value) ? fallback : value!.Trim();
        var characters = source.Select(character => character < 32 || "<>:\"/\\|?*".IndexOf(character) >= 0 ? '_' : character).ToArray();
        string result = new string(characters).Trim().TrimEnd('.', ' ');
        if (result.Length == 0 || result == "." || result == "..") result = fallback;
        if (result.Length > 120) result = result.Substring(0, 120).TrimEnd('.', ' ');
        return result;
    }

    private static bool IsSafeName(string value) =>
        !string.IsNullOrWhiteSpace(value) && value != "." && value != ".." && value.All(character => character >= 32 && "<>:\"/\\|?*".IndexOf(character) < 0);

    private static string UniqueName(string requested, ISet<string> used) {
        if (used.Add(requested)) return requested;
        string extension = Path.GetExtension(requested);
        string stem = requested.Substring(0, requested.Length - extension.Length);
        for (int suffix = 2; suffix < int.MaxValue; suffix++) {
            string candidate = stem + " (" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")" + extension;
            if (used.Add(candidate)) return candidate;
        }
        throw new IOException("No unique OneNote notebook entry name could be generated.");
    }

    private static string Combine(string prefix, string name) => string.IsNullOrEmpty(prefix) ? name : prefix + "/" + name;

    private static void ValidateOptions(OneNoteWriterOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (options.MaxOutputBytes < 1) throw new ArgumentOutOfRangeException(nameof(options), "MaxOutputBytes must be greater than zero.");
        if (options.MaxPackageEntries < 1 || options.MaxPackageEntries > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(options), "MaxPackageEntries must be between 1 and 65535.");
    }
}
