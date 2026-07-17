namespace OfficeIMO.OneNote;

/// <summary>Reads a local OneNote notebook hierarchy from <c>.onetoc2</c> files and sibling sections.</summary>
public static class OneNoteNotebookReader {
    private static readonly StringComparer FileSystemPathComparer = Path.DirectorySeparatorChar == '\\'
        ? StringComparer.OrdinalIgnoreCase
        : StringComparer.Ordinal;

    /// <summary>Reads a notebook from a local <c>.onetoc2</c> file.</summary>
    public static OneNoteNotebook Read(string tableOfContentsPath, OneNoteNotebookReaderOptions? options = null) {
        if (tableOfContentsPath == null) throw new ArgumentNullException(nameof(tableOfContentsPath));
        string fullPath = Path.GetFullPath(tableOfContentsPath);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("OneNote table-of-contents file does not exist.", fullPath);
        var effective = options ?? new OneNoteNotebookReaderOptions();
        ValidateOptions(effective);

        string directory = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();
        var notebook = new OneNoteNotebook {
            Name = new DirectoryInfo(directory).Name,
            SourcePath = directory
        };
        var state = new ReadState(effective, directory);
        ReadToc(fullPath, notebook, null, state, 0);
        return notebook;
    }

    /// <summary>
    /// Reads the root hierarchy from a caller-owned <c>.onetoc2</c> stream. Because a standalone
    /// stream cannot resolve sibling files, returned sections and groups contain TOC metadata only.
    /// </summary>
    public static OneNoteNotebook Read(
        Stream tableOfContentsStream,
        string sourceName = "Open Notebook.onetoc2",
        OneNoteNotebookReaderOptions? options = null) {
        if (tableOfContentsStream == null) throw new ArgumentNullException(nameof(tableOfContentsStream));
        if (!tableOfContentsStream.CanRead || !tableOfContentsStream.CanSeek) throw new ArgumentException("The OneNote table-of-contents stream must be readable and seekable.", nameof(tableOfContentsStream));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        var effective = options ?? new OneNoteNotebookReaderOptions();
        ValidateOptions(effective);

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(tableOfContentsStream, effective.OneNoteOptions);
        if (store.Header.FileKind != OneNoteFileKind.TableOfContents) {
            throw new OneNoteFormatException("ONENOTE_TOC_KIND", "The requested stream is not a .onetoc2 revision store.");
        }
        OneNoteTocData toc = OneNoteTocMapper.Map(store);
        string name = Path.GetFileNameWithoutExtension(sourceName);
        var notebook = new OneNoteNotebook {
            Name = string.IsNullOrWhiteSpace(name) || string.Equals(name, "Open Notebook", StringComparison.OrdinalIgnoreCase) ? "OneNote notebook" : name,
            SourcePath = sourceName
        };
        ApplyRootTocMetadata(notebook, store, toc);

        uint tableOfContentsOrder = 0;
        foreach (OneNoteTocEntry entry in toc.Entries.OrderBy(item => item.Order).ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)) {
            if (notebook.Sections.Count + notebook.SectionGroups.Count >= effective.MaxNotebookEntries) {
                throw new OneNoteFormatException("ONENOTE_TOC_ENTRY_LIMIT", "The notebook TOC entry limit was exceeded.");
            }
            if (!IsSafeStandaloneEntryName(entry.Name)) {
                notebook.Diagnostics.Add(new OneNoteDiagnostic {
                    Code = "ONENOTE_TOC_PATH",
                    Message = "A TOC entry contains an unsafe child path: " + entry.Name,
                    Severity = OneNoteDiagnosticSeverity.Warning,
                    SourcePath = sourceName
                });
                continue;
            }
            if (entry.IsSection) {
                notebook.Sections.Add(new OneNoteSection {
                    Id = entry.Id,
                    Name = Path.GetFileNameWithoutExtension(entry.Name),
                    ColorArgb = entry.ColorArgb,
                    SourcePath = entry.Name,
                    TableOfContentsOrder = tableOfContentsOrder++
                });
            } else if (effective.IncludeRecycleBin || !string.Equals(entry.Name, "OneNote_RecycleBin", StringComparison.OrdinalIgnoreCase)) {
                notebook.SectionGroups.Add(new OneNoteSectionGroup {
                    Id = entry.Id,
                    Name = entry.Name,
                    RelativePath = entry.Name,
                    IsRecycleBin = string.Equals(entry.Name, "OneNote_RecycleBin", StringComparison.OrdinalIgnoreCase),
                    TableOfContentsOrder = tableOfContentsOrder++
                });
            }
        }
        if (effective.LoadSectionContent && (notebook.Sections.Count > 0 || notebook.SectionGroups.Count > 0)) {
            notebook.Diagnostics.Add(new OneNoteDiagnostic {
                Code = "ONENOTE_TOC_STREAM_METADATA_ONLY",
                Message = "A standalone .onetoc2 stream exposes hierarchy metadata only; use a path or .onepkg to load section content.",
                Severity = OneNoteDiagnosticSeverity.Warning,
                SourcePath = sourceName
            });
        }
        return notebook;
    }

    /// <summary>Applies identity and preserved metadata from a root table of contents to its notebook.</summary>
    internal static void ApplyRootTocMetadata(
        OneNoteNotebook notebook,
        OneNoteRevisionStore store,
        OneNoteTocData toc) {
        notebook.Id = store.Header.FileId;
        notebook.ColorArgb = toc.ColorArgb;
        notebook.HistoryEnabled = toc.HistoryEnabled;
        notebook.TableOfContentsRootObjectId = toc.RootObjectId;
        notebook.TableOfContentsStorageFormat = toc.StorageFormat;
        foreach (OneNoteOpaqueObject item in toc.PreservedObjects) notebook.UnknownObjects.Add(item);
    }

    private static void ReadToc(
        string tocPath,
        OneNoteNotebook notebook,
        OneNoteSectionGroup? group,
        ReadState state,
        int depth) {
        string fullPath = Path.GetFullPath(tocPath);
        if (!state.VisitedTocs.Add(fullPath)) {
            notebook.Diagnostics.Add(new OneNoteDiagnostic {
                Code = "ONENOTE_TOC_CYCLE",
                Message = "A section-group table of contents was already visited: " + fullPath,
                Severity = OneNoteDiagnosticSeverity.Warning,
                SourcePath = fullPath
            });
            return;
        }
        if (depth > state.Options.MaxSectionGroupDepth) {
            throw new OneNoteFormatException("ONENOTE_TOC_DEPTH", "The notebook section-group depth limit was exceeded.");
        }

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(fullPath, state.Options.OneNoteOptions);
        if (store.Header.FileKind != OneNoteFileKind.TableOfContents) {
            throw new OneNoteFormatException("ONENOTE_TOC_KIND", "The requested notebook index is not a .onetoc2 revision store.");
        }
        OneNoteTocData toc = OneNoteTocMapper.Map(store);
        if (group == null) {
            ApplyRootTocMetadata(notebook, store, toc);
        } else {
            group.TableOfContentsRootObjectId = toc.RootObjectId;
            group.TableOfContentsStorageFormat = toc.StorageFormat;
            foreach (OneNoteOpaqueObject item in toc.PreservedObjects) group.UnknownObjects.Add(item);
        }

        string directory = Path.GetDirectoryName(fullPath) ?? state.RootDirectory;
        uint tableOfContentsOrder = 0;
        foreach (OneNoteTocEntry entry in toc.Entries.OrderBy(item => item.Order).ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)) {
            state.EntryCount++;
            if (state.EntryCount > state.Options.MaxNotebookEntries) {
                throw new OneNoteFormatException("ONENOTE_TOC_ENTRY_LIMIT", "The notebook TOC entry limit was exceeded.");
            }
            string? childPath = ResolveChildPath(directory, entry.Name);
            if (childPath == null) {
                notebook.Diagnostics.Add(new OneNoteDiagnostic {
                    Code = "ONENOTE_TOC_PATH",
                    Message = "A TOC entry contains an unsafe child path: " + entry.Name,
                    Severity = OneNoteDiagnosticSeverity.Warning,
                    SourcePath = fullPath
                });
                continue;
            }

            if (entry.IsSection) {
                OneNoteSection section = LoadSection(entry, childPath, state, notebook);
                section.TableOfContentsOrder = tableOfContentsOrder++;
                if (group == null) notebook.Sections.Add(section); else group.Sections.Add(section);
                continue;
            }

            bool recycleBin = string.Equals(entry.Name, "OneNote_RecycleBin", StringComparison.OrdinalIgnoreCase);
            if (recycleBin && !state.Options.IncludeRecycleBin) continue;
            var childGroup = new OneNoteSectionGroup {
                Id = entry.Id,
                Name = entry.Name,
                RelativePath = MakeRelativePath(state.RootDirectory, childPath),
                IsRecycleBin = recycleBin,
                TableOfContentsOrder = tableOfContentsOrder++
            };
            if (group == null) notebook.SectionGroups.Add(childGroup); else group.SectionGroups.Add(childGroup);
            if (!state.Options.RecurseSectionGroups || !Directory.Exists(childPath)) continue;
            string? childToc = FindChildToc(childPath);
            if (childToc == null) {
                notebook.Diagnostics.Add(new OneNoteDiagnostic {
                    Code = "ONENOTE_TOC_GROUP_MISSING",
                    Message = "A section group has no .onetoc2 file: " + childPath,
                    Severity = OneNoteDiagnosticSeverity.Warning,
                    SourcePath = childPath
                });
                continue;
            }
            ReadToc(childToc, notebook, childGroup, state, depth + 1);
        }
    }

    private static OneNoteSection LoadSection(OneNoteTocEntry entry, string path, ReadState state, OneNoteNotebook notebook) {
        OneNoteSection section;
        if (state.Options.LoadSectionContent && File.Exists(path)) {
            try {
                section = OneNoteSectionReader.Read(path, state.Options.OneNoteOptions);
            } catch (Exception exception) when (state.Options.ContinueOnSectionError && IsRecoverableSectionError(exception)) {
                section = new OneNoteSection();
                notebook.Diagnostics.Add(new OneNoteDiagnostic {
                    Code = exception is OneNoteFormatException format ? format.Code : "ONENOTE_SECTION_READ",
                    Message = "A notebook section could not be read: " + exception.Message,
                    Severity = OneNoteDiagnosticSeverity.Warning,
                    SourcePath = path
                });
            }
        } else {
            section = new OneNoteSection { SourcePath = path };
            if (state.Options.LoadSectionContent) {
                notebook.Diagnostics.Add(new OneNoteDiagnostic {
                    Code = "ONENOTE_TOC_SECTION_MISSING",
                    Message = "A section referenced by the notebook TOC does not exist: " + path,
                    Severity = OneNoteDiagnosticSeverity.Warning,
                    SourcePath = path
                });
            }
        }
        section.Id = entry.Id;
        if (string.IsNullOrWhiteSpace(section.Name)) section.Name = Path.GetFileNameWithoutExtension(entry.Name);
        section.ColorArgb = entry.ColorArgb ?? section.ColorArgb;
        section.SourcePath = path;
        return section;
    }

    private static bool IsRecoverableSectionError(Exception exception) {
        return exception is OneNoteFormatException || exception is EndOfStreamException || exception is InvalidDataException;
    }

    internal static string? ResolveChildPath(string parentDirectory, string childName) {
        if (!IsSafeStandaloneEntryName(childName)) return null;
        string parent = EnsureTrailingSeparator(Path.GetFullPath(parentDirectory));
        string candidate = Path.GetFullPath(Path.Combine(parent, childName));
        StringComparison comparison = Path.DirectorySeparatorChar == '\\'
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        if (!candidate.StartsWith(parent, comparison)) return null;
        return IsSafeFileSystemEntry(candidate) ? candidate : null;
    }

    private static bool IsSafeStandaloneEntryName(string name) {
        return !string.IsNullOrWhiteSpace(name) &&
               !Path.IsPathRooted(name) &&
               name.IndexOf('/') < 0 &&
               name.IndexOf('\\') < 0 &&
               name.IndexOf(':') < 0 &&
               name != "." &&
               name != "..";
    }

    private static string? FindChildToc(string directory) {
        string conventional = Path.Combine(directory, "Open Notebook.onetoc2");
        if (File.Exists(conventional) && IsSafeFileSystemEntry(conventional)) return conventional;
        return Directory.EnumerateFiles(directory, "*.onetoc2", SearchOption.TopDirectoryOnly)
            .Where(IsSafeFileSystemEntry)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();
    }

    /// <summary>
    /// Rejects symbolic links, junctions, and other reparse-point children before a TOC-controlled
    /// path is opened. Missing entries remain safe to resolve so the caller can report them normally.
    /// Attribute failures other than a missing entry fail closed.
    /// </summary>
    private static bool IsSafeFileSystemEntry(string path) {
        try {
            return (File.GetAttributes(path) & FileAttributes.ReparsePoint) == 0;
        } catch (FileNotFoundException) {
            return true;
        } catch (DirectoryNotFoundException) {
            return true;
        } catch (IOException) {
            return false;
        } catch (UnauthorizedAccessException) {
            return false;
        }
    }

    private static string MakeRelativePath(string root, string child) {
        Uri rootUri = new Uri(EnsureTrailingSeparator(Path.GetFullPath(root)));
        Uri childUri = new Uri(EnsureTrailingSeparator(Path.GetFullPath(child)));
        return Uri.UnescapeDataString(rootUri.MakeRelativeUri(childUri).ToString()).TrimEnd('/').Replace('/', Path.DirectorySeparatorChar);
    }

    private static string EnsureTrailingSeparator(string path) {
        return path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? path
            : path + Path.DirectorySeparatorChar;
    }

    private static void ValidateOptions(OneNoteNotebookReaderOptions options) {
        if (options.OneNoteOptions == null) throw new ArgumentException("OneNoteOptions cannot be null.", nameof(options));
        if (options.MaxSectionGroupDepth < 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSectionGroupDepth));
        if (options.MaxNotebookEntries < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxNotebookEntries));
    }

    private sealed class ReadState {
        internal ReadState(OneNoteNotebookReaderOptions options, string rootDirectory) {
            Options = options;
            RootDirectory = Path.GetFullPath(rootDirectory);
        }

        internal OneNoteNotebookReaderOptions Options { get; }
        internal string RootDirectory { get; }
        internal HashSet<string> VisitedTocs { get; } = new HashSet<string>(FileSystemPathComparer);
        internal int EntryCount { get; set; }
    }
}
