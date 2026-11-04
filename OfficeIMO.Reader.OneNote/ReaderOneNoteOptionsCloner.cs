using OfficeIMO.OneNote;

namespace OfficeIMO.Reader.OneNote;

internal static class ReaderOneNoteOptionsCloner {
    internal static ReaderOneNoteOptions CloneOrDefault(ReaderOneNoteOptions? options) {
        options ??= new ReaderOneNoteOptions();
        OneNoteLoadOptions native = CloneNative(options.OneNoteOptions ?? new OneNoteLoadOptions());
        OneNoteNotebookReaderOptions notebook = CloneNotebook(options.NotebookOptions ?? new OneNoteNotebookReaderOptions());
        notebook.OneNoteOptions = native;
        return new ReaderOneNoteOptions {
            AllowTableOfContentsSiblingFileReads = options.AllowTableOfContentsSiblingFileReads,
            IncludeAssetPayloads = options.IncludeAssetPayloads,
            IncludeConflictPages = options.IncludeConflictPages,
            IncludeVersionHistory = options.IncludeVersionHistory,
            OneNoteOptions = native,
            NotebookOptions = notebook
        };
    }

    internal static ReaderOneNoteOptions? CloneNullable(ReaderOneNoteOptions? options) {
        return options == null ? null : CloneOrDefault(options);
    }

    private static OneNoteLoadOptions CloneNative(OneNoteLoadOptions options) {
        return new OneNoteLoadOptions {
            MaxInputBytes = options.MaxInputBytes,
            MaxFileNodeListFragments = options.MaxFileNodeListFragments,
            MaxFileNodes = options.MaxFileNodes,
            MaxTransactionLogFragments = options.MaxTransactionLogFragments,
            MaxTransactionEntries = options.MaxTransactionEntries,
            MaxObjects = options.MaxObjects,
            MaxPropertiesPerObject = options.MaxPropertiesPerObject,
            MaxPropertySetDepth = options.MaxPropertySetDepth,
            MaxPageGraphNodes = options.MaxPageGraphNodes,
            MaxPageRelationshipDepth = options.MaxPageRelationshipDepth,
            MaxAssetBytes = options.MaxAssetBytes,
            MaxTotalAssetBytes = options.MaxTotalAssetBytes,
            MaxStreamObjects = options.MaxStreamObjects,
            MaxStreamObjectDepth = options.MaxStreamObjectDepth,
            StrictHeaderValidation = options.StrictHeaderValidation,
            ValidateTransactionChecksums = options.ValidateTransactionChecksums,
            PreserveUnknownData = options.PreserveUnknownData
        };
    }

    private static OneNoteNotebookReaderOptions CloneNotebook(OneNoteNotebookReaderOptions options) {
        return new OneNoteNotebookReaderOptions {
            LoadSectionContent = options.LoadSectionContent,
            ContinueOnSectionError = options.ContinueOnSectionError,
            RecurseSectionGroups = options.RecurseSectionGroups,
            IncludeRecycleBin = options.IncludeRecycleBin,
            MaxSectionGroupDepth = options.MaxSectionGroupDepth,
            MaxNotebookEntries = options.MaxNotebookEntries,
            MaxPackageEntries = options.MaxPackageEntries,
            MaxPackageExpandedBytes = options.MaxPackageExpandedBytes,
            MaxPackageEntryBytes = options.MaxPackageEntryBytes
        };
    }
}
