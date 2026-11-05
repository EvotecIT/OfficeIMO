namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRouteSettings {
    private static OfficeIMO.AsciiDoc.AsciiDocParseOptions? Snapshot(OfficeIMO.AsciiDoc.AsciiDocParseOptions? value) => value == null ? null : new OfficeIMO.AsciiDoc.AsciiDocParseOptions {
            Profile = value.Profile,
            MaxInputLength = value.MaxInputLength,
            MaxBlockCount = value.MaxBlockCount,
            MaxInlineNestingDepth = value.MaxInlineNestingDepth,
            MaxInlineNodeCount = value.MaxInlineNodeCount,
            MaxTableCellCount = value.MaxTableCellCount,
            MaxTableColumnCount = value.MaxTableColumnCount,
            MaxTableSpan = value.MaxTableSpan
    };

    private static OfficeIMO.AsciiDoc.Markdown.AsciiDocMarkdownExportOptions? Snapshot(OfficeIMO.AsciiDoc.Markdown.AsciiDocMarkdownExportOptions? value) => value == null ? null : new OfficeIMO.AsciiDoc.Markdown.AsciiDocMarkdownExportOptions {
            FailOnLoss = value.FailOnLoss,
            IncludeDocumentAttributesAsFrontMatter = value.IncludeDocumentAttributesAsFrontMatter,
            PreserveUnsupportedAsSource = value.PreserveUnsupportedAsSource,
            PreserveCommentsAsSource = value.PreserveCommentsAsSource,
            ExpandDocumentAttributes = value.ExpandDocumentAttributes,
            UndefinedAttributeBehavior = value.UndefinedAttributeBehavior,
            MaximumBlockNestingDepth = value.MaximumBlockNestingDepth,
            References = value.References
    };

    private static OfficeIMO.AsciiDoc.Markdown.MarkdownAsciiDocExportOptions? Snapshot(OfficeIMO.AsciiDoc.Markdown.MarkdownAsciiDocExportOptions? value) => value == null ? null : new OfficeIMO.AsciiDoc.Markdown.MarkdownAsciiDocExportOptions {
            FailOnLoss = value.FailOnLoss,
            LineEnding = value.LineEnding,
            FirstLevelOneHeadingIsDocumentTitle = value.FirstLevelOneHeadingIsDocumentTitle,
            PreserveUnsupportedAsSource = value.PreserveUnsupportedAsSource
    };

    private static OfficeIMO.Latex.LatexParseOptions? Snapshot(OfficeIMO.Latex.LatexParseOptions? value) {
        if (value == null) return null;
        var copy = new OfficeIMO.Latex.LatexParseOptions {
            Profile = value.Profile,
            MaxInputLength = value.MaxInputLength,
            MaxInputBytes = value.MaxInputBytes,
            MaxTokenCount = value.MaxTokenCount,
            MaxNestingDepth = value.MaxNestingDepth,
            MacroExpansion = value.MacroExpansion,
            MaxExpansionDepth = value.MaxExpansionDepth,
            MaxExpansionLength = value.MaxExpansionLength,
            MaxExpansionInputLength = value.MaxExpansionInputLength,
            MaxExpansionTokenCount = value.MaxExpansionTokenCount
        };
        copy.VerbatimEnvironmentNames.Clear();
        foreach (string name in value.VerbatimEnvironmentNames) copy.VerbatimEnvironmentNames.Add(name);
        return copy;
    }

    private static OfficeIMO.Latex.Markdown.LatexMarkdownExportOptions? Snapshot(OfficeIMO.Latex.Markdown.LatexMarkdownExportOptions? value) => value == null ? null : new OfficeIMO.Latex.Markdown.LatexMarkdownExportOptions {
            FailOnLoss = value.FailOnLoss,
            PreserveUnsupportedAsSource = value.PreserveUnsupportedAsSource,
            IncludePreambleAsFrontMatter = value.IncludePreambleAsFrontMatter
    };

    private static OfficeIMO.Latex.Markdown.MarkdownLatexExportOptions? Snapshot(OfficeIMO.Latex.Markdown.MarkdownLatexExportOptions? value) => value == null ? null : new OfficeIMO.Latex.Markdown.MarkdownLatexExportOptions {
            FailOnLoss = value.FailOnLoss,
            DocumentClass = value.DocumentClass,
            LineEnding = value.LineEnding,
            FirstHeadingIsTitle = value.FirstHeadingIsTitle,
            PreserveUnsupportedAsSource = value.PreserveUnsupportedAsSource
    };

    private static OfficeIMO.Rtf.RtfLoadOptions? Snapshot(OfficeIMO.Rtf.RtfLoadOptions? value) => value == null ? null : new OfficeIMO.Rtf.RtfLoadOptions {
            DeferBinaryPayloads = value.DeferBinaryPayloads,
            Limits = value.Limits.Clone(),
            ReadEmbeddedObjects = value.ReadEmbeddedObjects,
            ReadFileReferences = value.ReadFileReferences,
            HyperlinkPolicy = value.HyperlinkPolicy,
            WarnOnUnsupportedDestinations = value.WarnOnUnsupportedDestinations,
            WarnOnUnsupportedCodePages = value.WarnOnUnsupportedCodePages
    };

    private static OfficeIMO.Rtf.RtfSaveOptions? Snapshot(OfficeIMO.Rtf.RtfSaveOptions? value) => value == null ? null : new OfficeIMO.Rtf.RtfSaveOptions {
            IncludeGenerator = value.IncludeGenerator,
            IncludeHtmlEncapsulation = value.IncludeHtmlEncapsulation,
            MaterializeStyleFormatting = value.MaterializeStyleFormatting,
            DefaultFontName = value.DefaultFontName
    };

    private static OfficeIMO.OpenDocument.OdfLoadOptions? Snapshot(OfficeIMO.OpenDocument.OdfLoadOptions? value) => value == null ? null : new OfficeIMO.OpenDocument.OdfLoadOptions {
            DeferBinaryPayloads = value.DeferBinaryPayloads,
            Password = value.Password,
            MaxPackageBytes = value.MaxPackageBytes,
            MaxEntries = value.MaxEntries,
            MaxEntryUncompressedBytes = value.MaxEntryUncompressedBytes,
            MaxTotalUncompressedBytes = value.MaxTotalUncompressedBytes,
            MaxTotalKdfIterations = value.MaxTotalKdfIterations,
            MaxCompressionRatio = value.MaxCompressionRatio,
            MaxDepth = value.MaxDepth,
            MaxXmlCharacters = value.MaxXmlCharacters,
            MaxXmlDepth = value.MaxXmlDepth
    };

    private static OfficeIMO.OpenDocument.OdfSaveOptions? Snapshot(OfficeIMO.OpenDocument.OdfSaveOptions? value) => value == null ? null : new OfficeIMO.OpenDocument.OdfSaveOptions {
            CompatibilityProfile = value.CompatibilityProfile,
            SignatureHandling = value.SignatureHandling,
            Encryption = value.Encryption == null ? null : new OfficeIMO.OpenDocument.OdfEncryptionOptions { Password = value.Encryption.Password, IterationCount = value.Encryption.IterationCount },
            EncryptionHandling = value.EncryptionHandling,
            Deterministic = value.Deterministic
    };

    private static OfficeIMO.OneNote.OneNoteLoadOptions? Snapshot(OfficeIMO.OneNote.OneNoteLoadOptions? value) => value == null ? null : new OfficeIMO.OneNote.OneNoteLoadOptions {
            MaxInputBytes = value.MaxInputBytes,
            MaxFileNodeListFragments = value.MaxFileNodeListFragments,
            MaxFileNodes = value.MaxFileNodes,
            MaxTransactionLogFragments = value.MaxTransactionLogFragments,
            MaxTransactionEntries = value.MaxTransactionEntries,
            MaxObjects = value.MaxObjects,
            MaxPropertiesPerObject = value.MaxPropertiesPerObject,
            MaxPropertySetDepth = value.MaxPropertySetDepth,
            MaxPageGraphNodes = value.MaxPageGraphNodes,
            MaxPageRelationshipDepth = value.MaxPageRelationshipDepth,
            MaxInkPathValues = value.MaxInkPathValues,
            MaxAssetBytes = value.MaxAssetBytes,
            MaxTotalAssetBytes = value.MaxTotalAssetBytes,
            MaxStreamObjects = value.MaxStreamObjects,
            MaxStreamObjectDepth = value.MaxStreamObjectDepth,
            StrictHeaderValidation = value.StrictHeaderValidation,
            ValidateTransactionChecksums = value.ValidateTransactionChecksums,
            PreserveUnknownData = value.PreserveUnknownData
    };

    private static OfficeIMO.OneNote.OneNoteSaveOptions? Snapshot(OfficeIMO.OneNote.OneNoteSaveOptions? value) => value == null ? null : new OfficeIMO.OneNote.OneNoteSaveOptions {
            StorageFormat = value.StorageFormat,
            MaxOutputBytes = value.MaxOutputBytes,
            MaxInkPathValues = value.MaxInkPathValues,
            ValidateRoundTrip = value.ValidateRoundTrip,
            MaxPackageEntries = value.MaxPackageEntries,
            MaxSectionGroupDepth = value.MaxSectionGroupDepth,
            MaxPageRelationshipDepth = value.MaxPageRelationshipDepth,
            MaxContentDepth = value.MaxContentDepth,
            PreserveUnknownData = value.PreserveUnknownData
    };
    internal static OfficeIMO.OneNote.OneNoteNotebookReaderOptions? Snapshot(OfficeIMO.OneNote.OneNoteNotebookReaderOptions? value) =>
        value == null ? null : new OfficeIMO.OneNote.OneNoteNotebookReaderOptions {
            OneNoteOptions = Snapshot(value.OneNoteOptions)!,
            LoadSectionContent = value.LoadSectionContent,
            ContinueOnSectionError = value.ContinueOnSectionError,
            RecurseSectionGroups = value.RecurseSectionGroups,
            IncludeRecycleBin = value.IncludeRecycleBin,
            MaxSectionGroupDepth = value.MaxSectionGroupDepth,
            MaxNotebookEntries = value.MaxNotebookEntries,
            MaxPackageEntries = value.MaxPackageEntries,
            MaxPackageExpandedBytes = value.MaxPackageExpandedBytes,
            MaxPackageEntryBytes = value.MaxPackageEntryBytes
        };
}
