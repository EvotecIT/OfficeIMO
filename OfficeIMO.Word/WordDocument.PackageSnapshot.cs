using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Reopens snapshot bytes without resetting the source package's XML limits or compatibility policy.</summary>
    internal static WordDocument LoadPackageSnapshot(Stream snapshot, WordprocessingDocument source) =>
        LoadFromStream(snapshot, new WordLoadOptions {
            MaxInputBytes = Math.Max(WordLoadOptions.DefaultMaxInputBytes, snapshot.Length)
        }, captureSourcePackageSecurity: false, detectXml: false,
            packageSettings: WordPackageSnapshot.CreateOpenSettings(source, autoSave: false));

    /// <summary>The SDK package's parts saved and written out as package bytes.</summary>
    internal byte[] CreateSavedPackageBytes() {
        if (_wordprocessingDocument == null) {
            throw new InvalidOperationException("Document couldn't be saved as legacy DOC because WordDocument wasn't provided.");
        }
        _wordprocessingDocument.Save();
        return CreateOpenXmlBytesAfterSave();
    }

    /// <summary>The SDK package's limit on characters in one part, or zero for none.</summary>
    internal long PackageMaxCharactersInPart => _wordprocessingDocument.MaxCharactersInPart;

    /// <summary>Whether the package has a main document part with a body.</summary>
    internal bool HasMainDocumentBody() => _wordprocessingDocument.MainDocumentPart?.Document?.Body != null;

    /// <summary>
    /// A writable copy of the document as it stands, carrying the in-memory table and table-of-contents state a save
    /// reads, so saving the copy leaves this document untouched. A document on OfficeIMO's own Open XML layer is copied
    /// from its package without opening the SDK.
    /// </summary>
    internal WordDocument CreateWritableSnapshot() {
        if (_ownedModel is { } model) {
            byte[] bytes;
            if (model.IsChanged) {
                using var buffer = new MemoryStream(model.SourceBytes.Length + 4096);
                model.Save(buffer);
                bytes = buffer.ToArray();
            } else {
                bytes = model.SourceBytes;
            }
            // The copy may keep reading the stream it was loaded from, so it stays open with the copy.
            var ownedSnapshot = new MemoryStream(bytes, writable: false);
            // The copy keeps the source's XML limits, so a part edited past them fails here as the SDK copy would.
            WordDocument copy = LoadFromStream(ownedSnapshot, new WordLoadOptions {
                MaxInputBytes = Math.Max(WordLoadOptions.DefaultMaxInputBytes, bytes.LongLength),
                OpenSettings = _sdkDeferredLoad?.OpenSettings
            }, captureSourcePackageSecurity: false, detectXml: false);
            copy._generatedTableOfContent = _generatedTableOfContent?.CopyToOwnedSnapshot(copy);
            return copy;
        }

        var cloneStream = new MemoryStream();
        using (var packageClone = WordPackageSnapshot.Clone(_wordprocessingDocument, cloneStream)) { }
        cloneStream.Position = 0;
        WordDocument clone = LoadPackageSnapshot(cloneStream, _wordprocessingDocument);
        WordTable.CopyGridNormalizationState(_wordprocessingDocument, clone._wordprocessingDocument);
        if (_generatedTableOfContent != null) {
            var sourceBlocks = _wordprocessingDocument.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>().ToList();
            int index = sourceBlocks.IndexOf(_generatedTableOfContent.SdtBlock);
            if (index >= 0) {
                SdtBlock block = clone._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>().ElementAt(index);
                clone._generatedTableOfContent = new WordTableOfContent(clone, block, queueUpdateOnOpen: false);
                _generatedTableOfContent.CopyGeneratedRefreshStateTo(clone._generatedTableOfContent);
            }
        }
        return clone;
    }
}
