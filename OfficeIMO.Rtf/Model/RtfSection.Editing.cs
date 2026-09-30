namespace OfficeIMO.Rtf;

/// <content>Provides internal section synchronization for document-level block editing.</content>
public sealed partial class RtfSection {
    internal void Attach(RtfDocument document) {
        if (_document != null && !ReferenceEquals(_document, document))
            throw new InvalidOperationException("The section already belongs to another document.");
        _document = document;
    }
    internal int IndexOfBlock(IRtfBlock block) => _blocks.IndexOf(block);

    internal void InsertBlock(int index, IRtfBlock block) => _blocks.Insert(index, block);

    internal bool RemoveBlock(IRtfBlock block) => _blocks.Remove(block);
}
