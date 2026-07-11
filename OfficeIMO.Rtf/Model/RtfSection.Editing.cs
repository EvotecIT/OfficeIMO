namespace OfficeIMO.Rtf;

/// <content>Provides internal section synchronization for document-level block editing.</content>
public sealed partial class RtfSection {
    internal int IndexOfBlock(IRtfBlock block) => _blocks.IndexOf(block);

    internal void InsertBlock(int index, IRtfBlock block) => _blocks.Insert(index, block);

    internal bool RemoveBlock(IRtfBlock block) => _blocks.Remove(block);
}
