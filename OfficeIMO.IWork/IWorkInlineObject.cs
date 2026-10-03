namespace OfficeIMO.IWork;

/// <summary>A drawable attachment at one character position in its source text storage.</summary>
public sealed class IWorkInlineObject {
    internal IWorkInlineObject(int characterOffset, IWorkObjectIdentity attachment, IWorkObjectIdentity drawable) {
        CharacterOffset = characterOffset;
        Attachment = attachment;
        Drawable = drawable;
    }

    /// <summary>Gets the zero-based UTF-16 character offset in the original text storage, before marker removal.</summary>
    public int CharacterOffset { get; }
    /// <summary>Gets the native attachment record identity.</summary>
    public IWorkObjectIdentity Attachment { get; }
    /// <summary>Gets the native image or table record identity referenced by the attachment.</summary>
    public IWorkObjectIdentity Drawable { get; }
}
