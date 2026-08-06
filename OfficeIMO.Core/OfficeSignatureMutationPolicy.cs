namespace OfficeIMO;

/// <summary>
/// Controls how a save handles digital signatures that package mutation would invalidate.
/// </summary>
public enum OfficeSignatureMutationPolicy {
    /// <summary>Block the save rather than silently invalidating an existing signature.</summary>
    BlockSave,

    /// <summary>Remove invalidated signature parts and related application metadata.</summary>
    RemoveInvalidatedSignatures,

    /// <summary>Preserve signature markup even though the rewritten package may invalidate it.</summary>
    PreserveSignatureMarkup
}
