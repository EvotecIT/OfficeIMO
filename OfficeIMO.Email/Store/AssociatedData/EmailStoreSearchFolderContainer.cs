using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Search-folder container evidence retained on the folder object itself.</summary>
public sealed class EmailStoreSearchFolderContainer {
    internal EmailStoreSearchFolderContainer(EmailStoreFolderInfo folder) {
        Folder = folder;
        PstSearchCriteriaFlags = folder.Mapi.GetNullableValue(MapiKnownProperties.PidTag.PstSearchCriteriaFlags);
        ExtendedFolderFlags = Copy(folder.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.ExtendedFolderFlags));
    }

    /// <summary>Underlying folder metadata and detached MAPI bag.</summary>
    public EmailStoreFolderInfo Folder { get; }
    /// <summary>PST-local search criteria state flags when present.</summary>
    public int? PstSearchCriteriaFlags { get; }
    /// <summary>Exact extended-folder flags used to correlate persistent definitions.</summary>
    public byte[]? ExtendedFolderFlags { get; }
    /// <summary>Whether either source classification or search-specific properties identify a search folder.</summary>
    public bool HasSearchEvidence => Folder.IsSearchFolder || PstSearchCriteriaFlags.HasValue || ExtendedFolderFlags != null;

    private static byte[]? Copy(byte[]? value) => value == null ? null : (byte[])value.Clone();
}
