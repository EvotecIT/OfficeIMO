namespace OfficeIMO.Email.AddressBook;

/// <summary>Safety and compatibility limits for read-only OAB sessions.</summary>
public sealed class OfflineAddressBookReaderOptions {
    /// <summary>Default bounded reader profile.</summary>
    public static OfflineAddressBookReaderOptions Default { get; } = new OfflineAddressBookReaderOptions();

    /// <summary>Creates an immutable reader profile.</summary>
    public OfflineAddressBookReaderOptions(
        long maxInputBytes = 64L * 1024 * 1024 * 1024,
        int maxDiscoveredFiles = 4096,
        int maxDirectoryDepth = 16,
        int maxMetadataBytes = 8 * 1024 * 1024,
        int maxPropertiesPerTable = 4096,
        int maxRecordBytes = 16 * 1024 * 1024,
        int maxStringBytes = 4 * 1024 * 1024,
        int maxBinaryBytes = 16 * 1024 * 1024,
        int maxValuesPerProperty = 100_000,
        long maxDeclaredEntries = 100_000_000,
        int string8CodePage = 1252,
        bool retainRawPropertyBytes = true,
        int maxDirectoryEntries = 100_000) {
        if (maxInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxInputBytes));
        if (maxDiscoveredFiles <= 0) throw new ArgumentOutOfRangeException(nameof(maxDiscoveredFiles));
        if (maxDirectoryDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxDirectoryDepth));
        if (maxMetadataBytes < 16) throw new ArgumentOutOfRangeException(nameof(maxMetadataBytes));
        if (maxPropertiesPerTable <= 0) throw new ArgumentOutOfRangeException(nameof(maxPropertiesPerTable));
        if (maxRecordBytes < 5) throw new ArgumentOutOfRangeException(nameof(maxRecordBytes));
        if (maxStringBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxStringBytes));
        if (maxBinaryBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBinaryBytes));
        if (maxValuesPerProperty <= 0) throw new ArgumentOutOfRangeException(nameof(maxValuesPerProperty));
        if (maxDeclaredEntries <= 0) throw new ArgumentOutOfRangeException(nameof(maxDeclaredEntries));
        if (string8CodePage <= 0) throw new ArgumentOutOfRangeException(nameof(string8CodePage));
        if (maxDirectoryEntries <= 0) throw new ArgumentOutOfRangeException(nameof(maxDirectoryEntries));

        MaxInputBytes = maxInputBytes;
        MaxDiscoveredFiles = maxDiscoveredFiles;
        MaxDirectoryDepth = maxDirectoryDepth;
        MaxMetadataBytes = maxMetadataBytes;
        MaxPropertiesPerTable = maxPropertiesPerTable;
        MaxRecordBytes = maxRecordBytes;
        MaxStringBytes = maxStringBytes;
        MaxBinaryBytes = maxBinaryBytes;
        MaxValuesPerProperty = maxValuesPerProperty;
        MaxDeclaredEntries = maxDeclaredEntries;
        String8CodePage = string8CodePage;
        RetainRawPropertyBytes = retainRawPropertyBytes;
        MaxDirectoryEntries = maxDirectoryEntries;
    }

    /// <summary>
    /// Creates the reader profile exposed before directory-entry discovery was bounded.
    /// Retained so already-compiled callers continue to bind to the original constructor signature.
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    public OfflineAddressBookReaderOptions(
        long maxInputBytes,
        int maxDiscoveredFiles,
        int maxDirectoryDepth,
        int maxMetadataBytes,
        int maxPropertiesPerTable,
        int maxRecordBytes,
        int maxStringBytes,
        int maxBinaryBytes,
        int maxValuesPerProperty,
        long maxDeclaredEntries,
        int string8CodePage,
        bool retainRawPropertyBytes)
        : this(
            maxInputBytes,
            maxDiscoveredFiles,
            maxDirectoryDepth,
            maxMetadataBytes,
            maxPropertiesPerTable,
            maxRecordBytes,
            maxStringBytes,
            maxBinaryBytes,
            maxValuesPerProperty,
            maxDeclaredEntries,
            string8CodePage,
            retainRawPropertyBytes,
            maxDirectoryEntries: 100_000) {
    }

    /// <summary>Maximum bytes in one OAB component.</summary>
    public long MaxInputBytes { get; }
    /// <summary>Maximum .oab files discovered below a directory root.</summary>
    public int MaxDiscoveredFiles { get; }
    /// <summary>Maximum directory recursion depth.</summary>
    public int MaxDirectoryDepth { get; }
    /// <summary>Maximum bytes in the version 4 schema metadata structure.</summary>
    public int MaxMetadataBytes { get; }
    /// <summary>Maximum definitions in either header or entry property table.</summary>
    public int MaxPropertiesPerTable { get; }
    /// <summary>Maximum bytes in one header or address-book-object record.</summary>
    public int MaxRecordBytes { get; }
    /// <summary>Maximum encoded bytes in one string value.</summary>
    public int MaxStringBytes { get; }
    /// <summary>Maximum bytes in one binary value.</summary>
    public int MaxBinaryBytes { get; }
    /// <summary>Maximum values in one multi-valued property.</summary>
    public int MaxValuesPerProperty { get; }
    /// <summary>Maximum declared records in one Full Details file.</summary>
    public long MaxDeclaredEntries { get; }
    /// <summary>Code page used for OAB PtypString8 values.</summary>
    public int String8CodePage { get; }
    /// <summary>Whether decoded MAPI properties retain their original OAB value encoding.</summary>
    public bool RetainRawPropertyBytes { get; }
    /// <summary>Maximum total filesystem entries inspected during directory discovery, including non-OAB entries.</summary>
    public int MaxDirectoryEntries { get; }
}
