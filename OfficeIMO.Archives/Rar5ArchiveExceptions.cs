namespace OfficeIMO.Archives;

/// <summary>A selected RAR5 entry uses a method outside the stored reader profile.</summary>
public sealed class Rar5CompressionNotSupportedException : NotSupportedException {
    internal Rar5CompressionNotSupportedException(string entryName, int method, long dictionarySize)
        : base(FormattableString.Invariant($"RAR 5 entry '{entryName}' uses unsupported compression method {method} and a {dictionarySize}-byte dictionary. Only stored entries (method 0) are supported.")) {
        EntryName = entryName;
        Method = method;
        DictionarySize = dictionarySize;
    }

    /// <summary>Authored archive entry name.</summary>
    public string EntryName { get; }
    /// <summary>Compression method from the RAR5 file header.</summary>
    public int Method { get; }
    /// <summary>Dictionary size in bytes declared by the file header.</summary>
    public long DictionarySize { get; }
}

/// <summary>A RAR5 archive has encrypted headers or encrypted member data.</summary>
public sealed class Rar5EncryptionNotSupportedException : NotSupportedException {
    internal Rar5EncryptionNotSupportedException(long byteOffset, string? entryName)
        : base(entryName is null
            ? FormattableString.Invariant($"Encrypted RAR5 headers are not supported (byte offset {byteOffset}).")
            : FormattableString.Invariant($"Encrypted RAR5 data for '{entryName}' is not supported (byte offset {byteOffset}).")) {
        ByteOffset = byteOffset;
        EntryName = entryName;
    }

    /// <summary>Offset of the encryption declaration in its volume.</summary>
    public long ByteOffset { get; }
    /// <summary>Encrypted member name, or null for archive-header encryption.</summary>
    public string? EntryName { get; }
}
