namespace OfficeIMO.Email;

internal sealed class EmailSemanticEntry {
    internal EmailSemanticEntry(string path, byte[] digest, long? length) {
        Path = path;
        Digest = digest;
        Length = length;
    }
    internal string Path { get; }
    internal byte[] Digest { get; }
    internal long? Length { get; }
}

internal sealed class EmailSemanticSnapshot {
    internal EmailSemanticSnapshot(EmailSemanticFingerprint fingerprint,
        IReadOnlyDictionary<string, EmailSemanticEntry> entries) {
        Fingerprint = fingerprint;
        Entries = entries;
    }
    internal EmailSemanticFingerprint Fingerprint { get; }
    internal IReadOnlyDictionary<string, EmailSemanticEntry> Entries { get; }
}
