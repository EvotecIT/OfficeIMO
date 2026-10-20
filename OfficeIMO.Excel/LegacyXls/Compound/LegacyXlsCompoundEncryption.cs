using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Diagnostics;

namespace OfficeIMO.Excel.LegacyXls.Compound;

/// <summary>Applies the workbook password independently to each required BIFF record stream.</summary>
internal static partial class LegacyXlsCompoundEncryption {
    internal static bool IsAuxiliaryBiffStream(string name) =>
        name.StartsWith("_SX_DB_CUR/", StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "Revision Log", StringComparison.OrdinalIgnoreCase);

    internal static void AddEncryptedAuxiliaryStreams(OfficeCompoundFile source,
        Dictionary<string, byte[]> replacements, Func<byte[], byte[]> transform) {
        foreach (var item in source.Streams)
            if (IsAuxiliaryBiffStream(item.Key)) replacements[item.Key] = transform(item.Value);
    }

    internal static OfficeCompoundFile RestorePlaintextStreams(OfficeCompoundFile source,
        byte[] originalWorkbook, byte[] decodedWorkbook, LegacyXlsImportOptions options,
        List<LegacyXlsImportDiagnostic> diagnostics, IDictionary<string, bool> protectedEntries) {
        string workbookName = source.Streams.First(item => ReferenceEquals(item.Value, originalWorkbook)).Key;
        var streams = source.Streams.ToDictionary(item => item.Key, item => item.Value, StringComparer.OrdinalIgnoreCase);
        streams[workbookName] = decodedWorkbook;
        BiffRecord filePass = FindFilePass(decodedWorkbook);
        foreach (var item in source.Streams) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!IsAuxiliaryBiffStream(item.Key)) continue;
            BiffRc4RecordStreamCipher.ValidateRecordStream(item.Value);
            if (!LegacyBiffWorkbookParser.TryDecryptRecordStream(item.Value, filePass, options.Password!, diagnostics, out byte[] plaintext))
                throw new InvalidDataException("The encrypted auxiliary BIFF stream could not be decrypted: " + item.Key);
            streams[item.Key] = plaintext;
        }
        var entries = source.Entries.ToList();
        if (source.Streams.TryGetValue("encryption", out byte[]? summary)) {
            var session = BiffRc4CryptoApiEncryption.OpenSession(filePass, options.Password!);
            streams.Remove("encryption");
            entries.RemoveAll(item => string.Equals(item.Path, "encryption", StringComparison.OrdinalIgnoreCase));
            RestoreEncryptedSummary(summary, session, streams, entries, options, protectedEntries);
        }
        return new OfficeCompoundFile(streams, entries, source.RootEntry);
    }

    private static void RestoreEncryptedSummary(byte[] encrypted, OfficeBinaryRc4CryptoApiSession session,
        Dictionary<string, byte[]> streams, List<OfficeCompoundFileEntry> entries, LegacyXlsImportOptions options,
        IDictionary<string, bool> protectedEntries) {
        long decoded = 0;
        void Consume(long size) {
            options.CancellationToken.ThrowIfCancellationRequested();
            decoded = checked(decoded + size);
            if (decoded > options.MaxInputBytes) throw new InvalidDataException("Encrypted XLS compound data exceeds the decoded input limit.");
        }
        var decodedEntries = OfficeBinaryEncryptedSummary.Decrypt(encrypted, session, options.MaxInputBytes,
            (isStream, size) => { if (isStream) Consume(size); }, options.CancellationToken);
        var restored = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in decodedEntries) {
            if (item.Name.Equals("Workbook", StringComparison.OrdinalIgnoreCase)
                || item.Name.Equals("Book", StringComparison.OrdinalIgnoreCase)
                || item.Name.Equals("encryption", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Encrypted summary contains a reserved XLS entry.");
            protectedEntries.Add(item.Name, item.IsStream);
            if (item.IsStream) {
                Add(item.Name, item.Bytes, new OfficeCompoundFileEntry(item.Name, item.Name, 2, item.Bytes.Length),
                    item.Name == "\u0005DocumentSummaryInformation");
                continue;
            }
            var readOptions = new OfficeCompoundReadOptions(65536, 32768, maxStreamBytes: options.MaxInputBytes,
                maxTotalStreamBytes: options.MaxInputBytes,
                streamSizeValidator: (_, size) => Consume(size), preserveResourceLimits: true);
            if (!OfficeCompoundFileReader.TryRead(item.Bytes, readOptions, out OfficeCompoundFile? storage, out string? error) || storage == null)
                throw new InvalidDataException("Encrypted XLS storage is not a valid compound file: " + error);
            if (entries.Any(entry => string.Equals(entry.Path, item.Name, StringComparison.OrdinalIgnoreCase))) {
                ValidateStorageMirror(item.Name, storage, streams, entries, options);
                continue;
            }
            entries.Add(OfficeCompoundStorage.Repath(storage.RootEntry, item.Name, item.Name, 1));
            foreach (var entry in storage.Entries.Where(entry => !entry.IsFallback && entry.ObjectType != 5)) {
                string path = item.Name + "/" + entry.Path;
                if (entry.IsStream) Add(path, storage.Streams[entry.Path], OfficeCompoundStorage.Repath(entry, entry.Name, path, 2), false);
                else entries.Add(OfficeCompoundStorage.Repath(entry, entry.Name, path, entry.ObjectType));
                if (entries.Count > 65536) throw new InvalidDataException("Decrypted XLS compound directory exceeds the entry limit.");
            }
        }
        bool hasSummary = restored.Contains("\u0005SummaryInformation");
        bool hasDocumentSummary = restored.Contains("\u0005DocumentSummaryInformation");
        if (session.EncryptsDocumentProperties && (!hasSummary || !hasDocumentSummary))
            throw new InvalidDataException("The XLS encrypted summary must carry both document-property streams.");
        if (!session.EncryptsDocumentProperties && (hasSummary || hasDocumentSummary))
            throw new InvalidDataException("Encrypted XLS document properties disagree with the password descriptor flags.");

        void Add(string path, byte[] bytes, OfficeCompoundFileEntry entry, bool replacePlaceholder) {
            if (!restored.Add(path) || (!replacePlaceholder && streams.ContainsKey(path)))
                throw new InvalidDataException("Encrypted XLS summary contains a duplicate compound path: " + path);
            streams[path] = bytes;
            entries.RemoveAll(existing => string.Equals(existing.Path, path, StringComparison.OrdinalIgnoreCase));
            entries.Add(entry);
        }
    }

    /// <summary>Accepts the equivalent clear storage copy emitted alongside encrypted embeddings by native Excel.</summary>
    private static void ValidateStorageMirror(string name, OfficeCompoundFile storage,
        Dictionary<string, byte[]> streams, List<OfficeCompoundFileEntry> entries, LegacyXlsImportOptions options) {
        const string collision = "Encrypted XLS storage conflicts with an existing compound entry.";
        string prefix = name + "/";
        var actual = entries.Where(entry => !entry.IsFallback
            && (entry.Path.Equals(name, StringComparison.OrdinalIgnoreCase)
                || entry.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
            .ToDictionary(entry => entry.Path, StringComparer.OrdinalIgnoreCase);
        var expected = storage.Entries.Where(entry => !entry.IsFallback && entry.ObjectType != 5).ToArray();
        if (actual.Count != expected.Length + 1 || !actual.TryGetValue(name, out var root)
            || root.ObjectType != 1 || root.ClassId != storage.RootEntry.ClassId
            || root.StateBits != storage.RootEntry.StateBits)
            throw new InvalidDataException(collision);
        // Native storage serialization resets creation time and snapshots modification
        // time independently. Compare hierarchy, identities and bytes; retain the clear
        // directory's timestamps when the two payloads are otherwise equivalent.
        foreach (var entry in expected) {
            options.CancellationToken.ThrowIfCancellationRequested();
            string path = prefix + entry.Path;
            if (!actual.TryGetValue(path, out var mirror) || mirror.ObjectType != entry.ObjectType
                || mirror.ClassId != entry.ClassId || mirror.StateBits != entry.StateBits)
                throw new InvalidDataException(collision);
            if (entry.IsStream && (!streams.TryGetValue(path, out var bytes)
                || !bytes.SequenceEqual(storage.Streams[entry.Path])))
                throw new InvalidDataException(collision);
        }
        if (streams.Keys.Count(path => path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            != expected.Count(entry => entry.IsStream))
            throw new InvalidDataException(collision);
    }

    private static BiffRecord FindFilePass(byte[] bytes) {
        int offset = 0;
        while (offset <= bytes.Length - 4) {
            ushort type = BiffRecordReader.ReadUInt16(bytes, offset);
            int length = BiffRecordReader.ReadUInt16(bytes, offset + 2);
            if (length > bytes.Length - offset - 4) break;
            if (type == (ushort)BiffRecordType.FilePass) {
                var payload = new byte[length]; Buffer.BlockCopy(bytes, offset + 4, payload, 0, length);
                return new BiffRecord(type, offset, payload);
            }
            if (type == (ushort)BiffRecordType.Eof) break;
            offset += length + 4;
        }
        throw new InvalidDataException("The decrypted workbook is missing its password descriptor.");
    }
}
