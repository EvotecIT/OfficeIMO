using System.Collections.ObjectModel;
using OfficeIMO.Drawing.Internal;
using System.Threading;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Retains the complete binary source package, compound streams, edit chain, live persist directory,
    /// and exact live persist-object bytes needed by preservation-aware saves.
    /// </summary>
    internal sealed class LegacyPptPackage {
        private const ushort RecordUserEdit = 0x0FF5;
        private const ushort RecordPersistDirectory = 0x1772;
        private readonly byte[] _originalBytes;
        private readonly byte[]? _originalEncryptedBytes;

        private LegacyPptPackage(byte[] originalBytes, OfficeCompoundFile compoundFile,
            byte[] documentStream, byte[] currentUserStream, uint currentEditOffset,
            uint documentPersistId, IReadOnlyList<LegacyPptUserEdit> userEdits,
            IReadOnlyDictionary<uint, uint> persistObjectOffsets,
            IReadOnlyDictionary<uint, LegacyPptPersistObject> persistObjects,
            bool wasEncryptedSource, int? encryptionKeySizeBits,
            bool? encryptedDocumentProperties, string? encryptionPassword,
            byte[]? originalEncryptedBytes) {
            _originalBytes = (byte[])originalBytes.Clone();
            _originalEncryptedBytes = originalEncryptedBytes == null
                ? null
                : (byte[])originalEncryptedBytes.Clone();
            CompoundFile = compoundFile;
            DocumentStream = documentStream;
            CurrentUserStream = currentUserStream;
            CurrentEditOffset = currentEditOffset;
            DocumentPersistId = documentPersistId;
            UserEdits = new ReadOnlyCollection<LegacyPptUserEdit>(userEdits.ToArray());
            PersistObjectOffsets = new ReadOnlyDictionary<uint, uint>(
                persistObjectOffsets.ToDictionary(pair => pair.Key, pair => pair.Value));
            PersistObjects = new ReadOnlyDictionary<uint, LegacyPptPersistObject>(
                persistObjects.ToDictionary(pair => pair.Key, pair => pair.Value));
            WasEncryptedSource = wasEncryptedSource;
            EncryptionKeySizeBits = encryptionKeySizeBits;
            EncryptedDocumentProperties = encryptedDocumentProperties;
            EncryptionPassword = encryptionPassword;
        }

        internal OfficeCompoundFile CompoundFile { get; }

        internal byte[] DocumentStream { get; }

        internal byte[] CurrentUserStream { get; }

        internal byte[]? PicturesStream => CompoundFile.Streams.TryGetValue("Pictures", out byte[]? stream)
            ? stream
            : null;

        internal uint CurrentEditOffset { get; }

        internal uint DocumentPersistId { get; }

        internal IReadOnlyList<LegacyPptUserEdit> UserEdits { get; }

        internal IReadOnlyDictionary<uint, uint> PersistObjectOffsets { get; }

        internal IReadOnlyDictionary<uint, LegacyPptPersistObject> PersistObjects { get; }

        internal bool WasEncryptedSource { get; }

        internal int? EncryptionKeySizeBits { get; }

        internal bool? EncryptedDocumentProperties { get; }

        internal string? EncryptionPassword { get; }

        internal int CompoundStreamCount => CompoundFile.Entries.Count(entry => entry.IsStream && !entry.IsFallback);

        internal bool HasBinarySignatureStream => CompoundFile.Entries.Any(entry =>
            entry.IsStream && !entry.IsFallback
            && string.Equals(entry.Path, "_signatures", StringComparison.OrdinalIgnoreCase));

        internal bool HasXmlSignatureStorage => CompoundFile.Entries.Any(entry =>
            !entry.IsFallback
            && (string.Equals(entry.Path, "_xmlsignatures", StringComparison.OrdinalIgnoreCase)
                || entry.Path.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)));

        internal byte[] CopyOriginalBytes() => (byte[])_originalBytes.Clone();

        internal bool TryCopyOriginalEncryptedBytes(out byte[] bytes) {
            if (_originalEncryptedBytes == null) {
                bytes = Array.Empty<byte>();
                return false;
            }
            bytes = (byte[])_originalEncryptedBytes.Clone();
            return true;
        }

        internal IReadOnlyDictionary<string, byte[]> CopyCompoundStreams() => CompoundFile.Streams.ToDictionary(
            pair => pair.Key, pair => (byte[])pair.Value.Clone(), StringComparer.OrdinalIgnoreCase);

        internal byte[] RewriteCompoundStreams(
            IReadOnlyDictionary<string, byte[]> replacementStreams,
            IReadOnlyCollection<string>? removedPaths = null) =>
            OfficeCompoundFileWriter.Rewrite(CompoundFile,
                replacementStreams, removedPaths);

        internal static LegacyPptPackage Read(byte[] bytes, LegacyPptImportOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            return Read(bytes, options,
                new LegacyPptRecordTraversalBudget(options.MaxRecordCount),
                CancellationToken.None);
        }

        internal static LegacyPptPackage Read(byte[] bytes,
            LegacyPptImportOptions options,
            LegacyPptRecordTraversalBudget recordBudget,
            CancellationToken cancellationToken = default) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (recordBudget == null) {
                throw new ArgumentNullException(nameof(recordBudget));
            }
            cancellationToken.ThrowIfCancellationRequested();
            if (!OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error)
                || compound == null) {
                throw new InvalidDataException(error ?? "The input is not a valid OLE compound file.");
            }
            if (!compound.Streams.TryGetValue("PowerPoint Document", out byte[]? documentStream)) {
                throw new InvalidDataException("The OLE compound file does not contain a PowerPoint Document stream.");
            }
            if (!compound.Streams.TryGetValue("Current User", out byte[]? currentUserStream)) {
                throw new InvalidDataException("The OLE compound file does not contain a Current User stream.");
            }
            if (documentStream.Length > options.MaxInputBytes) {
                throw new InvalidDataException($"The PowerPoint Document stream exceeds {options.MaxInputBytes} bytes.");
            }

            bool wasEncryptedSource = false;
            int? encryptionKeySizeBits = null;
            bool? encryptedDocumentProperties = null;
            string? encryptionPassword = null;
            byte[]? originalEncryptedBytes = null;
            LegacyPptCurrentUserAtom currentUser = LegacyPptCurrentUserAtom.Read(
                currentUserStream);
            if (currentUser.HeaderToken
                != LegacyPptCurrentUserAtom.UnencryptedHeaderToken) {
                originalEncryptedBytes = (byte[])bytes.Clone();
                bytes = LegacyPptRc4CryptoApi.DecryptPackage(bytes, compound,
                    options, recordBudget, cancellationToken,
                    out int keySizeBits,
                    out bool documentPropertiesEncrypted);
                cancellationToken.ThrowIfCancellationRequested();
                if (!OfficeCompoundFileReader.TryRead(bytes,
                        out compound, out error) || compound == null
                    || !compound.Streams.TryGetValue("PowerPoint Document",
                        out documentStream)
                    || !compound.Streams.TryGetValue("Current User",
                        out currentUserStream)) {
                    throw new InvalidDataException(error
                        ?? "The decrypted binary PowerPoint package is invalid.");
                }
                wasEncryptedSource = true;
                encryptionKeySizeBits = keySizeBits;
                encryptedDocumentProperties = documentPropertiesEncrypted;
                encryptionPassword = options.Password;
            }

            uint currentEditOffset = ReadCurrentEditOffset(currentUserStream);
            var liveOffsets = new Dictionary<uint, uint>();
            var edits = new List<LegacyPptUserEdit>();
            var visitedEdits = new HashSet<uint>();
            uint editOffset = currentEditOffset;
            uint documentPersistId = 0;
            while (editOffset != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!visitedEdits.Add(editOffset)) {
                    throw new InvalidDataException("The UserEditAtom chain contains a cycle.");
                }
                LegacyPptRecord edit = LegacyPptRecordReader.ReadSingle(documentStream,
                    ToBoundedOffset(editOffset, documentStream.Length,
                        "UserEditAtom"), options, recordBudget);
                if (edit.Type != RecordUserEdit || edit.PayloadLength < 20) {
                    throw new InvalidDataException("The current edit pointer does not reference a valid UserEditAtom.");
                }

                uint previousEditOffset = edit.ReadUInt32(8);
                uint persistDirectoryOffset = edit.ReadUInt32(12);
                uint editDocumentPersistId = edit.ReadUInt32(16);
                uint persistIdSeed = edit.PayloadLength >= 24 ? edit.ReadUInt32(20) : 0;
                if (documentPersistId == 0) documentPersistId = editDocumentPersistId;
                IReadOnlyDictionary<uint, uint> editOffsets = ReadPersistDirectory(
                    documentStream, persistDirectoryOffset, options,
                    recordBudget);
                foreach (KeyValuePair<uint, uint> pair in editOffsets) {
                    if (!liveOffsets.ContainsKey(pair.Key)) liveOffsets.Add(pair.Key, pair.Value);
                }
                edits.Add(new LegacyPptUserEdit(editOffset, previousEditOffset, persistDirectoryOffset,
                    editDocumentPersistId, persistIdSeed, editOffsets));
                editOffset = previousEditOffset;
            }
            if (documentPersistId == 0) {
                throw new InvalidDataException("The UserEditAtom has no document persist id.");
            }

            var persistObjects = new Dictionary<uint, LegacyPptPersistObject>();
            foreach (KeyValuePair<uint, uint> pair in liveOffsets) {
                cancellationToken.ThrowIfCancellationRequested();
                persistObjects.Add(pair.Key, ReadPersistObject(documentStream, pair.Key, pair.Value));
            }
            return new LegacyPptPackage(bytes, compound, documentStream, currentUserStream,
                currentEditOffset, documentPersistId, edits, liveOffsets,
                persistObjects, wasEncryptedSource, encryptionKeySizeBits,
                encryptedDocumentProperties, encryptionPassword,
                originalEncryptedBytes);
        }

        private static uint ReadCurrentEditOffset(byte[] currentUserStream) {
            LegacyPptCurrentUserAtom currentUser = LegacyPptCurrentUserAtom.Read(currentUserStream);
            if (currentUser.HeaderToken
                != LegacyPptCurrentUserAtom.UnencryptedHeaderToken) {
                throw new NotSupportedException("Encrypted PowerPoint 97-2003 binary presentations are not supported.");
            }
            return currentUser.CurrentEditOffset;
        }

        private static IReadOnlyDictionary<uint, uint> ReadPersistDirectory(byte[] documentStream, uint offset,
            LegacyPptImportOptions options,
            LegacyPptRecordTraversalBudget recordBudget) {
            LegacyPptRecord directory = LegacyPptRecordReader.ReadSingle(documentStream,
                ToBoundedOffset(offset, documentStream.Length,
                    "PersistDirectoryAtom"), options, recordBudget);
            if (directory.Type != RecordPersistDirectory) {
                throw new InvalidDataException("The UserEditAtom does not point to a PersistDirectoryAtom.");
            }

            var offsets = new Dictionary<uint, uint>();
            int position = 0;
            while (position < directory.PayloadLength) {
                if (directory.PayloadLength - position < 4) {
                    throw new InvalidDataException("A PersistDirectoryEntry header is truncated.");
                }
                uint packed = directory.ReadUInt32(position);
                position += 4;
                uint persistId = packed & 0x000FFFFF;
                int count = unchecked((int)(packed >> 20));
                if (count == 0 || count > (directory.PayloadLength - position) / 4) {
                    throw new InvalidDataException("A PersistDirectoryEntry has an invalid object count.");
                }
                for (int index = 0; index < count; index++) {
                    uint objectOffset = directory.ReadUInt32(position);
                    position += 4;
                    uint objectId = checked(persistId + unchecked((uint)index));
                    if (offsets.ContainsKey(objectId)) {
                        throw new InvalidDataException($"PersistDirectoryAtom contains duplicate persist id {objectId}.");
                    }
                    offsets.Add(objectId, objectOffset);
                }
            }
            return offsets;
        }

        private static LegacyPptPersistObject ReadPersistObject(byte[] documentStream, uint persistId,
            uint streamOffset) {
            int offset = ToBoundedOffset(streamOffset, documentStream.Length, $"persist object {persistId}");
            ushort recordType = ReadUInt16(documentStream, offset + 2);
            uint payloadLength = ReadUInt32(documentStream, offset + 4);
            long totalLength = 8L + payloadLength;
            if (totalLength > int.MaxValue || offset > documentStream.Length - totalLength) {
                throw new InvalidDataException($"Persist object {persistId} extends beyond the PowerPoint Document stream.");
            }
            var recordBytes = new byte[unchecked((int)totalLength)];
            Buffer.BlockCopy(documentStream, offset, recordBytes, 0, recordBytes.Length);
            return new LegacyPptPersistObject(persistId, streamOffset, recordType, recordBytes);
        }

        private static int ToBoundedOffset(uint offset, int length, string description) {
            if (offset > int.MaxValue || offset > unchecked((uint)Math.Max(0, length - 8))) {
                throw new InvalidDataException($"The {description} offset 0x{offset:X} is outside the PowerPoint Document stream.");
            }
            return unchecked((int)offset);
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) =>
            unchecked((ushort)(bytes[offset] | (bytes[offset + 1] << 8)));

        private static uint ReadUInt32(byte[] bytes, int offset) => unchecked((uint)(bytes[offset]
            | (bytes[offset + 1] << 8)
            | (bytes[offset + 2] << 16)
            | (bytes[offset + 3] << 24)));
    }
}
