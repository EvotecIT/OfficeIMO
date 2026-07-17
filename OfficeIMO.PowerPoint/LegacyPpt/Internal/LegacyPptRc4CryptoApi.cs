using OfficeIMO.Drawing.Internal;
using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Reads and writes record-scoped RC4 CryptoAPI encryption for binary PowerPoint packages.</summary>
    internal static class LegacyPptRc4CryptoApi {
        private const ushort RecordUserEdit = 0x0FF5;
        private const ushort RecordPersistDirectory = 0x1772;
        private const ushort RecordCryptSession10 = 0x2F14;
        private const int DefaultKeySizeBits = 128;

        internal static byte[] DecryptPackage(byte[] sourceBytes,
            OfficeCompoundFile source, LegacyPptImportOptions options,
            LegacyPptRecordTraversalBudget recordBudget,
            LegacyPptDecodedStorageBudget decodedStorageBudget,
            CancellationToken cancellationToken,
            out int keySizeBits, out bool encryptedDocumentProperties) {
            if (sourceBytes == null) throw new ArgumentNullException(nameof(sourceBytes));
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (recordBudget == null) {
                throw new ArgumentNullException(nameof(recordBudget));
            }
            if (decodedStorageBudget == null) {
                throw new ArgumentNullException(
                    nameof(decodedStorageBudget));
            }
            if (options.Password == null) {
                throw new CryptographicException(
                    "The binary PowerPoint presentation is encrypted. Provide LegacyPptImportOptions.Password or use PowerPointPresentation.LoadEncrypted.");
            }
            cancellationToken.ThrowIfCancellationRequested();
            byte[] documentStream = GetRequiredStream(source,
                "PowerPoint Document");
            byte[] currentUserStream = GetRequiredStream(source,
                "Current User");
            LegacyPptCurrentUserAtom currentUser =
                LegacyPptCurrentUserAtom.Read(currentUserStream);
            if (currentUser.HeaderToken
                != LegacyPptCurrentUserAtom.EncryptedHeaderToken) {
                throw new InvalidDataException(
                    "The Current User stream does not advertise RC4 CryptoAPI encryption.");
            }

            LegacyPptRecord edit = LegacyPptRecordReader.ReadSingle(
                documentStream, ToBoundedOffset(currentUser.CurrentEditOffset,
                    documentStream.Length, "UserEditAtom"), options,
                recordBudget);
            if (edit.Type != RecordUserEdit || edit.PayloadLength < 32
                || edit.ReadUInt32(8) != 0) {
                throw new InvalidDataException(
                    "An encrypted binary PowerPoint stream must contain one complete UserEditAtom.");
            }
            uint directoryOffset = edit.ReadUInt32(12);
            IReadOnlyDictionary<uint, uint> offsets = ReadPersistDirectory(
                documentStream, directoryOffset, options, recordBudget);
            uint encryptionPersistId = edit.ReadUInt32(28);
            if (encryptionPersistId == 0
                || !offsets.TryGetValue(encryptionPersistId,
                    out uint encryptionOffset)) {
                throw new InvalidDataException(
                    "The encrypted presentation does not reference a valid CryptSession10Container.");
            }
            byte[] encryptionRecord = ReadPlainRecord(documentStream,
                encryptionOffset, RecordCryptSession10,
                "CryptSession10Container");
            byte[] encryptionHeader = CopyBytes(encryptionRecord, 8,
                encryptionRecord.Length - 8);
            OfficeBinaryRc4CryptoApiSession session =
                OfficeBinaryRc4CryptoApiSession.Open(encryptionHeader,
                    options.Password);
            cancellationToken.ThrowIfCancellationRequested();
            keySizeBits = session.KeySizeBits;
            encryptedDocumentProperties = session.EncryptsDocumentProperties;

            EncryptedPersistObject[] encryptedObjects = offsets
                .Where(pair => pair.Key != encryptionPersistId)
                .Select(pair => InspectEncryptedPersistObject(documentStream,
                    pair.Key, pair.Value, session, cancellationToken))
                .OrderBy(item => item.Offset)
                .ThenBy(item => item.PersistId)
                .ToArray();
            ValidatePersistObjectRanges(encryptedObjects,
                new EncryptedPersistObject(encryptionPersistId,
                    ToBoundedOffset(encryptionOffset,
                        documentStream.Length,
                        "CryptSession10Container"),
                    encryptionRecord.Length));
            foreach (EncryptedPersistObject item in encryptedObjects) {
                decodedStorageBudget.Consume(item.RecordLength);
            }

            var plainObjects = new SortedDictionary<uint, byte[]>();
            foreach (EncryptedPersistObject item in encryptedObjects) {
                cancellationToken.ThrowIfCancellationRequested();
                plainObjects.Add(item.PersistId, DecryptPersistObject(
                    documentStream, item, session, cancellationToken));
            }
            cancellationToken.ThrowIfCancellationRequested();
            byte[] plainDocument = BuildDocumentStream(plainObjects, edit,
                encryptionPersistId: null, out uint plainEditOffset);
            var replacements = new Dictionary<string, byte[]>(
                StringComparer.OrdinalIgnoreCase) {
                ["PowerPoint Document"] = plainDocument,
                ["Current User"] = PatchCurrentUser(currentUserStream,
                    LegacyPptCurrentUserAtom.UnencryptedHeaderToken,
                    plainEditOffset)
            };
            if (source.Streams.TryGetValue("Pictures", out byte[]? pictures)) {
                replacements["Pictures"] = TransformPictures(pictures,
                    session, encrypting: false,
                    cancellationToken: cancellationToken);
            }
            if (session.EncryptsDocumentProperties) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (KeyValuePair<string, byte[]> replacement
                         in LegacyPptEncryptedSummary.Decrypt(source, session,
                             options, decodedStorageBudget,
                             cancellationToken)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    replacements[replacement.Key] = replacement.Value;
                }
            }
            cancellationToken.ThrowIfCancellationRequested();
            return OfficeCompoundFileWriter.Rewrite(source, replacements,
                new[] { "EncryptedSummary" });
        }

        internal static byte[] EncryptPackage(byte[] plainBytes,
            string password, int keySizeBits = DefaultKeySizeBits,
            bool encryptDocumentProperties = true) {
            if (plainBytes == null) throw new ArgumentNullException(nameof(plainBytes));
            if (password == null) throw new ArgumentNullException(nameof(password));
            LegacyPptPackage package = LegacyPptPackage.Read(plainBytes,
                new LegacyPptImportOptions());
            uint encryptionPersistId = package.PersistObjectOffsets.Count == 0
                ? 1U
                : checked(package.PersistObjectOffsets.Keys.Max() + 1U);
            if (encryptionPersistId > 0x000FFFFE) {
                throw new NotSupportedException(
                    "The presentation has no available persist identifier for its encryption session.");
            }
            OfficeBinaryRc4CryptoApiSession session =
                OfficeBinaryRc4CryptoApiSession.Create(password, keySizeBits,
                    encryptDocumentProperties,
                    out byte[] encryptionHeader);

            var encryptedObjects = new SortedDictionary<uint, byte[]>();
            foreach (KeyValuePair<uint, LegacyPptPersistObject> pair
                     in package.PersistObjects) {
                byte[] encrypted = (byte[])pair.Value.RecordBytes.Clone();
                session.TransformInPlace(encrypted, 0, encrypted.Length,
                    pair.Key);
                encryptedObjects.Add(pair.Key, encrypted);
            }
            encryptedObjects.Add(encryptionPersistId,
                BuildRecord(version: 0x0F, instance: 0,
                    RecordCryptSession10, encryptionHeader));
            LegacyPptRecord sourceEdit = LegacyPptRecordReader.ReadSingle(
                package.DocumentStream, checked((int)package.CurrentEditOffset),
                new LegacyPptImportOptions());
            byte[] encryptedDocument = BuildDocumentStream(encryptedObjects,
                sourceEdit, encryptionPersistId, out uint encryptedEditOffset);
            var replacements = new Dictionary<string, byte[]>(
                StringComparer.OrdinalIgnoreCase) {
                ["PowerPoint Document"] = encryptedDocument,
                ["Current User"] = PatchCurrentUser(package.CurrentUserStream,
                    LegacyPptCurrentUserAtom.EncryptedHeaderToken,
                    encryptedEditOffset)
            };
            if (package.PicturesStream != null) {
                replacements["Pictures"] = TransformPictures(
                    package.PicturesStream, session, encrypting: true);
            }
            if (encryptDocumentProperties) {
                replacements["EncryptedSummary"] =
                    LegacyPptEncryptedSummary.Encrypt(
                        package.CompoundFile, session);
                replacements["\u0005DocumentSummaryInformation"] =
                    LegacyPptEncryptedSummary
                        .CreateEmptyDocumentSummaryInformation();
                return package.RewriteCompoundStreams(replacements,
                    new[] { "\u0005SummaryInformation" });
            }
            return package.RewriteCompoundStreams(replacements,
                new[] { "EncryptedSummary" });
        }

        private static EncryptedPersistObject InspectEncryptedPersistObject(
            byte[] documentStream,
            uint persistId, uint streamOffset,
            OfficeBinaryRc4CryptoApiSession session,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            int offset = ToBoundedOffset(streamOffset, documentStream.Length,
                $"persist object {persistId}");
            byte[] header = CopyBytes(documentStream, offset, 8);
            session.TransformInPlace(header, 0, header.Length, persistId,
                cancellationToken);
            uint payloadLength = ReadUInt32(header, 4);
            long recordLength = 8L + payloadLength;
            if (recordLength > int.MaxValue
                || offset > documentStream.Length - recordLength) {
                throw new InvalidDataException(
                    $"Encrypted persist object {persistId} has an invalid record length.");
            }
            return new EncryptedPersistObject(persistId, offset,
                unchecked((int)recordLength));
        }

        private static void ValidatePersistObjectRanges(
            IReadOnlyCollection<EncryptedPersistObject> encryptedObjects,
            EncryptedPersistObject encryptionObject) {
            EncryptedPersistObject? previous = null;
            foreach (EncryptedPersistObject current in encryptedObjects
                         .Append(encryptionObject)
                         .OrderBy(item => item.Offset)
                         .ThenBy(item => item.PersistId)) {
                if (previous.HasValue
                    && current.Offset < previous.Value.EndOffset) {
                    throw new InvalidDataException(
                        $"Encrypted persist object {current.PersistId} overlaps persist object {previous.Value.PersistId} in the PowerPoint Document stream.");
                }
                previous = current;
            }
        }

        private static byte[] DecryptPersistObject(byte[] documentStream,
            EncryptedPersistObject item,
            OfficeBinaryRc4CryptoApiSession session,
            CancellationToken cancellationToken) {
            byte[] record = CopyBytes(documentStream, item.Offset,
                item.RecordLength);
            session.TransformInPlace(record, 0, record.Length, item.PersistId,
                cancellationToken);
            return record;
        }

        private readonly struct EncryptedPersistObject {
            internal EncryptedPersistObject(uint persistId, int offset,
                int recordLength) {
                PersistId = persistId;
                Offset = offset;
                RecordLength = recordLength;
            }

            internal uint PersistId { get; }

            internal int Offset { get; }

            internal int RecordLength { get; }

            internal int EndOffset => checked(Offset + RecordLength);
        }

        private static byte[] BuildDocumentStream(
            IReadOnlyDictionary<uint, byte[]> persistObjects,
            LegacyPptRecord sourceEdit, uint? encryptionPersistId,
            out uint editOffset) {
            using var output = new MemoryStream();
            var offsets = new SortedDictionary<uint, uint>();
            foreach (KeyValuePair<uint, byte[]> pair in persistObjects
                         .OrderBy(pair => pair.Key)) {
                offsets.Add(pair.Key, checked((uint)output.Position));
                output.Write(pair.Value, 0, pair.Value.Length);
            }
            uint directoryOffset = checked((uint)output.Position);
            byte[] directory = BuildPersistDirectory(offsets);
            output.Write(directory, 0, directory.Length);
            editOffset = checked((uint)output.Position);
            byte[] edit = BuildUserEdit(sourceEdit, directoryOffset,
                offsets.Keys.Count == 0 ? 0U : offsets.Keys.Max(),
                encryptionPersistId);
            output.Write(edit, 0, edit.Length);
            return output.ToArray();
        }

        private static byte[] BuildUserEdit(LegacyPptRecord source,
            uint directoryOffset, uint maximumPersistId,
            uint? encryptionPersistId) {
            if (source.Type != RecordUserEdit || source.PayloadLength < 28) {
                throw new InvalidDataException(
                    "The source UserEditAtom is incomplete.");
            }
            var payload = new byte[encryptionPersistId.HasValue ? 32 : 28];
            byte[] sourceBytes = source.CopyRecordBytes();
            Buffer.BlockCopy(sourceBytes, 8, payload, 0, 8);
            WriteUInt32(payload, 8, 0);
            WriteUInt32(payload, 12, directoryOffset);
            WriteUInt32(payload, 16, source.ReadUInt32(16));
            WriteUInt32(payload, 20, Math.Max(source.ReadUInt32(20),
                maximumPersistId));
            Buffer.BlockCopy(sourceBytes, 32, payload, 24, 4);
            if (encryptionPersistId.HasValue) {
                WriteUInt32(payload, 28, encryptionPersistId.Value);
            }
            return BuildRecord(version: 0, instance: 0, RecordUserEdit,
                payload);
        }

        private static byte[] BuildPersistDirectory(
            IReadOnlyDictionary<uint, uint> offsets) {
            var payload = new List<byte>();
            KeyValuePair<uint, uint>[] entries = offsets
                .OrderBy(pair => pair.Key).ToArray();
            for (int index = 0; index < entries.Length;) {
                int count = 1;
                while (index + count < entries.Length && count < 0x0FFF
                       && entries[index + count].Key
                       == entries[index].Key + unchecked((uint)count)) {
                    count++;
                }
                AppendUInt32(payload, (unchecked((uint)count) << 20)
                    | entries[index].Key);
                for (int item = 0; item < count; item++) {
                    AppendUInt32(payload, entries[index + item].Value);
                }
                index += count;
            }
            return BuildRecord(version: 0, instance: 0,
                RecordPersistDirectory, payload.ToArray());
        }

        private static IReadOnlyDictionary<uint, uint> ReadPersistDirectory(
            byte[] documentStream, uint offset, LegacyPptImportOptions options,
            LegacyPptRecordTraversalBudget recordBudget) {
            LegacyPptRecord directory = LegacyPptRecordReader.ReadSingle(
                documentStream, ToBoundedOffset(offset, documentStream.Length,
                    "PersistDirectoryAtom"), options, recordBudget);
            if (directory.Type != RecordPersistDirectory) {
                throw new InvalidDataException(
                    "The encrypted UserEditAtom does not reference a PersistDirectoryAtom.");
            }
            var offsets = new Dictionary<uint, uint>();
            int position = 0;
            while (position < directory.PayloadLength) {
                if (directory.PayloadLength - position < 4) {
                    throw new InvalidDataException(
                        "An encrypted PersistDirectoryEntry is truncated.");
                }
                uint packed = directory.ReadUInt32(position);
                position += 4;
                uint persistId = packed & 0x000FFFFF;
                int count = unchecked((int)(packed >> 20));
                if (count == 0
                    || count > (directory.PayloadLength - position) / 4) {
                    throw new InvalidDataException(
                        "An encrypted PersistDirectoryEntry has an invalid object count.");
                }
                recordBudget.Consume(count);
                for (int index = 0; index < count; index++) {
                    uint objectId = checked(persistId
                        + unchecked((uint)index));
                    if (offsets.ContainsKey(objectId)) {
                        throw new InvalidDataException(
                            $"The encrypted persist directory contains duplicate id {objectId}.");
                    }
                    offsets.Add(objectId, directory.ReadUInt32(position));
                    position += 4;
                }
            }
            return offsets;
        }

        private static byte[] TransformPictures(byte[] source,
            OfficeBinaryRc4CryptoApiSession session, bool encrypting,
            CancellationToken cancellationToken = default) {
            byte[] pictures = (byte[])source.Clone();
            int position = 0;
            while (position < pictures.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                if (pictures.Length - position < 8) {
                    throw new InvalidDataException(
                        "The encrypted Pictures stream has a truncated record header.");
                }
                int recordStart = position;
                ushort versionAndInstance = encrypting
                    ? ReadUInt16(pictures, position)
                    : (ushort)0;
                ushort recordType = encrypting
                    ? ReadUInt16(pictures, position + 2)
                    : (ushort)0;
                uint payloadLength = encrypting
                    ? ReadUInt32(pictures, position + 4)
                    : 0;
                TransformPictureField(pictures, position, 8, session,
                    cancellationToken);
                if (!encrypting) {
                    versionAndInstance = ReadUInt16(pictures, position);
                    recordType = ReadUInt16(pictures, position + 2);
                    payloadLength = ReadUInt32(pictures, position + 4);
                }
                long end = position + 8L + payloadLength;
                if (end > pictures.Length) {
                    throw new InvalidDataException(
                        "The encrypted Pictures stream contains an invalid record length.");
                }
                position += 8;
                TransformPicturePayload(pictures, ref position,
                    checked((int)end), versionAndInstance, recordType,
                    session, encrypting, cancellationToken);
                if (position != end) {
                    throw new InvalidDataException(
                        $"Picture record at offset 0x{recordStart:X} was not transformed completely.");
                }
            }
            return pictures;
        }

        private static void TransformPicturePayload(byte[] pictures,
            ref int position, int end, ushort versionAndInstance,
            ushort recordType, OfficeBinaryRc4CryptoApiSession session,
            bool encrypting,
            CancellationToken cancellationToken) {
            if (recordType == 0xF007) {
                int[] parts = { 1, 1, 16, 2, 4, 4, 4, 1, 1, 1, 1 };
                int nameLength = encrypting
                    ? ReadPictureNameLength(pictures, position, end)
                    : 0;
                foreach (int part in parts) {
                    TransformBoundedPictureField(pictures, ref position,
                        end, part, session, cancellationToken);
                }
                if (!encrypting) {
                    nameLength = pictures[position - 3];
                }
                if (nameLength > 0) {
                    TransformBoundedPictureField(pictures, ref position,
                        end, nameLength, session, cancellationToken);
                }
                if (position == end) return;
                if (end - position < 8) {
                    throw new InvalidDataException(
                        "An embedded picture record header is truncated.");
                }
                versionAndInstance = encrypting
                    ? ReadUInt16(pictures, position)
                    : (ushort)0;
                recordType = encrypting
                    ? ReadUInt16(pictures, position + 2)
                    : (ushort)0;
                TransformPictureField(pictures, position, 8, session,
                    cancellationToken);
                if (!encrypting) {
                    versionAndInstance = ReadUInt16(pictures, position);
                    recordType = ReadUInt16(pictures, position + 2);
                }
                position += 8;
            }

            int instance = versionAndInstance >> 4;
            int uidCount = instance == 0x217 || instance == 0x3D5
                || instance == 0x46B || instance == 0x543
                || instance == 0x6E1 || instance == 0x6E3
                || instance == 0x6E5 || instance == 0x7A9 ? 2 : 1;
            for (int index = 0; index < uidCount; index++) {
                TransformBoundedPictureField(pictures, ref position,
                    end, 16, session, cancellationToken);
            }
            int metadataLength = recordType == 0xF01A
                || recordType == 0xF01B || recordType == 0xF01C ? 34 : 1;
            TransformBoundedPictureField(pictures, ref position, end,
                metadataLength, session, cancellationToken);
            TransformBoundedPictureField(pictures, ref position, end,
                end - position, session, cancellationToken);
        }

        private static int ReadPictureNameLength(byte[] pictures,
            int payloadOffset, int end) {
            if (end - payloadOffset < 36) {
                throw new InvalidDataException(
                    "The Pictures stream contains a truncated File BLIP Store Entry.");
            }
            return pictures[payloadOffset + 33];
        }

        private static void TransformBoundedPictureField(byte[] pictures,
            ref int position, int end, int length,
            OfficeBinaryRc4CryptoApiSession session,
            CancellationToken cancellationToken) {
            if (length < 0 || position > end - length) {
                throw new InvalidDataException(
                    "The encrypted picture field extends beyond its record.");
            }
            TransformPictureField(pictures, position, length, session,
                cancellationToken);
            position += length;
        }

        private static void TransformPictureField(byte[] pictures, int offset,
            int length, OfficeBinaryRc4CryptoApiSession session,
            CancellationToken cancellationToken) =>
            session.TransformInPlace(pictures, offset, length,
                blockNumber: 0,
                cancellationToken: cancellationToken);

        private static byte[] ReadPlainRecord(byte[] stream, uint streamOffset,
            ushort expectedType, string description) {
            int offset = ToBoundedOffset(streamOffset, stream.Length,
                description);
            ushort type = ReadUInt16(stream, offset + 2);
            uint payloadLength = ReadUInt32(stream, offset + 4);
            long length = 8L + payloadLength;
            if (type != expectedType || length > int.MaxValue
                || offset > stream.Length - length) {
                throw new InvalidDataException(
                    $"The {description} record is invalid.");
            }
            return CopyBytes(stream, offset, unchecked((int)length));
        }

        private static byte[] PatchCurrentUser(byte[] source, uint token,
            uint editOffset) {
            byte[] patched = (byte[])source.Clone();
            _ = LegacyPptCurrentUserAtom.Read(patched);
            WriteUInt32(patched, 12, token);
            WriteUInt32(patched, 16, editOffset);
            return patched;
        }

        private static byte[] GetRequiredStream(OfficeCompoundFile compound,
            string name) => compound.Streams.TryGetValue(name,
            out byte[]? stream) ? stream : throw new InvalidDataException(
            $"The OLE compound file does not contain a {name} stream.");

        private static byte[] BuildRecord(byte version, ushort instance,
            ushort type, byte[] payload) {
            var bytes = new byte[checked(8 + payload.Length)];
            WriteUInt16(bytes, 0,
                unchecked((ushort)((instance << 4) | version)));
            WriteUInt16(bytes, 2, type);
            WriteUInt32(bytes, 4, unchecked((uint)payload.Length));
            Buffer.BlockCopy(payload, 0, bytes, 8, payload.Length);
            return bytes;
        }

        private static int ToBoundedOffset(uint offset, int length,
            string description) {
            if (offset > int.MaxValue
                || offset > unchecked((uint)Math.Max(0, length - 8))) {
                throw new InvalidDataException(
                    $"The {description} offset 0x{offset:X} is outside the PowerPoint Document stream.");
            }
            return unchecked((int)offset);
        }

        private static byte[] CopyBytes(byte[] source, int offset, int length) {
            var result = new byte[length];
            Buffer.BlockCopy(source, offset, result, 0, length);
            return result;
        }

        private static void AppendUInt32(ICollection<byte> bytes,
            uint value) {
            bytes.Add(unchecked((byte)value));
            bytes.Add(unchecked((byte)(value >> 8)));
            bytes.Add(unchecked((byte)(value >> 16)));
            bytes.Add(unchecked((byte)(value >> 24)));
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) =>
            unchecked((ushort)(bytes[offset] | bytes[offset + 1] << 8));

        private static uint ReadUInt32(byte[] bytes, int offset) =>
            unchecked((uint)(bytes[offset] | bytes[offset + 1] << 8
                | bytes[offset + 2] << 16 | bytes[offset + 3] << 24));

        private static void WriteUInt16(byte[] bytes, int offset,
            ushort value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
        }

        private static void WriteUInt32(byte[] bytes, int offset, uint value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
            bytes[offset + 2] = unchecked((byte)(value >> 16));
            bytes[offset + 3] = unchecked((byte)(value >> 24));
        }
    }
}
