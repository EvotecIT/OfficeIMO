using OfficeIMO.Drawing.Internal;
using System.Text;
using System.Threading;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Restores streams carried by an RC4 CryptoAPI EncryptedSummary stream.</summary>
    internal static class LegacyPptEncryptedSummary {
        private const string EncryptedSummaryStream = "EncryptedSummary";
        private const string SummaryInformationStream = "\u0005SummaryInformation";
        private const string DocumentSummaryInformationStream =
            "\u0005DocumentSummaryInformation";

        internal static byte[] Encrypt(OfficeCompoundFile source,
            OfficeBinaryRc4CryptoApiSession session) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (session == null) throw new ArgumentNullException(nameof(session));

            byte[] summary = GetPropertyStreamOrEmpty(source,
                SummaryInformationStream,
                OfficeOlePropertySetWriter.SummaryInformationFormatId);
            byte[] documentSummary = GetPropertyStreamOrEmpty(source,
                DocumentSummaryInformationStream,
                OfficeOlePropertySetWriter
                    .DocumentSummaryInformationFormatId);
            var streams = new[] {
                (Name: SummaryInformationStream, Bytes: summary),
                (Name: DocumentSummaryInformationStream,
                    Bytes: documentSummary)
            };
            using var output = new MemoryStream();
            output.Write(new byte[8], 0, 8);
            var descriptors = new List<EncryptedStreamDescriptor>(
                streams.Length);
            for (ushort block = 0; block < streams.Length; block++) {
                byte[] encrypted = (byte[])streams[block].Bytes.Clone();
                session.TransformInPlace(encrypted, 0, encrypted.Length,
                    block);
                int offset = checked((int)output.Position);
                output.Write(encrypted, 0, encrypted.Length);
                descriptors.Add(new EncryptedStreamDescriptor(offset,
                    encrypted.Length, block, isStream: true,
                    streams[block].Name));
            }

            int descriptorOffset = checked((int)output.Position);
            byte[] descriptorBytes = WriteDescriptors(descriptors);
            session.TransformInPlace(descriptorBytes, 0,
                descriptorBytes.Length, blockNumber: 0);
            output.Write(descriptorBytes, 0, descriptorBytes.Length);
            byte[] result = output.ToArray();
            var header = new byte[8];
            WriteUInt32(header, 0, unchecked((uint)descriptorOffset));
            WriteUInt32(header, 4,
                unchecked((uint)descriptorBytes.Length));
            session.TransformInPlace(header, 0, header.Length,
                blockNumber: 0);
            Buffer.BlockCopy(header, 0, result, 0, header.Length);
            return result;
        }

        internal static byte[] CreateEmptyDocumentSummaryInformation() =>
            OfficeOlePropertySetWriter.CreatePropertySet((
                OfficeOlePropertySetWriter
                    .DocumentSummaryInformationFormatId,
                OfficeOlePropertySetWriter.CreateSection(
                    Array.Empty<OfficeOleProperty>())));

        internal static IReadOnlyDictionary<string, byte[]> Decrypt(
            OfficeCompoundFile source,
            OfficeBinaryRc4CryptoApiSession session,
            LegacyPptImportOptions options,
            LegacyPptDecodedStorageBudget decodedStorageBudget,
            CancellationToken cancellationToken = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (session == null) throw new ArgumentNullException(nameof(session));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (decodedStorageBudget == null) {
                throw new ArgumentNullException(
                    nameof(decodedStorageBudget));
            }
            cancellationToken.ThrowIfCancellationRequested();
            if (!source.Streams.TryGetValue(EncryptedSummaryStream,
                    out byte[]? encrypted)) {
                throw new InvalidDataException(
                    "The RC4 CryptoAPI encryption header requires an EncryptedSummary stream.");
            }
            if (encrypted.Length < 12) {
                throw new InvalidDataException(
                    "The EncryptedSummary stream is truncated.");
            }

            byte[] header = CopyBytes(encrypted, 0, 8);
            session.TransformInPlace(header, 0, header.Length,
                blockNumber: 0, cancellationToken);
            uint descriptorOffsetValue = ReadUInt32(header, 0);
            uint descriptorSizeValue = ReadUInt32(header, 4);
            if (descriptorOffsetValue < 8
                || descriptorOffsetValue > int.MaxValue
                || descriptorSizeValue < 4
                || descriptorSizeValue > int.MaxValue) {
                throw new InvalidDataException(
                    "The EncryptedSummary descriptor location is invalid.");
            }
            int descriptorOffset = unchecked((int)descriptorOffsetValue);
            int descriptorSize = unchecked((int)descriptorSizeValue);
            if (descriptorOffset > encrypted.Length - descriptorSize) {
                throw new InvalidDataException(
                    "The EncryptedSummary descriptor array extends beyond the stream.");
            }

            byte[] descriptors = CopyBytes(encrypted, descriptorOffset,
                descriptorSize);
            session.TransformInPlace(descriptors, 0, descriptors.Length,
                blockNumber: 0, cancellationToken);
            IReadOnlyList<EncryptedStreamDescriptor> entries =
                ReadDescriptors(descriptors, descriptorOffset,
                    cancellationToken);
            var replacements = new Dictionary<string, byte[]>(
                StringComparer.OrdinalIgnoreCase);
            foreach (EncryptedStreamDescriptor entry in entries) {
                cancellationToken.ThrowIfCancellationRequested();
                if (entry.StreamOffset < 8
                    || entry.StreamOffset > descriptorOffset
                    || entry.StreamSize > descriptorOffset - entry.StreamOffset) {
                    throw new InvalidDataException(
                        $"EncryptedSummary entry '{entry.Name}' has an invalid data range.");
                }
                byte[] data = CopyBytes(encrypted, entry.StreamOffset,
                    entry.StreamSize);
                session.TransformInPlace(data, 0, data.Length,
                    entry.BlockNumber, cancellationToken);
                if (entry.IsStream) {
                    AddReplacement(replacements, entry.Name, data);
                } else {
                    AddStorageReplacements(replacements, entry.Name, data,
                        options, decodedStorageBudget, cancellationToken);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (!replacements.ContainsKey(SummaryInformationStream)
                || !replacements.ContainsKey(DocumentSummaryInformationStream)) {
                throw new InvalidDataException(
                    "The EncryptedSummary stream must carry both Office document-property streams.");
            }
            return replacements;
        }

        private static byte[] GetPropertyStreamOrEmpty(
            OfficeCompoundFile source, string name, Guid formatId) {
            if (source.Streams.TryGetValue(name, out byte[]? bytes)) {
                return bytes;
            }
            return OfficeOlePropertySetWriter.CreatePropertySet((formatId,
                OfficeOlePropertySetWriter.CreateSection(
                    Array.Empty<OfficeOleProperty>())));
        }

        private static byte[] WriteDescriptors(
            IReadOnlyList<EncryptedStreamDescriptor> entries) {
            using var output = new MemoryStream();
            WriteUInt32(output, unchecked((uint)entries.Count));
            foreach (EncryptedStreamDescriptor entry in entries) {
                if (entry.Name.Length == 0 || entry.Name.Length > byte.MaxValue) {
                    throw new InvalidDataException(
                        "An EncryptedSummary stream name must contain 1 through 255 characters.");
                }
                WriteUInt32(output, unchecked((uint)entry.StreamOffset));
                WriteUInt32(output, unchecked((uint)entry.StreamSize));
                WriteUInt16(output, entry.BlockNumber);
                output.WriteByte(unchecked((byte)entry.Name.Length));
                output.WriteByte(entry.IsStream ? (byte)0x01 : (byte)0x00);
                WriteUInt32(output, 0);
                byte[] nameBytes = Encoding.Unicode.GetBytes(entry.Name);
                output.Write(nameBytes, 0, nameBytes.Length);
                WriteUInt16(output, 0);
            }
            return output.ToArray();
        }

        private static IReadOnlyList<EncryptedStreamDescriptor>
            ReadDescriptors(byte[] bytes, int descriptorOffset,
                CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            uint countValue = ReadUInt32(bytes, 0);
            if (countValue > int.MaxValue
                || countValue > unchecked((uint)((bytes.Length - 4) / 18))) {
                throw new InvalidDataException(
                    "The EncryptedSummary descriptor count is invalid.");
            }
            int count = unchecked((int)countValue);
            var entries = new List<EncryptedStreamDescriptor>(count);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int position = 4;
            for (int index = 0; index < count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (position > bytes.Length - 18) {
                    throw new InvalidDataException(
                        "An EncryptedSummary descriptor is truncated.");
                }
                uint offsetValue = ReadUInt32(bytes, position);
                uint sizeValue = ReadUInt32(bytes, position + 4);
                ushort block = ReadUInt16(bytes, position + 8);
                int nameLength = bytes[position + 10];
                byte flags = bytes[position + 11];
                position += 16;
                int nameByteLength = checked(nameLength * 2);
                if (nameLength == 0
                    || position > bytes.Length - nameByteLength - 2) {
                    throw new InvalidDataException(
                        "An EncryptedSummary descriptor has an invalid stream name.");
                }
                string name = Encoding.Unicode.GetString(bytes, position,
                    nameByteLength);
                position += nameByteLength;
                if (ReadUInt16(bytes, position) != 0
                    || name.IndexOf('\0') >= 0
                    || name.IndexOf('/') >= 0
                    || name.IndexOf('\\') >= 0
                    || !names.Add(name)) {
                    throw new InvalidDataException(
                        "An EncryptedSummary descriptor has an invalid or duplicate stream name.");
                }
                position += 2;
                if (offsetValue > int.MaxValue || sizeValue > int.MaxValue
                    || offsetValue >= unchecked((uint)descriptorOffset)) {
                    throw new InvalidDataException(
                        $"EncryptedSummary entry '{name}' has an invalid data location.");
                }
                entries.Add(new EncryptedStreamDescriptor(
                    unchecked((int)offsetValue),
                    unchecked((int)sizeValue), block,
                    (flags & 0x01) != 0, name));
            }
            if (position != bytes.Length) {
                throw new InvalidDataException(
                    "The EncryptedSummary descriptor array has trailing data.");
            }
            return entries;
        }

        internal static void AddStorageReplacements(
            IDictionary<string, byte[]> replacements, string storageName,
            byte[] bytes, LegacyPptImportOptions options,
            LegacyPptDecodedStorageBudget decodedStorageBudget,
            CancellationToken cancellationToken = default) {
            if (replacements == null) {
                throw new ArgumentNullException(nameof(replacements));
            }
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (decodedStorageBudget == null) {
                throw new ArgumentNullException(
                    nameof(decodedStorageBudget));
            }
            cancellationToken.ThrowIfCancellationRequested();
            int maxDirectoryEntries = Math.Max(1,
                Math.Min(65536, options.MaxRecordCount));
            int maxStreamCount = Math.Max(1,
                Math.Min(32768, options.MaxRecordCount));
            var readOptions = new OfficeCompoundReadOptions(
                maxDirectoryEntries, maxStreamCount,
                maxStreamBytes: int.MaxValue,
                maxTotalStreamBytes: long.MaxValue,
                streamSizeValidator: (_, size) => {
                    cancellationToken.ThrowIfCancellationRequested();
                    decodedStorageBudget.Consume(checked((int)size));
                });
            if (!OfficeCompoundFileReader.TryRead(bytes, readOptions,
                    out OfficeCompoundFile? storage, out string? error)
                || storage == null) {
                decodedStorageBudget.ThrowIfExceeded();
                throw new InvalidDataException(
                    $"EncryptedSummary storage '{storageName}' is not a valid compound file: {error}");
            }
            int streamCount = 0;
            foreach (OfficeCompoundFileEntry entry in storage.Entries) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!entry.IsStream || entry.IsFallback) continue;
                AddReplacement(replacements,
                    storageName + "/" + entry.Path,
                    storage.Streams[entry.Path]);
                streamCount++;
            }
            if (streamCount == 0) {
                throw new InvalidDataException(
                    $"EncryptedSummary storage '{storageName}' contains no streams.");
            }
        }

        private static void AddReplacement(
            IDictionary<string, byte[]> replacements, string path,
            byte[] bytes) {
            if (replacements.ContainsKey(path)) {
                throw new InvalidDataException(
                    $"EncryptedSummary contains duplicate compound path '{path}'.");
            }
            replacements.Add(path, bytes);
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte(unchecked((byte)value));
            stream.WriteByte(unchecked((byte)(value >> 8)));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte(unchecked((byte)value));
            stream.WriteByte(unchecked((byte)(value >> 8)));
            stream.WriteByte(unchecked((byte)(value >> 16)));
            stream.WriteByte(unchecked((byte)(value >> 24)));
        }

        private static byte[] CopyBytes(byte[] source, int offset,
            int length) {
            var result = new byte[length];
            Buffer.BlockCopy(source, offset, result, 0, length);
            return result;
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) =>
            unchecked((ushort)(bytes[offset] | bytes[offset + 1] << 8));

        private static uint ReadUInt32(byte[] bytes, int offset) =>
            unchecked((uint)(bytes[offset] | bytes[offset + 1] << 8
                | bytes[offset + 2] << 16 | bytes[offset + 3] << 24));

        private static void WriteUInt32(byte[] bytes, int offset,
            uint value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
            bytes[offset + 2] = unchecked((byte)(value >> 16));
            bytes[offset + 3] = unchecked((byte)(value >> 24));
        }

        private sealed class EncryptedStreamDescriptor {
            internal EncryptedStreamDescriptor(int streamOffset,
                int streamSize, ushort blockNumber, bool isStream,
                string name) {
                StreamOffset = streamOffset;
                StreamSize = streamSize;
                BlockNumber = blockNumber;
                IsStream = isStream;
                Name = name;
            }

            internal int StreamOffset { get; }
            internal int StreamSize { get; }
            internal ushort BlockNumber { get; }
            internal bool IsStream { get; }
            internal string Name { get; }
        }
    }
}
