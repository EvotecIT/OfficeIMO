using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Core.Internal {
    /// <summary>Decodes the MS-OFFCRYPTO encrypted stream and storage descriptor container.</summary>
    internal static class OfficeBinaryEncryptedSummary {
        internal sealed class Entry {
            internal Entry(string name, bool isStream, byte[] bytes) { Name = name; IsStream = isStream; Bytes = bytes; }
            internal string Name { get; }
            internal bool IsStream { get; }
            internal byte[] Bytes { get; }
        }

        internal static IReadOnlyList<Entry> Decrypt(byte[] encrypted,
            OfficeBinaryRc4CryptoApiSession session, long maximumDecodedBytes,
            Action<bool, int>? beforeDecode = null, CancellationToken cancellationToken = default, int maximumEntries = 65536) {
            if (encrypted == null) throw new ArgumentNullException(nameof(encrypted));
            if (session == null) throw new ArgumentNullException(nameof(session));
            if (maximumDecodedBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumDecodedBytes));
            if (maximumEntries < 1) throw new ArgumentOutOfRangeException(nameof(maximumEntries));
            cancellationToken.ThrowIfCancellationRequested();
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
                ReadDescriptors(descriptors, descriptorOffset, maximumEntries,
                    cancellationToken);
            var result = new List<Entry>(entries.Count);
            long total = 0;
            foreach (EncryptedStreamDescriptor entry in entries) {
                cancellationToken.ThrowIfCancellationRequested();
                if (entry.StreamOffset < 8 || entry.StreamOffset > descriptorOffset
                    || entry.StreamSize > descriptorOffset - entry.StreamOffset)
                    throw new InvalidDataException("An encrypted summary entry has an invalid data range.");
                total = checked(total + entry.StreamSize);
                if (total > maximumDecodedBytes)
                    throw new InvalidDataException("Encrypted summary data exceeds the decoded byte limit.");
                beforeDecode?.Invoke(entry.IsStream, entry.StreamSize);
                byte[] data = CopyBytes(encrypted, entry.StreamOffset, entry.StreamSize);
                session.TransformInPlace(data, 0, data.Length, entry.BlockNumber, cancellationToken);
                result.Add(new Entry(entry.Name, entry.IsStream, data));
            }
            return result;
        }

        private static IReadOnlyList<EncryptedStreamDescriptor>
            ReadDescriptors(byte[] bytes, int descriptorOffset, int maximumEntries,
                CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            uint countValue = ReadUInt32(bytes, 0);
            if (countValue > maximumEntries
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
