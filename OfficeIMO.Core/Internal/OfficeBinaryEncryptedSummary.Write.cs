using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeBinaryEncryptedSummary {
        /// <summary>Encodes the shared Office binary stream/storage container with independent RC4 block keys.</summary>
        internal static byte[] Encrypt(IReadOnlyList<Entry> entries, OfficeBinaryRc4CryptoApiSession session,
            long maximumBytes = int.MaxValue, CancellationToken cancellationToken = default) {
            if (entries == null) throw new ArgumentNullException(nameof(entries));
            if (session == null) throw new ArgumentNullException(nameof(session));
            if (maximumBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            if (entries.Count > 65536) throw new InvalidDataException("Encrypted summary has too many entries.");
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            long required = 12;
            foreach (var entry in entries) {
                cancellationToken.ThrowIfCancellationRequested();
                if (entry.Name.Length == 0 || entry.Name.Length > byte.MaxValue
                    || entry.Name.IndexOf('\0') >= 0 || entry.Name.IndexOf('/') >= 0 || entry.Name.IndexOf('\\') >= 0
                    || !names.Add(entry.Name)) throw new InvalidDataException("Encrypted summary has an invalid or duplicate entry name.");
                required = checked(required + entry.Bytes.LongLength + 18 + entry.Name.Length * 2L);
            }
            bool padLastEmpty = entries.Count != 0 && entries[entries.Count - 1].Bytes.Length == 0;
            if (padLastEmpty) required++;
            if (required > Math.Min(maximumBytes, int.MaxValue)) throw OfficeOutputLimit.Create("Encrypted summary exceeds the output materialization limit.");
            using var output = new OfficeBoundedMemoryStream(maximumBytes);
            using var data = new BinaryWriter(output, Encoding.Unicode, true);
            data.Write(new byte[8]);
            using var descriptions = new MemoryStream();
            using var descriptors = new BinaryWriter(descriptions, Encoding.Unicode, true);
            descriptors.Write(entries.Count);
            for (int index = 0; index < entries.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                var entry = entries[index];
                var encrypted = (byte[])entry.Bytes.Clone();
                session.TransformInPlace(encrypted, 0, encrypted.Length, checked((ushort)index), cancellationToken);
                descriptors.Write(checked((uint)output.Position));
                descriptors.Write(encrypted.Length);
                descriptors.Write(checked((ushort)index));
                descriptors.Write(checked((byte)entry.Name.Length));
                descriptors.Write(entry.IsStream ? (byte)1 : (byte)0);
                descriptors.Write(0U);
                descriptors.Write(Encoding.Unicode.GetBytes(entry.Name));
                descriptors.Write((ushort)0);
                data.Write(encrypted);
            }
            if (padLastEmpty) data.Write((byte)0);
            uint offset = checked((uint)output.Position);
            byte[] descriptorBytes = descriptions.ToArray();
            session.TransformInPlace(descriptorBytes, 0, descriptorBytes.Length, 0, cancellationToken);
            data.Write(descriptorBytes);
            using var header = new MemoryStream();
            using (var writer = new BinaryWriter(header, Encoding.Unicode, true)) {
                writer.Write(offset); writer.Write(descriptorBytes.Length);
            }
            byte[] location = header.ToArray();
            session.TransformInPlace(location, 0, location.Length, 0, cancellationToken);
            output.Position = 0; data.Write(location);
            cancellationToken.ThrowIfCancellationRequested();
            return output.ToArray();
        }
    }
}
