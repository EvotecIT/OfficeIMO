using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Produces compact, length-delimited SHA-256 guards for projected XML and
    /// related payload descriptors without retaining their full concatenation.
    /// </summary>
    internal static class LegacyPptProjectionDigest {
        internal static IncrementalHash CreateBuilder() =>
            IncrementalHash.CreateHash(HashAlgorithmName.SHA256);

        internal static void Append(IncrementalHash hash, string? value) {
            if (hash == null) throw new ArgumentNullException(nameof(hash));
            byte[] bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
            var length = new byte[4];
            uint count = checked((uint)bytes.Length);
            length[0] = unchecked((byte)count);
            length[1] = unchecked((byte)(count >> 8));
            length[2] = unchecked((byte)(count >> 16));
            length[3] = unchecked((byte)(count >> 24));
            hash.AppendData(length);
            hash.AppendData(bytes);
        }

        internal static string Finish(IncrementalHash hash) {
            if (hash == null) throw new ArgumentNullException(nameof(hash));
            return Convert.ToBase64String(hash.GetHashAndReset());
        }

        internal static string Create(params string?[] values) {
            using IncrementalHash hash = CreateBuilder();
            foreach (string? value in values) Append(hash, value);
            return Finish(hash);
        }
    }
}
