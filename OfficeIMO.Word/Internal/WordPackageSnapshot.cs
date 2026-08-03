#nullable enable
using System.Security.Cryptography;

namespace OfficeIMO.Word {
    /// <summary>Creates and fingerprints bounded Word package snapshots.</summary>
    internal static class WordPackageSnapshot {
        internal static void CopyBounded(string sourcePath, string snapshotPath, long maximumBytes) {
            using var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (source.Length > maximumBytes) {
                throw new InvalidDataException(
                    "The Word package length " + source.Length + " exceeds the configured limit of " + maximumBytes + " bytes.");
            }

            using var destination = new FileStream(snapshotPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            byte[] buffer = new byte[81920];
            long copied = 0;
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0) {
                copied = checked(copied + read);
                if (copied > maximumBytes) {
                    throw new InvalidDataException(
                        "The Word package exceeds the configured limit of " + maximumBytes + " bytes.");
                }
                destination.Write(buffer, 0, read);
            }
        }

        internal static string ComputeSha256(string filePath, long maximumBytes) {
            using FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (stream.Length > maximumBytes) {
                throw new InvalidDataException("The Word package exceeds its configured byte limit.");
            }

            using SHA256 hash = SHA256.Create();
            byte[] buffer = new byte[81920];
            long total = 0;
            while (true) {
                int read = stream.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                total = checked(total + read);
                if (total > maximumBytes) {
                    throw new InvalidDataException("The Word package exceeds its configured byte limit.");
                }
                hash.TransformBlock(buffer, 0, read, null, 0);
            }
            hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            return BitConverter.ToString(hash.Hash ?? Array.Empty<byte>()).Replace("-", string.Empty);
        }
    }
}
