using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.IWork;

public sealed partial class IWorkSourceDocument {
    /// <summary>Computes a deterministic SHA-256 hash of the captured normalized package entries.</summary>
    /// <remarks>This is a package-content identity, not a ZIP-file checksum. It includes normalized paths and bytes,
    /// including entries expanded from nested Index.zip. It uses the source's cancellation token and does not reopen files.</remarks>
    public string ComputePackageContentHash() {
        _cancellationToken.ThrowIfCancellationRequested();
        using IncrementalHash hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        hash.AppendData(Encoding.UTF8.GetBytes("OfficeIMO.IWork.PackageContent.v1\0"));
        AppendLength(Entries.Count);
        foreach (IWorkPackageEntry entry in Entries) {
            _cancellationToken.ThrowIfCancellationRequested();
            byte[] path = Encoding.UTF8.GetBytes(entry.Path);
            AppendLength(path.Length);
            hash.AppendData(path);
            AppendLength(entry.Length);
            for (int offset = 0; offset < entry.Length;) {
                _cancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(64 * 1024, entry.Length - offset);
                hash.AppendData(entry.Bytes, offset, count);
                offset += count;
            }
        }
        _cancellationToken.ThrowIfCancellationRequested();
        return BitConverter.ToString(hash.GetHashAndReset()).Replace("-", string.Empty).ToLowerInvariant();

        void AppendLength(int value) {
            byte[] bytes = { (byte)value, (byte)(value >> 8), (byte)(value >> 16), (byte)(value >> 24) };
            hash.AppendData(bytes);
        }
    }
}
