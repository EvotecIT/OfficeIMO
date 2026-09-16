using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Html.Runtime;

// The image ID pins the complete container. This digest separately proves that
// the published renderer/worker payloads in it match the operator's local builds.
internal static class HtmlPublicArtifactDigest {
    internal static string DirectorySha256(string directory, string? excludedTopLevelDirectory = null) {
        string root = Path.GetFullPath(directory);
        if (!Directory.Exists(root)) throw new DirectoryNotFoundException(root);
        var files = Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
            .Select(path => (Path: path, Relative: Path.GetRelativePath(root, path).Replace('\\', '/')))
            .Where(file => excludedTopLevelDirectory == null || !file.Relative.StartsWith(
                excludedTopLevelDirectory + "/", StringComparison.Ordinal))
            .OrderBy(file => file.Relative, StringComparer.Ordinal).ToArray();
        if (files.Length == 0) throw new IOException("The published artifact directory is empty.");
        using var digest = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        Span<byte> pathLength = stackalloc byte[4];
        foreach (var file in files) {
            byte[] pathBytes = Encoding.UTF8.GetBytes(file.Relative);
            BinaryPrimitives.WriteInt32LittleEndian(pathLength, pathBytes.Length);
            digest.AppendData(pathLength);
            digest.AppendData(pathBytes);
            using Stream input = File.OpenRead(file.Path);
            digest.AppendData(SHA256.HashData(input));
        }
        return Convert.ToHexString(digest.GetHashAndReset()).ToLowerInvariant();
    }
}
