using System;
using System.IO;
using System.Security.Cryptography;

namespace OfficeIMO.Security;

internal static class OfficePackageFileSnapshot {
    internal static void CopyBounded(string sourcePath, string destinationPath, long maxBytes) {
        if (maxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
        var info = new FileInfo(sourcePath);
        if (!info.Exists) throw new FileNotFoundException("The package file does not exist.", sourcePath);
        if (info.Length > maxBytes) throw new InvalidDataException("The package exceeds the configured byte limit.");
        using var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var destination = new FileStream(destinationPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        CopyBounded(source, destination, maxBytes);
    }

    internal static string ComputeSha256(string path, long maxBytes) {
        if (maxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
        var info = new FileInfo(path);
        if (!info.Exists) throw new FileNotFoundException("The package file does not exist.", path);
        if (info.Length > maxBytes) throw new InvalidDataException("The package exceeds the configured byte limit.");
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using SHA256 algorithm = SHA256.Create();
        return BitConverter.ToString(algorithm.ComputeHash(input)).Replace("-", string.Empty);
    }

    private static void CopyBounded(Stream source, Stream destination, long maxBytes) {
        var buffer = new byte[81920];
        long total = 0;
        while (true) {
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            total = checked(total + read);
            if (total > maxBytes) throw new InvalidDataException("The package exceeds the configured byte limit.");
            destination.Write(buffer, 0, read);
        }
    }
}
