using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace OfficeIMO.ContentSafety;

/// <summary>Applies encoded and package-expansion limits before format parsers allocate decoded content.</summary>
public static class OfficeContentSafetyInputGuard {
    /// <summary>Reads a file only after validating its encoded length and, when applicable, ZIP package metadata.</summary>
    public static byte[] ReadAllBytes(string filePath, OfficeContentSafetyOptions options, bool inspectZipPackage = false) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate();
        string fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The input file was not found.", fullPath);
        long length = new FileInfo(fullPath).Length;
        if (length > options.MaxInputBytes) throw new InvalidDataException("The encoded asset exceeds the configured input-byte limit.");
        byte[] bytes;
        using (FileStream stream = File.OpenRead(fullPath)) {
            bytes = ReadBounded(stream, options.MaxInputBytes);
        }
        ValidateBytes(bytes, options, inspectZipPackage);
        return bytes;
    }

    /// <summary>Validates already-buffered input before a parser or package loader is invoked.</summary>
    public static void ValidateBytes(byte[] input, OfficeContentSafetyOptions options, bool inspectZipPackage = false) {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate();
        if (input.LongLength > options.MaxInputBytes) throw new InvalidDataException("The encoded asset exceeds the configured input-byte limit.");
        if (inspectZipPackage && LooksLikeZip(input)) ValidateZipPackage(input, options);
    }

    /// <summary>Validates an in-memory text input before HTML or other text parsers are invoked.</summary>
    public static void ValidateText(string text, OfficeContentSafetyOptions options) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate();
        if (text.Length > options.MaxCharacters) throw new InvalidDataException("The asset exceeds the configured decoded-character limit.");
        if (Encoding.UTF8.GetByteCount(text) > options.MaxInputBytes) throw new InvalidDataException("The encoded asset exceeds the configured input-byte limit.");
    }

    /// <summary>Reads a BOM-aware UTF-8 text file through the encoded-input guard.</summary>
    public static string ReadUtf8Text(string filePath, OfficeContentSafetyOptions options) {
        byte[] bytes = ReadAllBytes(filePath, options);
        return DecodeText(bytes, options);
    }

    /// <summary>Decodes guarded BOM-aware text bytes, defaulting to strict UTF-8 when no BOM is present.</summary>
    internal static string DecodeText(byte[] bytes, OfficeContentSafetyOptions options) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (options == null) throw new ArgumentNullException(nameof(options));
        ValidateBytes(bytes, options);
        try {
            using var stream = new MemoryStream(bytes, writable: false);
            using var reader = new StreamReader(stream, new UTF8Encoding(false, true), detectEncodingFromByteOrderMarks: true);
            string text = reader.ReadToEnd();
            ValidateText(text, options);
            return text;
        } catch (DecoderFallbackException exception) {
            throw new InvalidDataException("The text input contains invalid encoded text.", exception);
        }
    }

    private static void ValidateZipPackage(byte[] input, OfficeContentSafetyOptions options) {
        try {
            using var stream = new MemoryStream(input, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            if (archive.Entries.Count > options.MaxPackageEntries) throw new InvalidDataException("The package exceeds the configured entry-count limit.");
            long expanded = 0;
            foreach (ZipArchiveEntry entry in archive.Entries) {
                if (entry.Length < 0 || expanded > options.MaxExpandedPackageBytes - entry.Length) {
                    throw new InvalidDataException("The package exceeds the configured expanded-byte limit.");
                }
                expanded += entry.Length;
            }
        } catch (InvalidDataException) {
            throw;
        } catch (Exception exception) when (exception is IOException || exception is NotSupportedException) {
            throw new InvalidDataException("The ZIP package could not be safely preflighted.", exception);
        }
    }

    private static byte[] ReadBounded(Stream stream, long maximumBytes) {
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
            if (total > maximumBytes - read) throw new InvalidDataException("The encoded asset exceeds the configured input-byte limit.");
            output.Write(buffer, 0, read);
            total += read;
        }
        return output.ToArray();
    }

    private static bool LooksLikeZip(byte[] input) => input.Length >= 4 && input[0] == 0x50 && input[1] == 0x4B &&
        ((input[2] == 0x03 && input[3] == 0x04) || (input[2] == 0x05 && input[3] == 0x06) || (input[2] == 0x07 && input[3] == 0x08));
}
