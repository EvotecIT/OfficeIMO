using OfficeIMO.Drawing.Internal;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

internal static class OfficeCompatibilitySourceCarrier {
    internal const string MetadataPath = "OfficeIMOCompatibility/SourceMetadata";
    internal const string PayloadPath = "OfficeIMOCompatibility/SourcePayload";
    private const string Magic = "OfficeIMOCompatibilitySource";
    private const int SchemaVersion = 1;
    private const string PackageMetadataPath = "OfficeIMOCompatibility/SourceMetadata";
    private const string PackagePayloadPath = "OfficeIMOCompatibility/SourcePayload";
    private const string PackageRelationshipType = "https://schemas.officeimo.com/relationships/compatibility-source";
    private const string PackageMetadataContentType = "application/vnd.officeimo.compatibility-source-metadata";
    private const string PackagePayloadContentType = "application/vnd.officeimo.compatibility-source-payload";
    private const int MaxMetadataBytes = 64 * 1024;
    private const long MaxSourcePayloadBytes = 512L * 1024L * 1024L;
    private static readonly DateTimeOffset ReproducibleEntryTime =
        new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    internal static byte[] AttachToCompound(
        byte[] compoundBytes,
        string formatId,
        string fileName,
        OfficeCompatibilityMode mode,
        byte[] sourceBytes) {
        if (compoundBytes == null) throw new ArgumentNullException(nameof(compoundBytes));
        if (string.IsNullOrWhiteSpace(formatId)) throw new ArgumentException("Format id cannot be empty.", nameof(formatId));
        if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("File name cannot be empty.", nameof(fileName));
        if (sourceBytes == null) throw new ArgumentNullException(nameof(sourceBytes));
        ValidateSourcePayloadSize(sourceBytes.LongLength);
        if (!OfficeCompoundFileReader.TryRead(compoundBytes, out OfficeCompoundFile? compound, out string? error)
            || compound == null) {
            throw new InvalidDataException("The generated legacy Office output is not a readable compound file. " + error);
        }

        string sha256 = ComputeSha256(sourceBytes);
        byte[] metadata = CreateMetadata(formatId.Trim(), Path.GetFileName(fileName), sha256, mode);
        return OfficeCompoundFileWriter.Rewrite(
            compound,
            new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
                [MetadataPath] = metadata,
                [PayloadPath] = (byte[])sourceBytes.Clone()
            });
    }

    internal static bool TryRead(
        OfficeCompoundFile? compound,
        out OfficeCompatibilitySourcePayload? payload,
        out string? error) {
        payload = null;
        error = null;
        if (compound == null
            || !compound.Streams.TryGetValue(MetadataPath, out byte[]? metadata)
            || !compound.Streams.TryGetValue(PayloadPath, out byte[]? sourceBytes)) {
            return false;
        }

        try {
            payload = ReadPayload(metadata, sourceBytes);
            return true;
        } catch (Exception exception) when (
            exception is EndOfStreamException
            || exception is IOException
            || exception is InvalidDataException) {
            error = exception.Message;
            return false;
        }
    }

    internal static byte[] AttachToPackage(
        byte[] packageBytes,
        string formatId,
        string fileName,
        OfficeCompatibilityMode mode,
        byte[] sourceBytes) {
        if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
        if (string.IsNullOrWhiteSpace(formatId)) throw new ArgumentException("Format id cannot be empty.", nameof(formatId));
        if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("File name cannot be empty.", nameof(fileName));
        if (sourceBytes == null) throw new ArgumentNullException(nameof(sourceBytes));
        ValidateSourcePayloadSize(sourceBytes.LongLength);

        using var input = new MemoryStream(packageBytes, writable: false);
        using var source = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
        ZipArchiveEntry? contentTypesEntry = FindEntry(source, "[Content_Types].xml");
        ZipArchiveEntry? relationshipsEntry = FindEntry(source, "_rels/.rels");
        if (contentTypesEntry == null || relationshipsEntry == null) {
            throw new InvalidDataException("The generated modern Office output is not a valid OPC package.");
        }

        XDocument contentTypes = ReadXml(contentTypesEntry);
        XDocument relationships = ReadXml(relationshipsEntry);
        AddPackageContentTypes(contentTypes);
        AddPackageRelationship(relationships);

        using var output = new MemoryStream(packageBytes.Length + sourceBytes.Length + 4096);
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in source.Entries) {
                if (IsPackageCarrierEntry(entry.FullName)
                    || string.Equals(entry.FullName, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(entry.FullName, "_rels/.rels", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }
                CopyEntry(entry, target);
            }

            WriteXmlEntry(target, "[Content_Types].xml", contentTypes);
            WriteXmlEntry(target, "_rels/.rels", relationships);
            WriteEntry(target, PackageMetadataPath, CreateMetadata(
                formatId.Trim(),
                Path.GetFileName(fileName),
                ComputeSha256(sourceBytes),
                mode));
            WriteEntry(target, PackagePayloadPath, sourceBytes);
        }
        return output.ToArray();
    }

    internal static bool TryReadPackage(
        byte[]? packageBytes,
        out OfficeCompatibilitySourcePayload? payload,
        out string? error) {
        payload = null;
        error = null;
        if (packageBytes == null || packageBytes.Length == 0) return false;

        try {
            using var input = new MemoryStream(packageBytes, writable: false);
            using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
            ZipArchiveEntry? metadataEntry = FindEntry(archive, PackageMetadataPath);
            ZipArchiveEntry? payloadEntry = FindEntry(archive, PackagePayloadPath);
            if (metadataEntry == null || payloadEntry == null) return false;
            payload = ReadPayload(
                ReadEntry(metadataEntry, MaxMetadataBytes, "compatibility source metadata"),
                ReadEntry(payloadEntry, MaxSourcePayloadBytes, "compatibility source payload"));
            return true;
        } catch (Exception exception) when (
            exception is EndOfStreamException
            || exception is IOException
            || exception is InvalidDataException
            || exception is NotSupportedException) {
            error = exception.Message;
            return false;
        }
    }

    internal static bool ContainsPackageCarrier(byte[]? packageBytes) {
        if (packageBytes == null || packageBytes.Length == 0) return false;
        try {
            using var input = new MemoryStream(packageBytes, writable: false);
            using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
            return FindEntry(archive, PackageMetadataPath) != null
                && FindEntry(archive, PackagePayloadPath) != null;
        } catch (Exception exception) when (
            exception is EndOfStreamException
            || exception is IOException
            || exception is InvalidDataException
            || exception is NotSupportedException) {
            return false;
        }
    }

    private static OfficeCompatibilitySourcePayload ReadPayload(byte[] metadata, byte[] sourceBytes) {
        if (metadata.LongLength > MaxMetadataBytes) {
            throw new InvalidDataException($"Compatibility source metadata exceeds {MaxMetadataBytes} bytes.");
        }
        ValidateSourcePayloadSize(sourceBytes.LongLength);
        using var stream = new MemoryStream(metadata, writable: false);
        using var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: false);
        if (!string.Equals(reader.ReadString(), Magic, StringComparison.Ordinal)) {
            throw new InvalidDataException("Compatibility source metadata has an invalid signature.");
        }
        int version = reader.ReadInt32();
        if (version != SchemaVersion) {
            throw new InvalidDataException($"Unsupported compatibility source metadata schema '{version}'.");
        }
        string formatId = reader.ReadString();
        string fileName = reader.ReadString();
        string expectedSha256 = reader.ReadString();
        int rawMode = reader.ReadInt32();
        if (!Enum.IsDefined(typeof(OfficeCompatibilityMode), rawMode)) {
            throw new InvalidDataException($"Unknown compatibility mode '{rawMode}'.");
        }
        if (stream.Position != stream.Length) {
            throw new InvalidDataException("Compatibility source metadata contains unexpected trailing bytes.");
        }

        string actualSha256 = ComputeSha256(sourceBytes);
        if (!string.Equals(expectedSha256, actualSha256, StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidDataException(
                $"Compatibility source payload SHA-256 mismatch. Expected {expectedSha256}, got {actualSha256}.");
        }

        return new OfficeCompatibilitySourcePayload(
            formatId,
            fileName,
            actualSha256,
            (OfficeCompatibilityMode)rawMode,
            (byte[])sourceBytes.Clone());
    }

    private static void AddPackageContentTypes(XDocument document) {
        XElement root = document.Root ?? throw new InvalidDataException("The OPC content-types part has no root element.");
        XNamespace ns = root.Name.Namespace;
        root.Elements(ns + "Override")
            .Where(element => IsCarrierPartName((string?)element.Attribute("PartName")))
            .Remove();
        root.Add(
            new XElement(ns + "Override",
                new XAttribute("PartName", "/" + PackageMetadataPath),
                new XAttribute("ContentType", PackageMetadataContentType)),
            new XElement(ns + "Override",
                new XAttribute("PartName", "/" + PackagePayloadPath),
                new XAttribute("ContentType", PackagePayloadContentType)));
    }

    private static void AddPackageRelationship(XDocument document) {
        XElement root = document.Root ?? throw new InvalidDataException("The OPC relationships part has no root element.");
        XNamespace ns = root.Name.Namespace;
        root.Elements(ns + "Relationship")
            .Where(element => string.Equals((string?)element.Attribute("Type"), PackageRelationshipType, StringComparison.Ordinal)
                || IsCarrierPartName((string?)element.Attribute("Target")))
            .Remove();
        var ids = new HashSet<string>(root.Elements(ns + "Relationship")
            .Select(element => (string?)element.Attribute("Id"))
            .Where(id => !string.IsNullOrWhiteSpace(id))!, StringComparer.Ordinal);
        int index = 1;
        while (ids.Contains("rIdOfficeIMOCompatibility" + index)) index++;
        root.Add(new XElement(ns + "Relationship",
            new XAttribute("Id", "rIdOfficeIMOCompatibility" + index),
            new XAttribute("Type", PackageRelationshipType),
            new XAttribute("Target", PackageMetadataPath)));
    }

    private static bool IsPackageCarrierEntry(string path) =>
        string.Equals(path, PackageMetadataPath, StringComparison.OrdinalIgnoreCase)
        || string.Equals(path, PackagePayloadPath, StringComparison.OrdinalIgnoreCase);

    private static bool IsCarrierPartName(string? path) {
        if (string.IsNullOrWhiteSpace(path)) return false;
        return IsPackageCarrierEntry(path!.TrimStart('/'));
    }

    private static ZipArchiveEntry? FindEntry(ZipArchive archive, string path) => archive.Entries
        .FirstOrDefault(entry => string.Equals(entry.FullName, path, StringComparison.OrdinalIgnoreCase));

    private static XDocument ReadXml(ZipArchiveEntry entry) {
        using Stream stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static byte[] ReadEntry(ZipArchiveEntry entry, long maxBytes, string description) {
        if (entry.Length > maxBytes) {
            throw new InvalidDataException($"The {description} exceeds the {maxBytes}-byte limit.");
        }
        using Stream input = entry.Open();
        using var output = new MemoryStream(entry.Length <= int.MaxValue ? (int)entry.Length : 0);
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = input.Read(buffer, 0, buffer.Length)) > 0) {
            total = checked(total + read);
            if (total > maxBytes) {
                throw new InvalidDataException($"The {description} expands beyond the {maxBytes}-byte limit.");
            }
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    private static void ValidateSourcePayloadSize(long length) {
        if (length > MaxSourcePayloadBytes) {
            throw new InvalidDataException(
                $"Compatibility source payload exceeds the {MaxSourcePayloadBytes}-byte limit.");
        }
    }

    private static void CopyEntry(ZipArchiveEntry source, ZipArchive target) {
        ZipArchiveEntry destination = target.CreateEntry(source.FullName, CompressionLevel.Optimal);
        destination.LastWriteTime = source.LastWriteTime.Year >= 1980 ? source.LastWriteTime : ReproducibleEntryTime;
        using Stream input = source.Open();
        using Stream output = destination.Open();
        input.CopyTo(output);
    }

    private static void WriteXmlEntry(ZipArchive archive, string path, XDocument document) {
        using var buffer = new MemoryStream();
        document.Save(buffer, SaveOptions.DisableFormatting);
        WriteEntry(archive, path, buffer.ToArray());
    }

    private static void WriteEntry(ZipArchive archive, string path, byte[] bytes) {
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        entry.LastWriteTime = ReproducibleEntryTime;
        using Stream output = entry.Open();
        output.Write(bytes, 0, bytes.Length);
    }

    private static byte[] CreateMetadata(
        string formatId,
        string fileName,
        string sha256,
        OfficeCompatibilityMode mode) {
        using var output = new MemoryStream();
        using (var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(Magic);
            writer.Write(SchemaVersion);
            writer.Write(formatId);
            writer.Write(fileName);
            writer.Write(sha256);
            writer.Write((int)mode);
        }
        return output.ToArray();
    }

    private static string ComputeSha256(byte[] bytes) {
        using SHA256 sha256 = SHA256.Create();
        byte[] hash = sha256.ComputeHash(bytes);
        var text = new StringBuilder(hash.Length * 2);
        foreach (byte value in hash) text.Append(value.ToString("x2"));
        return text.ToString();
    }
}
