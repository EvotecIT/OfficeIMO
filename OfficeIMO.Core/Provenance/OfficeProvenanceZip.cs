using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceZip {
    private const string ManifestPath = "META-INF/content_credential.c2pa";

    internal static bool HasSignature(byte[] data) => data.Length >= 4 && data[0] == 0x50 && data[1] == 0x4B &&
        ((data[2] == 0x03 && data[3] == 0x04) || (data[2] == 0x05 && data[3] == 0x06) || (data[2] == 0x07 && data[3] == 0x08));

    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        ValidateEntryCount(data, options.MaxContainerEntries);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        int index = 0;
        int embeddedCount = 0;
        long expandedBytes = 0;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            if (entry.FullName.Equals(ManifestPath, StringComparison.Ordinal)) {
                if (entry.Length > options.MaxManifestBytes || entry.Length > int.MaxValue) {
                    throw new InvalidDataException("ZIP provenance manifest exceeds the configured manifest limit.");
                }
                ReserveExpandedBytes(ref expandedBytes, entry.Length, options.MaxExpandedContainerBytes);
                byte[] manifest = ReadEntry(entry, (int)entry.Length);
                bool valid = OfficeC2paManifestStore.IsValid(
                    manifest, 0, manifest.Length, options.MaxManifestBytes, options.MaxContainerEntries, out _);
                context.Add(new OfficeProvenanceEvidence(
                    OfficeProvenanceCarrierKind.C2paManifest,
                    $"ZIP/{ManifestPath}[{index++}]",
                    valid,
                    manifest.Length));
                continue;
            }
            if (!options.ProcessEmbeddedAssets ||
                !IsSupportedEmbeddedAsset(entry, options, ref expandedBytes)) continue;
            embeddedCount++;
            if (embeddedCount > options.MaxEmbeddedAssets) throw new InvalidDataException("ZIP package exceeds the configured embedded-asset limit.");
            if (entry.Length > options.MaxAssetBytes || entry.Length > int.MaxValue) {
                throw new InvalidDataException("A supported embedded asset exceeds the configured asset limit.");
            }
            ReserveExpandedBytes(ref expandedBytes, entry.Length, options.MaxExpandedContainerBytes);
            byte[] asset = ReadEntry(entry, (int)entry.Length);
            if (!IsSupportedEmbeddedImage(asset, options)) continue;
            OfficeProvenanceReport nested;
            try {
                nested = OfficeProvenanceInspector.InspectCore(asset, entry.FullName, CreateNestedOptions(options));
            } catch (Exception exception) when (exception is InvalidDataException || exception is XmlException) {
                context.Diagnostics.Add($"ZIP/{entry.FullName}: embedded asset was preserved because inspection failed: {exception.Message}");
                continue;
            }
            foreach (OfficeProvenanceEvidence evidence in nested.Evidence) context.Add(PrefixEvidence(entry.FullName, evidence));
            foreach (string diagnostic in nested.Diagnostics) context.Diagnostics.Add($"ZIP/{entry.FullName}: {diagnostic}");
        }
        if (index > 1) context.Diagnostics.Add("The ZIP package contains multiple C2PA manifest entries.");
    }

    internal static byte[] Remove(
        byte[] data,
        OfficeProvenanceRemovalOptions options,
        List<OfficeProvenanceChange> changes,
        out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata && !options.RemoveExternalC2paReferences) return (byte[])data.Clone();
        ValidateEntryCount(data, options.Limits.MaxContainerEntries);
        using var inputStream = new MemoryStream(data, writable: false);
        using var input = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: false);
        var removable = new HashSet<string>(StringComparer.Ordinal);
        var embeddedRewrites = new Dictionary<ZipArchiveEntry, byte[]>();
        int occurrence = 0;
        int embeddedCount = 0;
        long inspectionBytes = 0;
        foreach (ZipArchiveEntry entry in input.Entries) {
            if (entry.FullName.Equals(ManifestPath, StringComparison.Ordinal)) {
                if (entry.Length > options.Limits.MaxManifestBytes || entry.Length > int.MaxValue) throw new InvalidDataException("ZIP provenance manifest exceeds the configured limit.");
                ReserveExpandedBytes(ref inspectionBytes, entry.Length, options.Limits.MaxExpandedContainerBytes);
                byte[] manifest = ReadEntry(entry, (int)entry.Length);
                bool valid = OfficeC2paManifestStore.IsValid(
                    manifest, 0, manifest.Length, options.Limits.MaxManifestBytes, options.Limits.MaxContainerEntries, out _);
                if (options.RemoveC2paManifests && (valid || !options.RequireStructurallyValidCarrier)) {
                    removable.Add(entry.FullName + "\0" + occurrence);
                    changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, $"ZIP/{ManifestPath}[{occurrence}]", 0));
                }
                occurrence++;
                continue;
            }
            if (!options.ProcessEmbeddedAssets ||
                !IsSupportedEmbeddedAsset(entry, options.Limits, ref inspectionBytes)) continue;
            embeddedCount++;
            if (embeddedCount > options.MaxEmbeddedAssets) throw new InvalidDataException("ZIP package exceeds the configured embedded-asset limit.");
            if (entry.Length > options.Limits.MaxAssetBytes || entry.Length > int.MaxValue) throw new InvalidDataException("A supported embedded asset exceeds the configured asset limit.");
            ReserveExpandedBytes(ref inspectionBytes, entry.Length, options.Limits.MaxExpandedContainerBytes);
            byte[] asset = ReadEntry(entry, (int)entry.Length);
            if (!IsSupportedEmbeddedImage(asset, options.Limits)) continue;
            OfficeProvenanceRemovalResult nested;
            try {
                nested = OfficeProvenanceRemover.Remove(asset, entry.FullName, CreateNestedRemovalOptions(options));
            } catch (Exception exception) when (exception is InvalidDataException || exception is XmlException) {
                // Malformed embedded assets are preserved; document-level diagnostics are available during inspection.
                continue;
            }
            if (!nested.WasChanged) continue;
            if (nested.Changes.Count > options.Limits.MaxCarriers - changes.Count) {
                throw new InvalidDataException($"The asset exceeds the configured carrier limit of {options.Limits.MaxCarriers}.");
            }
            embeddedRewrites.Add(entry, nested.ToArray());
            foreach (OfficeProvenanceChange change in nested.Changes) {
                changes.Add(new OfficeProvenanceChange(change.Carrier, $"ZIP/{entry.FullName}/{change.Location}", 0));
            }
            if (nested.WasReserialized) reserialized = true;
        }
        if (changes.Count == 0) return (byte[])data.Clone();

        bool hasSignature = options.SignatureMutationPolicy != OfficeIMO.OfficeSignatureMutationPolicy.PreserveSignatureMarkup &&
            HasPackageSignature(data, input, options);
        if (hasSignature && options.SignatureMutationPolicy == OfficeIMO.OfficeSignatureMutationPolicy.BlockSave) {
            throw new InvalidOperationException("Removing provenance would invalidate package signatures. Choose an explicit signature mutation policy.");
        }
        if (hasSignature && options.SignatureMutationPolicy == OfficeIMO.OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures) {
            throw new NotSupportedException("Generic ZIP provenance removal cannot safely rewrite package signatures. Use the owning OfficeIMO document format API.");
        }

        var outputEntries = new List<OfficeProvenanceZipWriteEntry>();
        Dictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata = GetEntryMetadata(data, input);
        occurrence = 0;
        IEnumerable<ZipArchiveEntry> orderedEntries = input.Entries
            .OrderByDescending(entry => entry.FullName.Equals("mimetype", StringComparison.Ordinal));
        foreach (ZipArchiveEntry entry in orderedEntries) {
            bool isManifest = entry.FullName.Equals(ManifestPath, StringComparison.Ordinal);
            string key = entry.FullName + "\0" + occurrence;
            if (isManifest) occurrence++;
            if (isManifest && removable.Contains(key)) continue;
            embeddedRewrites.TryGetValue(entry, out byte[]? rewritten);
            outputEntries.Add(CreateWriteEntry(entry, entryMetadata[entry], rewritten));
        }
        reserialized = true;
        return OfficeProvenanceZipWriter.Write(outputEntries, options.Limits.MaxExpandedContainerBytes, ReadArchiveComment(data));
    }

    private static uint GetExternalAttributes(ZipArchiveEntry source) {
        System.Reflection.PropertyInfo? property = typeof(ZipArchiveEntry).GetProperty("ExternalAttributes");
        object? value = property?.CanRead == true ? property.GetValue(source, null) : null;
        return value is int signed ? unchecked((uint)signed) : value is uint unsigned ? unsigned : 0u;
    }

    private static OfficeProvenanceZipWriteEntry CreateWriteEntry(
        ZipArchiveEntry entry,
        OfficeProvenanceZipEntryMetadata metadata,
        byte[]? replacement = null) {
        bool isMimetype = entry.FullName.Equals("mimetype", StringComparison.Ordinal);
        if (replacement != null) {
            return new OfficeProvenanceZipWriteEntry(
                entry.FullName,
                replacement.LongLength,
                compress: !isMimetype,
                entry.LastWriteTime,
                metadata.InternalAttributes,
                GetExternalAttributes(entry),
                metadata.LocalExtraField,
                metadata.CentralExtraField,
                metadata.Comment,
                () => new MemoryStream(replacement, writable: false));
        }
        return new OfficeProvenanceZipWriteEntry(
            entry.FullName,
            entry.Length,
            compress: !isMimetype,
            entry.LastWriteTime,
            metadata.InternalAttributes,
            GetExternalAttributes(entry),
            metadata.LocalExtraField,
            metadata.CentralExtraField,
            metadata.Comment,
            entry.Open);
    }

    /// <summary>
    /// Removes matching package entries while preserving the ODF/EPUB stored <c>mimetype</c> contract.
    /// </summary>
    internal static OfficeProvenanceSignatureStripResult RemoveEntries(
        byte[] data,
        Func<string, bool> shouldRemove) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (shouldRemove == null) throw new ArgumentNullException(nameof(shouldRemove));
        using var inputStream = new MemoryStream(data, writable: false);
        using var input = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: false);
        bool hadMatches = input.Entries.Any(entry => shouldRemove(entry.FullName));
        if (!hadMatches) return new OfficeProvenanceSignatureStripResult((byte[])data.Clone(), hadSignatures: false);

        Dictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata = GetEntryMetadata(data, input);
        List<OfficeProvenanceZipWriteEntry> outputEntries = input.Entries
            .OrderByDescending(entry => entry.FullName.Equals("mimetype", StringComparison.Ordinal))
            .Where(entry => !shouldRemove(entry.FullName))
            .Select(entry => CreateWriteEntry(entry, entryMetadata[entry]))
            .ToList();
        return new OfficeProvenanceSignatureStripResult(
            OfficeProvenanceZipWriter.Write(outputEntries, long.MaxValue, ReadArchiveComment(data)),
            hadSignatures: true);
    }

    internal static bool HasEntry(byte[] data, Func<string, bool> predicate) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (predicate == null) throw new ArgumentNullException(nameof(predicate));
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        return archive.Entries.Any(entry => predicate(entry.FullName));
    }

    internal static void ValidateEntryCount(byte[] data, int maximumEntries) {
        const uint endOfCentralDirectorySignature = 0x06054B50;
        const uint zip64LocatorSignature = 0x07064B50;
        const uint zip64EndOfCentralDirectorySignature = 0x06064B50;
        int minimumOffset = Math.Max(0, data.Length - (22 + ushort.MaxValue));
        int endOffset = -1;
        for (int offset = data.Length - 22; offset >= minimumOffset; offset--) {
            if (OfficeProvenanceBinary.ReadUInt32(data, offset, littleEndian: true) != endOfCentralDirectorySignature) continue;
            ushort commentLength = OfficeProvenanceBinary.ReadUInt16(data, offset + 20, littleEndian: true);
            if (offset + 22 + commentLength != data.Length) continue;
            endOffset = offset;
            break;
        }
        if (endOffset < 0) throw new InvalidDataException("ZIP package does not contain a valid end-of-central-directory record.");

        ushort diskNumber = OfficeProvenanceBinary.ReadUInt16(data, endOffset + 4, littleEndian: true);
        ushort directoryDisk = OfficeProvenanceBinary.ReadUInt16(data, endOffset + 6, littleEndian: true);
        ushort entriesOnDisk = OfficeProvenanceBinary.ReadUInt16(data, endOffset + 8, littleEndian: true);
        ushort totalEntries = OfficeProvenanceBinary.ReadUInt16(data, endOffset + 10, littleEndian: true);
        if (diskNumber != 0 || directoryDisk != 0 || entriesOnDisk != totalEntries) {
            throw new InvalidDataException("Multi-disk ZIP packages are not supported.");
        }

        ulong entryCount = totalEntries;
        uint centralDirectorySize = OfficeProvenanceBinary.ReadUInt32(data, endOffset + 12, littleEndian: true);
        uint centralDirectoryOffset = OfficeProvenanceBinary.ReadUInt32(data, endOffset + 16, littleEndian: true);
        if ((totalEntries == ushort.MaxValue || centralDirectorySize == uint.MaxValue || centralDirectoryOffset == uint.MaxValue) &&
            HasZip64Locator(data, endOffset, zip64LocatorSignature)) {
            int locatorOffset = endOffset - 20;
            if (OfficeProvenanceBinary.ReadUInt32(data, locatorOffset + 4, littleEndian: true) != 0 ||
                OfficeProvenanceBinary.ReadUInt32(data, locatorOffset + 16, littleEndian: true) != 1) {
                throw new InvalidDataException("ZIP64 package does not contain a valid locator.");
            }
            ulong recordOffset = OfficeProvenanceBinary.ReadUInt64(data, locatorOffset + 8, littleEndian: true);
            if (data.Length < 56 || recordOffset > (ulong)(data.Length - 56)) {
                throw new InvalidDataException("ZIP64 end-of-central-directory record is outside the package.");
            }
            int zip64Offset = (int)recordOffset;
            if (OfficeProvenanceBinary.ReadUInt32(data, zip64Offset, littleEndian: true) != zip64EndOfCentralDirectorySignature ||
                OfficeProvenanceBinary.ReadUInt32(data, zip64Offset + 16, littleEndian: true) != 0 ||
                OfficeProvenanceBinary.ReadUInt32(data, zip64Offset + 20, littleEndian: true) != 0) {
                throw new InvalidDataException("ZIP64 end-of-central-directory record is invalid.");
            }
            ulong entriesOnZip64Disk = OfficeProvenanceBinary.ReadUInt64(data, zip64Offset + 24, littleEndian: true);
            entryCount = OfficeProvenanceBinary.ReadUInt64(data, zip64Offset + 32, littleEndian: true);
            if (entriesOnZip64Disk != entryCount) throw new InvalidDataException("Multi-disk ZIP64 packages are not supported.");
        }
        if (entryCount > (ulong)maximumEntries) throw new InvalidDataException("ZIP package exceeds the configured entry limit.");
    }

    private static Dictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> GetEntryMetadata(byte[] data, ZipArchive archive) {
        List<OfficeProvenanceZipEntryMetadata> metadata = ReadCentralDirectoryMetadata(data, archive.Entries.Count);
        var result = new Dictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata>(archive.Entries.Count);
        for (int index = 0; index < archive.Entries.Count; index++) result.Add(archive.Entries[index], metadata[index]);
        return result;
    }

    private static byte[] ReadArchiveComment(byte[] data) {
        int endOffset = FindEndOfCentralDirectory(data);
        int commentLength = OfficeProvenanceBinary.ReadUInt16(data, endOffset + 20, littleEndian: true);
        byte[] comment = new byte[commentLength];
        if (commentLength != 0) Buffer.BlockCopy(data, endOffset + 22, comment, 0, commentLength);
        return comment;
    }

    private static int FindEndOfCentralDirectory(byte[] data) {
        const uint endSignature = 0x06054B50;
        int minimumOffset = Math.Max(0, data.Length - (22 + ushort.MaxValue));
        for (int offset = data.Length - 22; offset >= minimumOffset; offset--) {
            if (OfficeProvenanceBinary.ReadUInt32(data, offset, littleEndian: true) != endSignature) continue;
            ushort commentLength = OfficeProvenanceBinary.ReadUInt16(data, offset + 20, littleEndian: true);
            if (offset + 22 + commentLength == data.Length) return offset;
        }
        throw new InvalidDataException("ZIP package does not contain a valid end-of-central-directory record.");
    }

    private static List<OfficeProvenanceZipEntryMetadata> ReadCentralDirectoryMetadata(byte[] data, int expectedEntries) {
        const uint zip64LocatorSignature = 0x07064B50;
        const uint zip64EndSignature = 0x06064B50;
        const uint centralHeaderSignature = 0x02014B50;
        int endOffset = FindEndOfCentralDirectory(data);
        uint centralSize = OfficeProvenanceBinary.ReadUInt32(data, endOffset + 12, littleEndian: true);
        ulong centralOffset = OfficeProvenanceBinary.ReadUInt32(data, endOffset + 16, littleEndian: true);
        ushort totalEntries = OfficeProvenanceBinary.ReadUInt16(data, endOffset + 10, littleEndian: true);
        if ((totalEntries == ushort.MaxValue || centralSize == uint.MaxValue || centralOffset == uint.MaxValue) &&
            HasZip64Locator(data, endOffset, zip64LocatorSignature)) {
            int locatorOffset = endOffset - 20;
            ulong recordOffset = OfficeProvenanceBinary.ReadUInt64(data, locatorOffset + 8, littleEndian: true);
            if (recordOffset > (ulong)(data.Length - 56) ||
                OfficeProvenanceBinary.ReadUInt32(data, (int)recordOffset, littleEndian: true) != zip64EndSignature) {
                throw new InvalidDataException("ZIP64 end-of-central-directory record is invalid.");
            }
            centralOffset = OfficeProvenanceBinary.ReadUInt64(data, (int)recordOffset + 48, littleEndian: true);
        }
        if (centralOffset > int.MaxValue || centralOffset > (ulong)data.Length) {
            throw new InvalidDataException("ZIP central directory is outside the package.");
        }
        int cursor = (int)centralOffset;
        var metadata = new List<OfficeProvenanceZipEntryMetadata>(expectedEntries);
        for (int index = 0; index < expectedEntries; index++) {
            if (cursor > data.Length - 46 || OfficeProvenanceBinary.ReadUInt32(data, cursor, littleEndian: true) != centralHeaderSignature) {
                throw new InvalidDataException("ZIP central directory is malformed.");
            }
            int nameLength = OfficeProvenanceBinary.ReadUInt16(data, cursor + 28, littleEndian: true);
            int extraLength = OfficeProvenanceBinary.ReadUInt16(data, cursor + 30, littleEndian: true);
            int commentLength = OfficeProvenanceBinary.ReadUInt16(data, cursor + 32, littleEndian: true);
            ushort internalAttributes = OfficeProvenanceBinary.ReadUInt16(data, cursor + 36, littleEndian: true);
            ushort flags = OfficeProvenanceBinary.ReadUInt16(data, cursor + 8, littleEndian: true);
            uint localHeaderOffset = OfficeProvenanceBinary.ReadUInt32(data, cursor + 42, littleEndian: true);
            long recordEnd = (long)cursor + 46 + nameLength + extraLength + commentLength;
            if (recordEnd > data.Length) throw new InvalidDataException("ZIP central-directory entry exceeds the package bounds.");
            byte[] centralExtraField = new byte[extraLength];
            if (extraLength != 0) Buffer.BlockCopy(data, cursor + 46 + nameLength, centralExtraField, 0, extraLength);
            if (localHeaderOffset == uint.MaxValue) {
                localHeaderOffset = ResolveZip64LocalHeaderOffset(data, cursor, centralExtraField);
            }
            if (localHeaderOffset > int.MaxValue) {
                throw new InvalidDataException("ZIP local-header offset exceeds the supported package bounds.");
            }
            byte[] comment = new byte[commentLength];
            if (commentLength != 0) Buffer.BlockCopy(data, cursor + 46 + nameLength + extraLength, comment, 0, commentLength);
            if ((flags & 0x0800) == 0) comment = TranscodeLegacyZipComment(comment);
            int localOffset = (int)localHeaderOffset;
            if (localOffset > data.Length - 30 || OfficeProvenanceBinary.ReadUInt32(data, localOffset, littleEndian: true) != 0x04034B50U) {
                throw new InvalidDataException("ZIP local-file header is outside the package bounds.");
            }
            int localNameLength = OfficeProvenanceBinary.ReadUInt16(data, localOffset + 26, littleEndian: true);
            int localExtraLength = OfficeProvenanceBinary.ReadUInt16(data, localOffset + 28, littleEndian: true);
            long localHeaderEnd = (long)localOffset + 30 + localNameLength + localExtraLength;
            if (localHeaderEnd > data.Length) throw new InvalidDataException("ZIP local-file header exceeds the package bounds.");
            byte[] localExtraField = new byte[localExtraLength];
            if (localExtraLength != 0) Buffer.BlockCopy(data, localOffset + 30 + localNameLength, localExtraField, 0, localExtraLength);
            metadata.Add(new OfficeProvenanceZipEntryMetadata(localExtraField, centralExtraField, comment, internalAttributes));
            cursor = (int)recordEnd;
        }
        return metadata;
    }

    private static byte[] TranscodeLegacyZipComment(byte[] comment) {
        if (comment.Length == 0) return comment;
        const string highCharacters =
            "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»" +
            "░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌" +
            "█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ";
        var decoded = new char[comment.Length];
        for (int index = 0; index < comment.Length; index++) {
            byte value = comment[index];
            decoded[index] = value < 0x80 ? (char)value : highCharacters[value - 0x80];
        }
        int encodedLength = Encoding.UTF8.GetByteCount(decoded);
        if (encodedLength <= ushort.MaxValue) return Encoding.UTF8.GetBytes(decoded);
        byte[] bounded = new byte[ushort.MaxValue];
        Encoding.UTF8.GetEncoder().Convert(
            decoded,
            0,
            decoded.Length,
            bounded,
            0,
            bounded.Length,
            flush: true,
            out _,
            out int bytesUsed,
            out _);
        if (bytesUsed != bounded.Length) Array.Resize(ref bounded, bytesUsed);
        return bounded;
    }

    private static uint ResolveZip64LocalHeaderOffset(byte[] data, int centralHeaderOffset, byte[] extraField) {
        bool hasUncompressedSize = OfficeProvenanceBinary.ReadUInt32(data, centralHeaderOffset + 24, littleEndian: true) == uint.MaxValue;
        bool hasCompressedSize = OfficeProvenanceBinary.ReadUInt32(data, centralHeaderOffset + 20, littleEndian: true) == uint.MaxValue;
        int cursor = 0;
        while (cursor <= extraField.Length - 4) {
            ushort headerId = OfficeProvenanceBinary.ReadUInt16(extraField, cursor, littleEndian: true);
            int dataLength = OfficeProvenanceBinary.ReadUInt16(extraField, cursor + 2, littleEndian: true);
            cursor += 4;
            if (dataLength > extraField.Length - cursor) break;
            if (headerId == 0x0001) {
                int valueOffset = cursor;
                int remaining = dataLength;
                if (hasUncompressedSize) SkipZip64Value(ref valueOffset, ref remaining);
                if (hasCompressedSize) SkipZip64Value(ref valueOffset, ref remaining);
                if (remaining < 8) break;
                ulong resolved = OfficeProvenanceBinary.ReadUInt64(extraField, valueOffset, littleEndian: true);
                if (resolved > uint.MaxValue) {
                    throw new InvalidDataException("ZIP64 local-header offset exceeds the supported package bounds.");
                }
                return (uint)resolved;
            }
            cursor += dataLength;
        }
        throw new InvalidDataException("ZIP64 local-header offset metadata is missing or malformed.");
    }

    private static void SkipZip64Value(ref int offset, ref int remaining) {
        if (remaining < 8) throw new InvalidDataException("ZIP64 entry metadata is truncated.");
        offset += 8;
        remaining -= 8;
    }

    private static bool HasZip64Locator(byte[] data, int endOffset, uint locatorSignature) {
        int locatorOffset = endOffset - 20;
        return locatorOffset >= 0 &&
            OfficeProvenanceBinary.ReadUInt32(data, locatorOffset, littleEndian: true) == locatorSignature;
    }

    private static bool IsNonOpcSignatureEntry(ZipArchiveEntry entry) =>
        !entry.FullName.EndsWith("/", StringComparison.Ordinal) &&
        entry.FullName.StartsWith("META-INF/", StringComparison.Ordinal) &&
        entry.FullName.EndsWith("signatures.xml", StringComparison.OrdinalIgnoreCase);

    internal static bool HasPackageSignature(byte[] data, OfficeProvenanceRemovalOptions options) {
        ValidateEntryCount(data, options.Limits.MaxContainerEntries);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        return HasPackageSignature(data, archive, options);
    }

    internal static void ValidateForOwningPackageMutation(byte[] data, OfficeProvenanceOptions options) {
        ValidateEntryCount(data, options.MaxContainerEntries);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        long expandedBytes = 0;
        byte[] buffer = new byte[81920];
        foreach (ZipArchiveEntry entry in archive.Entries) {
            if (entry.Length > options.MaxAssetBytes) {
                throw new InvalidDataException("A package part exceeds the configured asset limit.");
            }
            using Stream source = entry.Open();
            long entryBytes = 0;
            while (true) {
                int read = source.Read(buffer, 0, buffer.Length);
                if (read <= 0) break;
                if (entryBytes > options.MaxAssetBytes - read) {
                    throw new InvalidDataException("A package part exceeds the configured asset limit.");
                }
                entryBytes += read;
                ReserveExpandedBytes(ref expandedBytes, read, options.MaxExpandedContainerBytes);
            }
            if (entryBytes != entry.Length) throw new InvalidDataException("A ZIP entry expanded to an unexpected length.");
        }
    }

    private static bool HasPackageSignature(
        byte[] data,
        ZipArchive archive,
        OfficeProvenanceRemovalOptions options) {
        if (archive.GetEntry("[Content_Types].xml") == null) return archive.Entries.Any(IsNonOpcSignatureEntry);

        var inspectionOptions = new OfficeIMO.Security.OfficePackageSignatureInspectionOptions {
            MaxPackageBytes = options.Limits.MaxAssetBytes,
            MaxPackageParts = options.Limits.MaxContainerEntries,
            MaxPartBytes = options.Limits.MaxAssetBytes,
            MaxSignatureBytes = options.Limits.MaxAssetBytes,
            MaxTotalDigestBytes = options.Limits.MaxExpandedContainerBytes,
            VerifyDigests = false
        };
        OfficeIMO.Security.OfficePackageSignatureInfo signatureInfo =
            OfficeIMO.Security.OfficePackageSignatureService.Inspect(data, inspectionOptions);
        if (!signatureInfo.SignatureDiscoveryComplete) {
            throw new InvalidDataException("The OPC package signature state could not be determined safely.");
        }
        return signatureInfo.HasSignatures;
    }

    private static bool IsSupportedEmbeddedAsset(
        ZipArchiveEntry entry,
        OfficeProvenanceOptions options,
        ref long expandedBytes) {
        string extension = Path.GetExtension(entry.FullName).ToLowerInvariant();
        if (extension is ".jpg" or ".jpeg" or ".png" or ".webp" or ".gif" or ".tif" or ".tiff" or ".svg") {
            return true;
        }
        if (entry.Length <= 0) return false;

        const int maximumSniffBytes = 16 * 1024;
        int sniffLength = (int)Math.Min(entry.Length, maximumSniffBytes);
        byte[] prefix = new byte[sniffLength];
        using Stream source = entry.Open();
        int read = 0;
        while (read < prefix.Length) {
            int current = source.Read(prefix, read, prefix.Length - read);
            if (current <= 0) break;
            ReserveExpandedBytes(ref expandedBytes, current, options.MaxExpandedContainerBytes);
            read += current;
        }
        if (read != prefix.Length) Array.Resize(ref prefix, read);
        OfficeProvenanceAssetFormat format = OfficeProvenanceInspector.DetectFormat(prefix, entry.FullName, options);
        bool supported = format is OfficeProvenanceAssetFormat.Jpeg or OfficeProvenanceAssetFormat.Png or
            OfficeProvenanceAssetFormat.Webp or OfficeProvenanceAssetFormat.Gif or
            OfficeProvenanceAssetFormat.Tiff or OfficeProvenanceAssetFormat.Svg;
        if (supported) expandedBytes -= read; // The complete entry is charged by the caller.
        return supported;
    }

    private static bool IsSupportedEmbeddedImage(byte[] data, OfficeProvenanceOptions options) {
        OfficeProvenanceAssetFormat format = OfficeProvenanceInspector.DetectFormat(data, fileName: null, options);
        return format is OfficeProvenanceAssetFormat.Jpeg or OfficeProvenanceAssetFormat.Png or
            OfficeProvenanceAssetFormat.Webp or OfficeProvenanceAssetFormat.Gif or
            OfficeProvenanceAssetFormat.Tiff or OfficeProvenanceAssetFormat.Svg;
    }

    private static OfficeProvenanceOptions CreateNestedOptions(OfficeProvenanceOptions source) => new OfficeProvenanceOptions {
        MaxAssetBytes = source.MaxAssetBytes,
        MaxManifestBytes = source.MaxManifestBytes,
        MaxCarriers = source.MaxCarriers,
        MaxContainerEntries = source.MaxContainerEntries,
        MaxExpandedContainerBytes = source.MaxExpandedContainerBytes,
        ProcessEmbeddedAssets = false,
        MaxEmbeddedAssets = source.MaxEmbeddedAssets
    };

    private static OfficeProvenanceRemovalOptions CreateNestedRemovalOptions(OfficeProvenanceRemovalOptions source) {
        var nested = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = source.RemoveC2paManifests,
            RemoveExternalC2paReferences = source.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = source.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = source.RequireStructurallyValidCarrier,
            SignatureMutationPolicy = source.SignatureMutationPolicy,
            ProcessEmbeddedAssets = false,
            MaxEmbeddedAssets = source.MaxEmbeddedAssets
        };
        nested.Limits.MaxAssetBytes = source.Limits.MaxAssetBytes;
        nested.Limits.MaxManifestBytes = source.Limits.MaxManifestBytes;
        nested.Limits.MaxCarriers = source.Limits.MaxCarriers;
        nested.Limits.MaxContainerEntries = source.Limits.MaxContainerEntries;
        nested.Limits.MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes;
        nested.Limits.ProcessEmbeddedAssets = false;
        nested.Limits.MaxEmbeddedAssets = source.Limits.MaxEmbeddedAssets;
        return nested;
    }

    private static OfficeProvenanceEvidence PrefixEvidence(string entryName, OfficeProvenanceEvidence evidence) =>
        new OfficeProvenanceEvidence(
            evidence.Carrier,
            $"ZIP/{entryName}/{evidence.Location}",
            evidence.IsStructurallyValid,
            evidence.PayloadLength,
            evidence.Value,
            evidence.DigitalSourceKind);

    private static byte[] ReadEntry(ZipArchiveEntry entry, int length) {
        byte[] data = new byte[length];
        using Stream stream = entry.Open();
        OfficeProvenanceBinary.ReadExactly(stream, data, 0, data.Length);
        if (stream.ReadByte() != -1) throw new InvalidDataException("ZIP entry expanded beyond its declared length.");
        return data;
    }

    private static void CopyBounded(Stream source, Stream destination, long maximumTotalBytes, ref long totalBytes) {
        byte[] buffer = new byte[81920];
        while (true) {
            int read = source.Read(buffer, 0, buffer.Length);
            if (read <= 0) break;
            ReserveExpandedBytes(ref totalBytes, read, maximumTotalBytes);
            destination.Write(buffer, 0, read);
        }
    }

    private static void ReserveExpandedBytes(ref long totalBytes, long bytes, long maximumBytes) {
        if (bytes < 0 || totalBytes > maximumBytes - bytes) throw new InvalidDataException("ZIP package exceeds the configured expanded-byte limit.");
        totalBytes += bytes;
    }

    private sealed class OfficeProvenanceZipEntryMetadata {
        internal OfficeProvenanceZipEntryMetadata(byte[] localExtraField, byte[] centralExtraField, byte[] comment, ushort internalAttributes) {
            LocalExtraField = localExtraField;
            CentralExtraField = centralExtraField;
            Comment = comment;
            InternalAttributes = internalAttributes;
        }

        internal byte[] LocalExtraField { get; }
        internal byte[] CentralExtraField { get; }
        internal byte[] Comment { get; }
        internal ushort InternalAttributes { get; }
    }
}
