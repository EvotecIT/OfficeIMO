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

    internal static void ValidateMimetypeEntry(byte[] data, string expectedValue, int maximumEntries) =>
        ValidateMimetypeEntry(data, new[] { expectedValue }, maximumEntries);

    internal static void ValidateMimetypeEntry(byte[] data, IReadOnlyCollection<string> expectedValues, int maximumEntries) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (expectedValues == null) throw new ArgumentNullException(nameof(expectedValues));
        if (expectedValues.Count == 0 || expectedValues.Any(string.IsNullOrEmpty)) throw new ArgumentException("At least one expected mimetype value is required.", nameof(expectedValues));
        ValidateEntryCount(data, maximumEntries);
        byte[] expectedName = Encoding.ASCII.GetBytes("mimetype");
        if (data.Length < 30 || OfficeProvenanceBinary.ReadUInt32(data, 0, littleEndian: true) != 0x04034B50U) {
            throw new InvalidDataException("The package does not start with a local mimetype entry.");
        }
        ushort flags = OfficeProvenanceBinary.ReadUInt16(data, 6, littleEndian: true);
        ushort compressionMethod = OfficeProvenanceBinary.ReadUInt16(data, 8, littleEndian: true);
        int nameLength = OfficeProvenanceBinary.ReadUInt16(data, 26, littleEndian: true);
        int extraFieldLength = OfficeProvenanceBinary.ReadUInt16(data, 28, littleEndian: true);
        if ((flags & 0x0001) != 0 || compressionMethod != 0 || nameLength != expectedName.Length || extraFieldLength != 0 ||
            !BytesEqual(data, 30, expectedName) ||
            !HasExactFirstEntryContent(data, expectedValues)) {
            throw new InvalidDataException("The package does not contain the required leading mimetype entry.");
        }
    }

    private static bool HasExactFirstEntryContent(byte[] data, IReadOnlyCollection<string> expectedValues) {
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        if (archive.Entries.Count == 0 || archive.Entries[0].Length > 256 || archive.Entries[0].Length > int.MaxValue) return false;
        byte[] content = ReadEntry(archive.Entries[0], (int)archive.Entries[0].Length);
        return expectedValues.Any(value => content.SequenceEqual(Encoding.ASCII.GetBytes(value)));
    }

    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        ValidateEntryCount(data, options.MaxContainerEntries);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        Dictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata = GetEntryMetadata(data, archive);
        bool hasDuplicateManifests = CountManifestEntries(archive, entryMetadata) > 1;
        int index = 0;
        int embeddedCount = 0;
        long expandedBytes = 0;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            string entryName = entryMetadata[entry].Name;
            if (entryName.Equals(ManifestPath, StringComparison.Ordinal)) {
                if (entry.Length > options.MaxManifestBytes || entry.Length > int.MaxValue) {
                    throw new InvalidDataException("ZIP provenance manifest exceeds the configured manifest limit.");
                }
                ReserveExpandedBytes(ref expandedBytes, entry.Length, options.MaxExpandedContainerBytes);
                byte[] manifest = ReadEntry(entry, (int)entry.Length);
                bool valid = !hasDuplicateManifests && OfficeC2paManifestStore.IsValid(
                    manifest, 0, manifest.Length, options.MaxManifestBytes, options.MaxContainerEntries, out _);
                context.Add(new OfficeProvenanceEvidence(
                    OfficeProvenanceCarrierKind.C2paManifest,
                    $"ZIP/{ManifestPath}[{index++}]",
                    valid,
                    manifest.Length));
                continue;
            }
            if (!options.ProcessEmbeddedAssets ||
                !IsSupportedEmbeddedAsset(entry, entryName, options, ref expandedBytes)) continue;
            if (entry.Length > options.MaxAssetBytes || entry.Length > int.MaxValue) {
                throw new InvalidDataException("A supported embedded asset exceeds the configured asset limit.");
            }
            ReserveExpandedBytes(ref expandedBytes, entry.Length, options.MaxExpandedContainerBytes);
            byte[] asset = ReadEntry(entry, (int)entry.Length);
            if (!IsSupportedEmbeddedImage(asset, options)) continue;
            embeddedCount++;
            if (embeddedCount > options.MaxEmbeddedAssets) throw new InvalidDataException("ZIP package exceeds the configured embedded-asset limit.");
            OfficeProvenanceReport nested;
            try {
                nested = OfficeProvenanceInspector.InspectCore(asset, entryName, CreateNestedOptions(options));
            } catch (Exception exception) when (exception is InvalidDataException || exception is XmlException) {
                context.Diagnostics.Add($"ZIP/{entryName}: embedded asset was preserved because inspection failed: {exception.Message}");
                continue;
            }
            foreach (OfficeProvenanceEvidence evidence in nested.Evidence) context.Add(PrefixEvidence(entryName, evidence));
            foreach (string diagnostic in nested.Diagnostics) context.Diagnostics.Add($"ZIP/{entryName}: {diagnostic}");
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
        Dictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata = GetEntryMetadata(data, input);
        bool hasDuplicateManifests = CountManifestEntries(input, entryMetadata) > 1;
        var removable = new HashSet<string>(StringComparer.Ordinal);
        var embeddedRewrites = new Dictionary<ZipArchiveEntry, byte[]>();
        int occurrence = 0;
        int embeddedCount = 0;
        long inspectionBytes = 0;
        foreach (ZipArchiveEntry entry in input.Entries) {
            string entryName = entryMetadata[entry].Name;
            if (entryName.Equals(ManifestPath, StringComparison.Ordinal)) {
                if (entry.Length > options.Limits.MaxManifestBytes || entry.Length > int.MaxValue) throw new InvalidDataException("ZIP provenance manifest exceeds the configured limit.");
                ReserveExpandedBytes(ref inspectionBytes, entry.Length, options.Limits.MaxExpandedContainerBytes);
                byte[] manifest = ReadEntry(entry, (int)entry.Length);
                bool valid = !hasDuplicateManifests && OfficeC2paManifestStore.IsValid(
                    manifest, 0, manifest.Length, options.Limits.MaxManifestBytes, options.Limits.MaxContainerEntries, out _);
                if (options.RemoveC2paManifests && (valid || !options.RequireStructurallyValidCarrier)) {
                    removable.Add(entryName + "\0" + occurrence);
                    changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, $"ZIP/{ManifestPath}[{occurrence}]", 0));
                }
                occurrence++;
                continue;
            }
            if (!options.ProcessEmbeddedAssets || !options.Limits.ProcessEmbeddedAssets ||
                !IsSupportedEmbeddedAsset(entry, entryName, options.Limits, ref inspectionBytes)) continue;
            if (entry.Length > options.Limits.MaxAssetBytes || entry.Length > int.MaxValue) throw new InvalidDataException("A supported embedded asset exceeds the configured asset limit.");
            ReserveExpandedBytes(ref inspectionBytes, entry.Length, options.Limits.MaxExpandedContainerBytes);
            byte[] asset = ReadEntry(entry, (int)entry.Length);
            if (!IsSupportedEmbeddedImage(asset, options.Limits)) continue;
            embeddedCount++;
            if (embeddedCount > Math.Min(options.MaxEmbeddedAssets, options.Limits.MaxEmbeddedAssets)) throw new InvalidDataException("ZIP package exceeds the configured embedded-asset limit.");
            OfficeProvenanceRemovalResult nested;
            try {
                nested = OfficeProvenanceRemover.Remove(asset, entryName, CreateNestedRemovalOptions(options));
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
                changes.Add(new OfficeProvenanceChange(change.Carrier, $"ZIP/{entryName}/{change.Location}", 0));
            }
            if (nested.WasReserialized) reserialized = true;
        }
        if (changes.Count == 0) return (byte[])data.Clone();

        bool hasSignature = options.SignatureMutationPolicy != OfficeIMO.OfficeSignatureMutationPolicy.PreserveSignatureMarkup &&
            HasPackageSignature(data, input, options, ref inspectionBytes);
        if (hasSignature && options.SignatureMutationPolicy == OfficeIMO.OfficeSignatureMutationPolicy.BlockSave) {
            throw new InvalidOperationException("Removing provenance would invalidate package signatures. Choose an explicit signature mutation policy.");
        }
        if (hasSignature && options.SignatureMutationPolicy == OfficeIMO.OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures) {
            throw new NotSupportedException("Generic ZIP provenance removal cannot safely rewrite package signatures. Use the owning OfficeIMO document format API.");
        }

        if (removable.Count != 0 && removable.Count == occurrence &&
            entryMetadata.Values.Any(metadata => metadata.Name.Equals("[Content_Types].xml", StringComparison.Ordinal))) {
            RemoveOpcManifestReferences(input, entryMetadata, embeddedRewrites, options.Limits, ref inspectionBytes);
        }

        var outputEntries = new List<OfficeProvenanceZipWriteEntry>();
        occurrence = 0;
        IEnumerable<ZipArchiveEntry> orderedEntries = input.Entries
            .OrderByDescending(entry => entryMetadata[entry].Name.Equals("mimetype", StringComparison.Ordinal));
        foreach (ZipArchiveEntry entry in orderedEntries) {
            string entryName = entryMetadata[entry].Name;
            bool isManifest = entryName.Equals(ManifestPath, StringComparison.Ordinal);
            string key = entryName + "\0" + occurrence;
            if (isManifest) occurrence++;
            if (isManifest && removable.Contains(key)) continue;
            embeddedRewrites.TryGetValue(entry, out byte[]? rewritten);
            outputEntries.Add(CreateWriteEntry(entry, entryMetadata[entry], rewritten));
        }
        reserialized = true;
        return OfficeProvenanceZipWriter.Write(outputEntries, options.Limits.MaxExpandedContainerBytes, ReadArchiveComment(data));
    }

    private static int CountManifestEntries(
        ZipArchive archive,
        IReadOnlyDictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata) {
        int count = 0;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            if (entryMetadata[entry].Name.Equals(ManifestPath, StringComparison.Ordinal) && ++count > 1) return count;
        }
        return count;
    }

    private static void RemoveOpcManifestReferences(
        ZipArchive archive,
        IReadOnlyDictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata,
        IDictionary<ZipArchiveEntry, byte[]> replacements,
        OfficeProvenanceOptions limits,
        ref long expandedBytes) {
        foreach (ZipArchiveEntry entry in archive.Entries) {
            string entryName = entryMetadata[entry].Name;
            bool isContentTypes = entryName.Equals("[Content_Types].xml", StringComparison.Ordinal);
            bool isRelationships = entryName.EndsWith(".rels", StringComparison.Ordinal) &&
                (entryName.StartsWith("_rels/", StringComparison.Ordinal) || entryName.Contains("/_rels/"));
            if (!isContentTypes && !isRelationships) continue;
            if (entry.Length > limits.MaxAssetBytes || entry.Length > int.MaxValue) {
                throw new InvalidDataException("An OPC relationship metadata part exceeds the configured asset limit.");
            }
            ReserveExpandedBytes(ref expandedBytes, entry.Length, limits.MaxExpandedContainerBytes);
            byte[] original = ReadEntry(entry, (int)entry.Length);
            OfficeProvenanceXml.ValidateMaterializedNodeBudget(original, limits, "OPC relationship metadata");
            var document = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using (var stream = new MemoryStream(original, writable: false))
            using (XmlReader reader = XmlReader.Create(stream, OfficeProvenanceXml.CreateReaderSettings(limits))) {
                document.Load(reader);
            }
            bool changed = false;
            if (isContentTypes) {
                foreach (XmlElement element in document.GetElementsByTagName("Override", "http://schemas.openxmlformats.org/package/2006/content-types").OfType<XmlElement>().ToArray()) {
                    string partName = element.GetAttribute("PartName");
                    if (TryNormalizeOpcTarget(string.Empty, partName, out string normalized) && normalized == ManifestPath) {
                        element.ParentNode?.RemoveChild(element);
                        changed = true;
                    }
                }
            } else {
                string ownerDirectory = GetOpcRelationshipOwnerDirectory(entryName);
                foreach (XmlElement relationship in document.GetElementsByTagName("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships").OfType<XmlElement>().ToArray()) {
                    if (string.Equals(relationship.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                    if (TryNormalizeOpcTarget(ownerDirectory, relationship.GetAttribute("Target"), out string normalized) && normalized == ManifestPath) {
                        relationship.ParentNode?.RemoveChild(relationship);
                        changed = true;
                    }
                }
            }
            if (!changed) continue;
            using var output = new MemoryStream();
            using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
                Encoding = new UTF8Encoding(false),
                Indent = false,
                OmitXmlDeclaration = document.FirstChild is not XmlDeclaration
            })) document.Save(writer);
            replacements[entry] = output.ToArray();
        }
    }

    private static string GetOpcRelationshipOwnerDirectory(string relationshipPart) {
        int marker = relationshipPart.LastIndexOf("/_rels/", StringComparison.Ordinal);
        if (marker >= 0) return relationshipPart.Substring(0, marker);
        return string.Empty;
    }

    private static bool TryNormalizeOpcTarget(string ownerDirectory, string target, out string normalized) {
        normalized = string.Empty;
        if (string.IsNullOrWhiteSpace(target)) return false;
        int suffix = target.IndexOfAny(new[] { '?', '#' });
        if (suffix >= 0) target = target.Substring(0, suffix);
        try { target = Uri.UnescapeDataString(target); } catch (UriFormatException) { return false; }
        if (target.IndexOf('\\') >= 0) return false;
        string combined = target.StartsWith("/", StringComparison.Ordinal)
            ? target.TrimStart('/')
            : (ownerDirectory.Length == 0 ? target : ownerDirectory + "/" + target);
        var segments = new List<string>();
        foreach (string segment in combined.Split('/')) {
            if (segment.Length == 0 || segment == ".") continue;
            if (segment == "..") {
                if (segments.Count == 0) return false;
                segments.RemoveAt(segments.Count - 1);
            } else segments.Add(segment);
        }
        normalized = string.Join("/", segments);
        return true;
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
        bool isMimetype = metadata.Name.Equals("mimetype", StringComparison.Ordinal);
        if (replacement != null) {
            return new OfficeProvenanceZipWriteEntry(
                metadata.Name,
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
            metadata.Name,
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
        Func<string, bool> shouldRemove,
        Func<string, bool>? shouldReplace = null,
        Func<string, byte[], byte[]>? replace = null,
        long maximumReplacementBytes = long.MaxValue) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (shouldRemove == null) throw new ArgumentNullException(nameof(shouldRemove));
        using var inputStream = new MemoryStream(data, writable: false);
        using var input = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: false);
        Dictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata = GetEntryMetadata(data, input);
        bool hadMatches = input.Entries.Any(entry => shouldRemove(entryMetadata[entry].Name));
        bool hadReplacements = shouldReplace != null && input.Entries.Any(entry => shouldReplace(entryMetadata[entry].Name));
        if (!hadMatches && !hadReplacements) return new OfficeProvenanceSignatureStripResult((byte[])data.Clone(), hadSignatures: false);

        var outputEntries = new List<OfficeProvenanceZipWriteEntry>();
        foreach (ZipArchiveEntry entry in input.Entries
            .OrderByDescending(candidate => entryMetadata[candidate].Name.Equals("mimetype", StringComparison.Ordinal))) {
            string entryName = entryMetadata[entry].Name;
            if (shouldRemove(entryName)) continue;
            byte[]? replacement = null;
            if (shouldReplace?.Invoke(entryName) == true) {
                if (replace == null || entry.Length > maximumReplacementBytes || entry.Length > int.MaxValue) {
                    throw new InvalidDataException("A package metadata entry exceeds its configured rewrite limit.");
                }
                replacement = replace(entryName, ReadEntry(entry, (int)entry.Length));
                if (replacement.LongLength > maximumReplacementBytes) {
                    throw new InvalidDataException("A rewritten package metadata entry exceeds its configured rewrite limit.");
                }
            }
            outputEntries.Add(CreateWriteEntry(entry, entryMetadata[entry], replacement));
        }
        return new OfficeProvenanceSignatureStripResult(
            OfficeProvenanceZipWriter.Write(outputEntries, long.MaxValue, ReadArchiveComment(data)),
            hadSignatures: hadMatches);
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
            byte[] rawName = new byte[nameLength];
            if (nameLength != 0) Buffer.BlockCopy(data, cursor + 46, rawName, 0, nameLength);
            string decodedName = DecodeZipEntryName(rawName, flags, centralExtraField);
            if (localHeaderOffset == uint.MaxValue) {
                localHeaderOffset = ResolveZip64LocalHeaderOffset(data, cursor, centralExtraField);
            }
            if (localHeaderOffset > int.MaxValue) {
                throw new InvalidDataException("ZIP local-header offset exceeds the supported package bounds.");
            }
            byte[] comment = new byte[commentLength];
            if (commentLength != 0) Buffer.BlockCopy(data, cursor + 46 + nameLength + extraLength, comment, 0, commentLength);
            if ((flags & 0x0800) == 0) {
                comment = TryReadUnicodeComment(centralExtraField, comment, out byte[]? unicodeComment)
                    ? unicodeComment!
                    : TranscodeLegacyZipComment(comment);
            }
            int localOffset = (int)localHeaderOffset;
            if (localOffset > data.Length - 30 || OfficeProvenanceBinary.ReadUInt32(data, localOffset, littleEndian: true) != 0x04034B50U) {
                throw new InvalidDataException("ZIP local-file header is outside the package bounds.");
            }
            int localNameLength = OfficeProvenanceBinary.ReadUInt16(data, localOffset + 26, littleEndian: true);
            int localExtraLength = OfficeProvenanceBinary.ReadUInt16(data, localOffset + 28, littleEndian: true);
            long localHeaderEnd = (long)localOffset + 30 + localNameLength + localExtraLength;
            if (localHeaderEnd > data.Length) throw new InvalidDataException("ZIP local-file header exceeds the package bounds.");
            if (localNameLength != rawName.Length || !BytesEqual(data, localOffset + 30, rawName)) {
                throw new InvalidDataException("ZIP local and central entry names do not match.");
            }
            byte[] localExtraField = new byte[localExtraLength];
            if (localExtraLength != 0) Buffer.BlockCopy(data, localOffset + 30 + localNameLength, localExtraField, 0, localExtraLength);
            metadata.Add(new OfficeProvenanceZipEntryMetadata(decodedName, localExtraField, centralExtraField, comment, internalAttributes));
            cursor = (int)recordEnd;
        }
        return metadata;
    }

    private static byte[] TranscodeLegacyZipComment(byte[] comment) {
        if (comment.Length == 0) return comment;
        char[] decoded = DecodeCp437(comment);
        int encodedLength = Encoding.UTF8.GetByteCount(decoded);
        if (encodedLength <= ushort.MaxValue) return Encoding.UTF8.GetBytes(decoded);
        throw new InvalidDataException("A legacy ZIP entry comment cannot be represented completely as UTF-8 within the ZIP length limit.");
    }

    private static string DecodeZipEntryName(byte[] rawName, ushort flags, byte[] extraField) {
        if ((flags & 0x0800) != 0) {
            try { return new UTF8Encoding(false, true).GetString(rawName); }
            catch (DecoderFallbackException exception) { throw new InvalidDataException("A UTF-8 ZIP entry name is malformed.", exception); }
        }
        if (TryReadUnicodePath(extraField, rawName, out string? unicodeName)) return unicodeName!;
        return new string(DecodeCp437(rawName));
    }

    private static bool TryReadUnicodePath(byte[] extraField, byte[] rawName, out string? name) {
        name = null;
        int cursor = 0;
        while (cursor <= extraField.Length - 4) {
            ushort fieldId = OfficeProvenanceBinary.ReadUInt16(extraField, cursor, littleEndian: true);
            int dataLength = OfficeProvenanceBinary.ReadUInt16(extraField, cursor + 2, littleEndian: true);
            cursor += 4;
            if (dataLength > extraField.Length - cursor) return false;
            if (fieldId == 0x7075 && dataLength >= 5 && extraField[cursor] == 1 &&
                OfficeProvenanceBinary.ReadUInt32(extraField, cursor + 1, littleEndian: true) == ComputeCrc32(rawName)) {
                try {
                    name = new UTF8Encoding(false, true).GetString(extraField, cursor + 5, dataLength - 5);
                    return true;
                } catch (DecoderFallbackException) { return false; }
            }
            cursor += dataLength;
        }
        return false;
    }

    private static bool TryReadUnicodeComment(byte[] extraField, byte[] rawComment, out byte[]? comment) {
        comment = null;
        int cursor = 0;
        while (cursor <= extraField.Length - 4) {
            ushort fieldId = OfficeProvenanceBinary.ReadUInt16(extraField, cursor, littleEndian: true);
            int dataLength = OfficeProvenanceBinary.ReadUInt16(extraField, cursor + 2, littleEndian: true);
            cursor += 4;
            if (dataLength > extraField.Length - cursor) return false;
            if (fieldId == 0x6375 && dataLength >= 5 && extraField[cursor] == 1 &&
                OfficeProvenanceBinary.ReadUInt32(extraField, cursor + 1, littleEndian: true) == ComputeCrc32(rawComment)) {
                try {
                    new UTF8Encoding(false, true).GetCharCount(extraField, cursor + 5, dataLength - 5);
                    comment = new byte[dataLength - 5];
                    Buffer.BlockCopy(extraField, cursor + 5, comment, 0, comment.Length);
                    return true;
                } catch (DecoderFallbackException) { return false; }
            }
            cursor += dataLength;
        }
        return false;
    }

    private static char[] DecodeCp437(byte[] bytes) {
        const string highCharacters =
            "ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»" +
            "░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌" +
            "█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ";
        var decoded = new char[bytes.Length];
        for (int index = 0; index < bytes.Length; index++) {
            byte value = bytes[index];
            decoded[index] = value < 0x80 ? (char)value : highCharacters[value - 0x80];
        }
        return decoded;
    }

    private static uint ComputeCrc32(byte[] bytes) {
        uint crc = uint.MaxValue;
        foreach (byte value in bytes) {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
        }
        return ~crc;
    }

    private static bool BytesEqual(byte[] data, int offset, byte[] expected) {
        if (offset < 0 || offset > data.Length - expected.Length) return false;
        for (int index = 0; index < expected.Length; index++) if (data[offset + index] != expected[index]) return false;
        return true;
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

    private static bool IsNonOpcSignatureEntry(string entryName) =>
        !entryName.EndsWith("/", StringComparison.Ordinal) &&
        entryName.StartsWith("META-INF/", StringComparison.Ordinal) &&
        entryName.EndsWith("signatures.xml", StringComparison.OrdinalIgnoreCase);

    private static bool IsOpcSignatureEvidenceEntry(string entryName) {
        if (entryName.EndsWith("/", StringComparison.Ordinal) ||
            !entryName.StartsWith("_xmlsignatures/", StringComparison.Ordinal)) return false;
        string relativeName = entryName.Substring("_xmlsignatures/".Length);
        return relativeName.Equals("origin.sigs", StringComparison.Ordinal) ||
            relativeName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
            relativeName.StartsWith("_rels/", StringComparison.Ordinal) &&
                relativeName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasOpcSignatureOriginRelationship(
        ZipArchive archive,
        IReadOnlyDictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata,
        long maximumBytes,
        long maximumExpandedBytes,
        ref long expandedBytes) {
        byte[] marker = System.Text.Encoding.ASCII.GetBytes(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin");
        foreach (ZipArchiveEntry relationships in archive.Entries) {
            if (!entryMetadata[relationships].Name.Equals("_rels/.rels", StringComparison.Ordinal) ||
                relationships.Length <= 0) continue;
            if (relationships.Length > maximumBytes) {
                throw new InvalidDataException("The OPC root relationships part exceeds the configured asset limit.");
            }
            ReserveExpandedBytes(ref expandedBytes, relationships.Length, maximumExpandedBytes);
            using Stream source = relationships.Open();
            byte[] bytes = OfficeProvenanceBinary.ReadBounded(source, maximumBytes);
            for (int offset = 0; offset <= bytes.Length - marker.Length; offset++) {
                int index = 0;
                while (index < marker.Length && bytes[offset + index] == marker[index]) index++;
                if (index == marker.Length) return true;
            }
        }
        return false;
    }

    internal static bool HasPackageSignature(byte[] data, OfficeProvenanceRemovalOptions options) {
        ValidateEntryCount(data, options.Limits.MaxContainerEntries);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        long expandedBytes = 0;
        return HasPackageSignature(data, archive, options, ref expandedBytes);
    }

    internal static bool HasApplicationSignatureMetadata(byte[] data, OfficeProvenanceOptions options) {
        ValidateEntryCount(data, options.MaxContainerEntries);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        Dictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata = GetEntryMetadata(data, archive);
        long expandedBytes = 0;
        var applicationMetadataParts = new HashSet<string>(StringComparer.Ordinal) { "docProps/app.xml" };
        ZipArchiveEntry[] rootRelationships = archive.Entries
            .Where(entry => entryMetadata[entry].Name.Equals("_rels/.rels", StringComparison.Ordinal))
            .ToArray();
        if (rootRelationships.Length > 1) throw new InvalidDataException("The OPC package contains duplicate root relationship parts.");
        if (rootRelationships.Length == 1) {
            ZipArchiveEntry relationships = rootRelationships[0];
            if (relationships.Length > options.MaxAssetBytes || relationships.Length > int.MaxValue) {
                throw new InvalidDataException("The OPC root relationships part exceeds the configured asset limit.");
            }
            ReserveExpandedBytes(ref expandedBytes, relationships.Length, options.MaxExpandedContainerBytes);
            byte[] relationshipsXml = ReadEntry(relationships, (int)relationships.Length);
            OfficeProvenanceXml.ValidateMaterializedNodeBudget(relationshipsXml, options, "OPC root relationships");
            using var relationshipsInput = new MemoryStream(relationshipsXml, writable: false);
            using XmlReader relationshipsReader = XmlReader.Create(relationshipsInput, OfficeProvenanceXml.CreateReaderSettings(options));
            while (relationshipsReader.Read()) {
                if (relationshipsReader.NodeType != XmlNodeType.Element ||
                    relationshipsReader.LocalName != "Relationship" ||
                    relationshipsReader.NamespaceURI != "http://schemas.openxmlformats.org/package/2006/relationships" ||
                    relationshipsReader.GetAttribute("Type") != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" ||
                    string.Equals(relationshipsReader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                if (!TryNormalizeOpcTarget(string.Empty, relationshipsReader.GetAttribute("Target") ?? string.Empty, out string target)) {
                    throw new InvalidDataException("The OPC extended-properties relationship has an invalid target.");
                }
                applicationMetadataParts.Add(target);
            }
        }

        foreach (string partName in applicationMetadataParts) {
            ZipArchiveEntry[] entries = archive.Entries
                .Where(entry => entryMetadata[entry].Name.Equals(partName, StringComparison.Ordinal))
                .ToArray();
            if (entries.Length == 0) continue;
            if (entries.Length != 1) throw new InvalidDataException("The OPC package contains duplicate application metadata parts.");
            ZipArchiveEntry entry = entries[0];
            if (entry.Length == 0) continue;
            if (entry.Length > options.MaxAssetBytes || entry.Length > int.MaxValue) {
                throw new InvalidDataException("Open XML application metadata exceeds the configured asset limit.");
            }
            ReserveExpandedBytes(ref expandedBytes, entry.Length, options.MaxExpandedContainerBytes);
            byte[] xml = ReadEntry(entry, (int)entry.Length);
            OfficeProvenanceXml.ValidateMaterializedNodeBudget(xml, options, "Open XML application metadata");
            using var input = new MemoryStream(xml, writable: false);
            using XmlReader reader = XmlReader.Create(input, OfficeProvenanceXml.CreateReaderSettings(options));
            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.Element &&
                    reader.LocalName == "DigSig" &&
                    reader.NamespaceURI == "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties") return true;
            }
        }
        return false;
    }

    internal static void ValidateForOwningPackageMutation(byte[] data, OfficeProvenanceOptions options) {
        ValidateEntryCount(data, options.MaxContainerEntries);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        Dictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata = GetEntryMetadata(data, archive);
        long expandedBytes = 0;
        byte[] buffer = new byte[81920];
        byte[]? rootRelationships = null;
        byte[]? contentTypes = null;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            string entryName = entryMetadata[entry].Name;
            if (entry.Length > options.MaxAssetBytes) {
                throw new InvalidDataException("A package part exceeds the configured asset limit.");
            }
            using Stream source = entry.Open();
            MemoryStream? metadata = entryName.Equals("_rels/.rels", StringComparison.Ordinal) ||
                entryName.Equals("[Content_Types].xml", StringComparison.Ordinal)
                ? new MemoryStream(entry.Length > int.MaxValue ? 0 : (int)entry.Length)
                : null;
            long entryBytes = 0;
            try {
                while (true) {
                    int read = source.Read(buffer, 0, buffer.Length);
                    if (read <= 0) break;
                    if (entryBytes > options.MaxAssetBytes - read) {
                        throw new InvalidDataException("A package part exceeds the configured asset limit.");
                    }
                    entryBytes += read;
                    ReserveExpandedBytes(ref expandedBytes, read, options.MaxExpandedContainerBytes);
                    metadata?.Write(buffer, 0, read);
                }
                if (entryBytes != entry.Length) throw new InvalidDataException("A ZIP entry expanded to an unexpected length.");
                if (metadata != null) {
                    byte[] xml = metadata.ToArray();
                    if (entryName.Equals("_rels/.rels", StringComparison.Ordinal)) {
                        if (rootRelationships != null) throw new InvalidDataException("The OPC package contains duplicate root relationships parts.");
                        rootRelationships = xml;
                    } else {
                        if (contentTypes != null) throw new InvalidDataException("The OPC package contains duplicate content-types parts.");
                        contentTypes = xml;
                    }
                }
            } finally {
                metadata?.Dispose();
            }
        }
        if (rootRelationships != null) OfficeProvenanceXml.ValidateMaterializedNodeBudget(rootRelationships, options, "OPC root relationships");
        if (contentTypes != null) OfficeProvenanceXml.ValidateMaterializedNodeBudget(contentTypes, options, "OPC content types");
    }

    private static bool HasPackageSignature(
        byte[] data,
        ZipArchive archive,
        OfficeProvenanceRemovalOptions options,
        ref long expandedBytes) {
        Dictionary<ZipArchiveEntry, OfficeProvenanceZipEntryMetadata> entryMetadata = GetEntryMetadata(data, archive);
        bool rawSignatureEvidence = archive.Entries.Any(entry => IsNonOpcSignatureEntry(entryMetadata[entry].Name)) ||
            archive.Entries.Any(entry => IsOpcSignatureEvidenceEntry(entryMetadata[entry].Name)) ||
            HasOpcSignatureOriginRelationship(
                archive,
                entryMetadata,
                options.Limits.MaxAssetBytes,
                options.Limits.MaxExpandedContainerBytes,
                ref expandedBytes);
        if (!entryMetadata.Values.Any(metadata => metadata.Name.Equals("[Content_Types].xml", StringComparison.Ordinal))) {
            return rawSignatureEvidence;
        }

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
        return signatureInfo.HasSignatures || rawSignatureEvidence;
    }

    private static bool IsSupportedEmbeddedAsset(
        ZipArchiveEntry entry,
        string entryName,
        OfficeProvenanceOptions options,
        ref long expandedBytes) {
        string extension = Path.GetExtension(entryName).ToLowerInvariant();
        if (extension is ".jpg" or ".jpeg" or ".png" or ".webp" or ".gif" or ".tif" or ".tiff" or ".svg") {
            return true;
        }
        if (entry.Length <= 0) return false;

        if (entry.Length > options.MaxAssetBytes || entry.Length > int.MaxValue) return false;
        int sniffLength = (int)entry.Length;
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
        OfficeProvenanceAssetFormat format = OfficeProvenanceInspector.DetectFormat(prefix, entryName, options);
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
        internal OfficeProvenanceZipEntryMetadata(string name, byte[] localExtraField, byte[] centralExtraField, byte[] comment, ushort internalAttributes) {
            Name = name;
            LocalExtraField = localExtraField;
            CentralExtraField = centralExtraField;
            Comment = comment;
            InternalAttributes = internalAttributes;
        }

        internal string Name { get; }
        internal byte[] LocalExtraField { get; }
        internal byte[] CentralExtraField { get; }
        internal byte[] Comment { get; }
        internal ushort InternalAttributes { get; }
    }
}
