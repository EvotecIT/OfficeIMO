using System.Runtime.CompilerServices;

namespace OfficeIMO.IWork.Internal;

internal static class IWorkDrawingReader {
    private const uint PackageMetadataArchive = 11006;
    private const uint ImageArchive = 3005;
    private static readonly ConditionalWeakTable<IWorkSourceDocument, ImageLookup> ImageLookups = new();

    internal static IWorkGeometry? ReadGeometry(IWorkWireMessage drawable) {
        return ReadGeometry(drawable, out _);
    }

    internal static string? ReadOptionalString(IWorkWireMessage? message, int field,
        IWorkProjectionBudget projectionBudget, ref bool complete) {
        if (message == null) return null;
        string? value = message.GetString(field, out bool fieldComplete);
        if (!fieldComplete) complete = false;
        if (value != null) projectionBudget.AddTextCharacters(value.Length);
        return value;
    }

    internal static IWorkGeometry? ReadGeometry(IWorkWireMessage drawable, out bool complete) {
        IWorkWireMessage? geometry = IWorkObjectIndex.TryGetMessage(drawable, 1, out bool malformedGeometry);
        complete = !malformedGeometry
            && !drawable.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            && (!drawable.HasField(1) || geometry != null);
        if (geometry == null) return null;
        IWorkGeometry? result = ReadGeometryArchive(geometry, out bool archiveComplete);
        if (!archiveComplete || result == null) complete = false;
        return result;
    }

    internal static IWorkGeometry? ReadGeometryArchive(IWorkWireMessage geometry) {
        return ReadGeometryArchive(geometry, out _);
    }

    internal static IWorkGeometry? ReadGeometryArchive(IWorkWireMessage geometry, out bool complete) {
        IWorkWireMessage? position = IWorkObjectIndex.TryGetMessage(geometry, 1, out bool malformedPosition);
        IWorkWireMessage? size = IWorkObjectIndex.TryGetMessage(geometry, 2, out bool malformedSize);
        complete = !malformedPosition && !malformedSize
            && !geometry.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            && !geometry.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
            && (!geometry.HasField(1) || position != null)
            && (!geometry.HasField(2) || size != null);
        if (!complete) return null;
        double left = position?.GetFloat(1) ?? 0;
        double top = position?.GetFloat(2) ?? 0;
        double width = size?.GetFloat(1) ?? 0;
        double height = size?.GetFloat(2) ?? 0;
        double rotation = geometry.GetFloat(4) ?? 0;
        if (position != null && (InvalidFloat(position, 1) || InvalidFloat(position, 2))
            || size != null && (InvalidFloat(size, 1) || InvalidFloat(size, 2))
            || InvalidFloat(geometry, 4)) {
            complete = false;
            return null;
        }
        if (!IsFinite(left) || !IsFinite(top) || !IsFinite(width) || !IsFinite(height)
            || !IsFinite(rotation) || width < 0 || height < 0) {
            complete = false;
            return null;
        }
        return new IWorkGeometry(left, top, width, height, rotation);
    }

    internal static IWorkWireMessage? DrawableMessage(IWorkObjectIndex index, IWorkArchiveRecord record) {
        return DrawableMessage(index, record, out _);
    }

    internal static IWorkWireMessage? DrawableMessage(IWorkObjectIndex index, IWorkArchiveRecord record,
        out bool complete) {
        complete = true;
        IWorkWireMessage message = index.Message(record);
        if (record.MessageType is ImageArchive or 6000) {
            IWorkWireMessage? drawable = IWorkObjectIndex.TryGetMessage(message, 1, out bool malformedDrawable);
            complete = !malformedDrawable
                && !message.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
                && (!message.HasField(1) || drawable != null);
            return drawable;
        }
        if (record.MessageType == 2011) {
            IWorkWireMessage? shape = IWorkObjectIndex.TryGetMessage(message, 1, out bool malformedShape);
            if (malformedShape || message.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
                || message.HasField(1) && shape == null) {
                complete = false;
                return null;
            }
            bool malformedDrawable = false;
            IWorkWireMessage? drawable = shape == null
                ? null
                : IWorkObjectIndex.TryGetMessage(shape, 1, out malformedDrawable);
            complete = shape == null || !malformedDrawable
                && !shape.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
                && (!shape.HasField(1) || drawable != null);
            return drawable;
        }
        return null;
    }

    internal static IWorkImageAsset? ReadImage(IWorkSourceDocument source,
        IWorkArchiveRecord record, IWorkProjectionBudget projectionBudget, out bool complete) {
        complete = true;
        if (record.MessageType != ImageArchive) return null;
        IWorkWireMessage message = source.Index.Message(record);
        IWorkWireMessage? drawable = IWorkObjectIndex.TryGetMessage(message, 1, out bool malformedDrawable);
        if (malformedDrawable || message.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            || drawable == null) {
            complete = false;
            return null;
        }
        var dataIdentifiers = new List<ulong>(6);
        var seenIdentifiers = new HashSet<ulong>();
        foreach (int field in new[] { 11, 15, 17, 13, 12, 23 }) {
            ulong? candidate = DataIdentifier(message, field, out bool identifierComplete);
            if (!identifierComplete) {
                complete = false;
                return null;
            }
            if (candidate.HasValue && seenIdentifiers.Add(candidate.Value)) {
                dataIdentifiers.Add(candidate.Value);
            }
        }
        if (dataIdentifiers.Count == 0) {
            complete = false;
            return null;
        }
        ImageLookup lookup = ImageLookups.GetValue(source, CreateImageLookup);
        if (!lookup.IsMetadataComplete) {
            complete = false;
            return null;
        }
        (DataEntry Data, IWorkPackageEntry Entry, string MediaType,
            int PixelWidth, int PixelHeight)? resolved = null;
        foreach (ulong dataIdentifier in dataIdentifiers) {
            if (lookup.DuplicateDataIdentifiers.Contains(dataIdentifier)
                || !lookup.DataEntries.TryGetValue(dataIdentifier, out DataEntry? candidateData)) continue;
            string candidatePath = "Data/" + candidateData.StoredFileName;
            if (!lookup.PackageEntries.TryGetValue(candidatePath,
                    out IWorkPackageEntry? candidateEntry)) continue;
            string candidateMediaType = MediaType(candidateEntry.Path, candidateEntry.Bytes);
            if (!IsEditableOwnerImageMediaType(candidateMediaType)) continue;
            (int? candidateWidth, int? candidateHeight) = IWorkImageInfo.Read(
                candidateEntry.Bytes, candidateMediaType,
                projectionBudget.RemainingDecodedImageBytes, out long candidateDecodedBytes);
            projectionBudget.AddDecodedImageBytes(candidateDecodedBytes);
            if (!candidateWidth.HasValue || !candidateHeight.HasValue) continue;
            resolved = (candidateData, candidateEntry, candidateMediaType,
                candidateWidth.Value, candidateHeight.Value);
            break;
        }
        if (!resolved.HasValue) {
            complete = false;
            return null;
        }
        (DataEntry data, IWorkPackageEntry entry, string mediaType,
            int pixelWidth, int pixelHeight) = resolved.Value;
        IWorkGeometry? geometry = ReadGeometry(drawable, out bool geometryComplete);
        if (!geometryComplete) complete = false;
        bool hasMask = message.HasBytes(5);
        if (hasMask || message.HasUnexpectedWireKind(5, IWorkWireKind.Bytes)) complete = false;
        string? hyperlink = drawable.GetString(4, out bool hyperlinkComplete);
        string? accessibilityDescription = drawable.GetString(8, out bool accessibilityComplete);
        if (!hyperlinkComplete || !accessibilityComplete) complete = false;
        projectionBudget.AddTextCharacters(data.PreferredFileName.Length);
        if (hyperlink != null) projectionBudget.AddTextCharacters(hyperlink.Length);
        if (accessibilityDescription != null) {
            projectionBudget.AddTextCharacters(accessibilityDescription.Length);
        }
        return new IWorkImageAsset(data.PreferredFileName, entry.Path, mediaType, entry.Bytes,
            pixelWidth, pixelHeight, geometry, hasMask,
            hyperlink, accessibilityDescription);
    }

    internal static bool IsEditableOwnerImageMediaType(string mediaType) =>
        mediaType is "image/png" or "image/jpeg";

    private static ImageLookup CreateImageLookup(IWorkSourceDocument source) {
        var entries = new Dictionary<string, IWorkPackageEntry>(StringComparer.Ordinal);
        foreach (IWorkPackageEntry entry in source.Entries) {
            if (!entries.ContainsKey(entry.Path)) entries.Add(entry.Path, entry);
        }
        IReadOnlyDictionary<ulong, DataEntry> dataEntries = ReadDataEntries(source,
            out ISet<ulong> duplicateIdentifiers, out bool metadataComplete);
        return new ImageLookup(dataEntries, entries, duplicateIdentifiers, metadataComplete);
    }

    private static IReadOnlyDictionary<ulong, DataEntry> ReadDataEntries(IWorkSourceDocument source,
        out ISet<ulong> duplicateIdentifiers, out bool metadataComplete) {
        IWorkArchiveRecord? metadata = source.Index.UniqueOfType(PackageMetadataArchive,
            out bool duplicateMetadata);
        var duplicates = new HashSet<ulong>();
        duplicateIdentifiers = duplicates;
        metadataComplete = metadata != null && !duplicateMetadata;
        if (metadata == null) return new Dictionary<ulong, DataEntry>();
        int metadataEntryCount = IWorkProtobuf.CountFields(metadata.Payload, 4,
            source.Options.MaximumProtobufFieldCount);
        if (metadataEntryCount > source.Options.MaximumProjectedImages) {
            throw new InvalidDataException(
                $"iWork image metadata exceeds the configured projection limit of {source.Options.MaximumProjectedImages} entries.");
        }
        var result = new Dictionary<ulong, DataEntry>();
        IReadOnlyList<IWorkWireMessage> messages = IWorkProtobuf.ParseRepeatedMessages(
            metadata.Payload, 4, source.Options.MaximumProjectedImages,
            source.Options, out bool malformedEntries);
        metadataComplete = !malformedEntries;
        foreach (IWorkWireMessage message in messages) {
            ulong? identifier = message.GetUnsigned(1);
            string? preferred = message.GetString(3, out bool preferredComplete);
            string? stored = message.GetString(4, out bool storedComplete) ?? preferred;
            if (message.FieldCount(1) != 1
                || message.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || !preferredComplete || !storedComplete) metadataComplete = false;
            if (identifier.HasValue && preferred != null && stored != null
                && IsSafeFileName(stored)) {
                if (!result.ContainsKey(identifier.Value)) {
                    result.Add(identifier.Value, new DataEntry(identifier.Value, preferred, stored));
                } else {
                    duplicates.Add(identifier.Value);
                }
            }
        }
        return result;
    }

    private static ulong? DataIdentifier(IWorkWireMessage image, int field, out bool complete) {
        complete = true;
        if (!image.HasField(field)) return null;
        IWorkWireMessage? identifier = IWorkObjectIndex.TryGetMessage(
            image, field, out bool malformedIdentifier);
        ulong? value = identifier?.GetUnsigned(1);
        if (malformedIdentifier || identifier == null
            || identifier.FieldCount(1) != 1
            || identifier.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || !value.HasValue) {
            complete = false;
            return null;
        }
        return value;
    }

    private static bool IsSafeFileName(string value) => value.Length > 0
        && value != "." && value != ".."
        && value.IndexOf('/') < 0 && value.IndexOf('\\') < 0;

    private static string MediaType(string path, byte[] bytes) {
        string extension = Path.GetExtension(path).ToLowerInvariant();
        if (extension == ".png") return "image/png";
        if (extension is ".jpg" or ".jpeg") return "image/jpeg";
        if (extension == ".svg") return "image/svg+xml";
        if (extension == ".pdf") return "application/pdf";
        if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50
            && bytes[2] == 0x4e && bytes[3] == 0x47) return "image/png";
        if (bytes.Length >= 3 && bytes[0] == 0xff && bytes[1] == 0xd8 && bytes[2] == 0xff) return "image/jpeg";
        return string.Empty;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool InvalidFloat(IWorkWireMessage message, int field) =>
        message.FieldCount(field) > 1
        || message.HasUnexpectedWireKind(field, IWorkWireKind.Fixed32)
        || message.HasField(field) && !message.GetFloat(field).HasValue;

    private sealed class DataEntry {
        internal DataEntry(ulong identifier, string preferredFileName, string storedFileName) {
            Identifier = identifier;
            PreferredFileName = preferredFileName;
            StoredFileName = storedFileName;
        }
        internal ulong Identifier { get; }
        internal string PreferredFileName { get; }
        internal string StoredFileName { get; }
    }

    private sealed class ImageLookup {
        internal ImageLookup(IReadOnlyDictionary<ulong, DataEntry> dataEntries,
            IReadOnlyDictionary<string, IWorkPackageEntry> packageEntries,
            ISet<ulong> duplicateDataIdentifiers, bool isMetadataComplete) {
            DataEntries = dataEntries;
            PackageEntries = packageEntries;
            DuplicateDataIdentifiers = duplicateDataIdentifiers;
            IsMetadataComplete = isMetadataComplete;
        }
        internal IReadOnlyDictionary<ulong, DataEntry> DataEntries { get; }
        internal IReadOnlyDictionary<string, IWorkPackageEntry> PackageEntries { get; }
        internal ISet<ulong> DuplicateDataIdentifiers { get; }
        internal bool IsMetadataComplete { get; }
    }
}
