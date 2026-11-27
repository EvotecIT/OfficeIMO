namespace OfficeIMO.Provenance;

/// <summary>
/// AI generator fields in a TIFF structure: a TIFF file itself, or the EXIF block of a JPEG, WebP or PNG image. Reports the
/// primary-IFD <c>Software</c> tag when it names an AI product and the Exif-IFD <c>UserComment</c> when it holds AI generation
/// settings. Removal deletes the entry from its IFD in place and clears its value bytes, so every other offset in the
/// structure stays valid and the containing segment or chunk keeps its size.
/// </summary>
internal static partial class OfficeProvenanceTiff {
    private const ushort AsciiType = 2;
    private const ushort SoftwareTag = 305;
    private const ushort ExifIfdTag = 34665;
    private const ushort UserCommentTag = 37510;

    internal static void InspectGeneratorFields(
        byte[] tiff,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext context,
        string location,
        bool containerIsStructurallyValid) {
        if (!TryFindGeneratorFields(tiff, options, out List<TiffIfd> ifds, out List<GeneratorField> fields, out bool complete)) return;
        bool storageValid = complete && HasValidEntryStorage(tiff, ifds);
        foreach (GeneratorField field in fields) {
            context.Add(new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.AiGeneratorMetadata,
                $"{location}/{field.Name}",
                containerIsStructurallyValid && storageValid && field.Ifd.TagsAreSorted,
                field.Length,
                field.Value,
                OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia));
        }
    }

    /// <summary>Removes AI generator fields; <paramref name="output"/> has the same length as <paramref name="tiff"/>.</summary>
    internal static bool TryRemoveGeneratorFields(
        byte[] tiff,
        OfficeProvenanceRemovalOptions options,
        string location,
        List<OfficeProvenanceChange> changes,
        bool containerIsStructurallyValid,
        out byte[] output) {
        output = tiff;
        if (!options.RemoveAiSourceMetadata ||
            !TryFindGeneratorFields(tiff, options.Limits, out List<TiffIfd> ifds, out List<GeneratorField> fields, out bool complete)) return false;
        bool storageValid = complete && HasValidEntryStorage(tiff, ifds);
        byte[]? rewritten = null;
        foreach (IGrouping<TiffIfd, GeneratorField> group in fields.GroupBy(static field => field.Ifd)) {
            TiffIfd ifd = group.Key;
            bool valid = containerIsStructurallyValid && storageValid && ifd.TagsAreSorted;
            if (!valid && options.RequireStructurallyValidCarrier) continue;
            if (HasOverlappingIfdRewriteStorage(tiff, ifds, ifd, options.Limits.MaxContainerEntries)) continue;
            rewritten ??= OfficeProvenanceBinary.CloneForOutput(tiff, options.EffectiveMaxOutputBytes);
            var removed = new HashSet<TiffEntry>();
            foreach (GeneratorField field in group) {
                // Out-of-line text shared with another entry or image data cannot be cleared, so strict removal keeps that entry;
                // inline values go with the entry.
                bool shared = field.Length > field.Entry.InlineSize &&
                    HasOverlappingValueStorage(tiff, ifds, field.Entry, field.Offset, field.Length, options.Limits.MaxContainerEntries);
                if (shared && options.RequireStructurallyValidCarrier) continue;
                if (field.Length > field.Entry.InlineSize && !shared) Array.Clear(rewritten, field.Offset, field.Length);
                removed.Add(field.Entry);
                changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.AiGeneratorMetadata, $"{location}/{field.Name}", field.Length));
            }
            if (removed.Count == 0) continue;
            RewriteIfd(rewritten, ifd, ifd.Entries.Where(entry => !removed.Contains(entry)).ToList());
        }
        if (rewritten == null) return false;
        output = rewritten;
        return true;
    }

    /// <summary>Whether the structure holds AI generator fields, so a container knows the carrier is there before validating it.</summary>
    internal static bool HasGeneratorFields(byte[] tiff, OfficeProvenanceOptions options) =>
        TryFindGeneratorFields(tiff, options, out _, out _, out _);

    /// <summary>
    /// Finds the fields. When a later IFD or a SubIFD link is malformed, the primary and Exif IFDs are still read on their own and
    /// <paramref name="complete"/> is false, so strict removal keeps the fields while they are still reported.
    /// </summary>
    private static bool TryFindGeneratorFields(
        byte[] tiff,
        OfficeProvenanceOptions options,
        out List<TiffIfd> ifds,
        out List<GeneratorField> fields,
        out bool complete) {
        fields = new List<GeneratorField>();
        complete = true;
        try {
            ifds = ReadIfds(tiff, options, followExifIfd: true);
        } catch (InvalidDataException exception) when (!OfficeProvenanceLimitException.Is(exception)) {
            complete = false;
            try {
                ifds = ReadIfds(tiff, options, followExifIfd: true, followLinkedIfds: false);
            } catch (InvalidDataException inner) when (!OfficeProvenanceLimitException.Is(inner)) {
                ifds = new List<TiffIfd>();
                return false;
            }
        }
        if (ifds.Count == 0) return false;
        TiffIfd primary = ifds[0];
        TiffEntry? exifPointer = primary.Entries.FirstOrDefault(static entry => entry.Tag == ExifIfdTag);
        TiffIfd? exif = exifPointer != null && TryReadExifIfdPointer(tiff, exifPointer, primary.BigTiff, out ulong exifOffset)
            ? ifds.FirstOrDefault(ifd => (ulong)ifd.Offset == exifOffset)
            : null;
        // Every decoded value draws from the expanded-byte budget first, so entries that share one large value cannot multiply the work.
        long decodedBytes = 0;
        void Charge(int length) {
            decodedBytes += length;
            if (decodedBytes > options.MaxExpandedContainerBytes) {
                throw OfficeProvenanceLimitException.Create("TIFF generator fields exceed the configured expanded-byte limit.");
            }
        }
        foreach (TiffEntry entry in primary.Entries) {
            if (entry.Tag != SoftwareTag || entry.Type != AsciiType || !TryGetValueStorageRange(tiff, entry, out int offset, out int length)) continue;
            Charge(length);
            string value = System.Text.Encoding.UTF8.GetString(tiff, offset, length).TrimEnd('\0');
            if (OfficeProvenanceAiGenerators.TryMatch(value, out string product)) {
                fields.Add(new GeneratorField(primary, entry, "IFD0/Software", product, offset, length));
            }
        }
        if (exif != null) {
            foreach (TiffEntry entry in exif.Entries) {
                if (entry.Tag != UserCommentTag || entry.Type != UndefinedType || !TryGetValueStorageRange(tiff, entry, out int offset, out int length)) continue;
                Charge(length);
                if (OfficeProvenanceAiGenerators.IsAiGenerationRecord(OfficeProvenanceAiGenerators.DecodeUserComment(tiff, offset, length))) {
                    fields.Add(new GeneratorField(exif, entry, "ExifIFD/UserComment", "UserComment", offset, length));
                }
            }
        }
        return fields.Count != 0;
    }

    private sealed class GeneratorField {
        internal GeneratorField(TiffIfd ifd, TiffEntry entry, string name, string value, int offset, int length) {
            Ifd = ifd; Entry = entry; Name = name; Value = value; Offset = offset; Length = length;
        }
        internal TiffIfd Ifd { get; }
        internal TiffEntry Entry { get; }
        internal string Name { get; }
        internal string Value { get; }
        internal int Offset { get; }
        internal int Length { get; }
    }
}
