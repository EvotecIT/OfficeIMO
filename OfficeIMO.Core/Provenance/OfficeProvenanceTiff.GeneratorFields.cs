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
        if (!TryFindGeneratorFields(tiff, options, out List<TiffIfd> ifds, out List<GeneratorField> fields)) return;
        bool storageValid = HasValidEntryStorage(tiff, ifds);
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
            !TryFindGeneratorFields(tiff, options.Limits, out List<TiffIfd> ifds, out List<GeneratorField> fields)) return false;
        bool storageValid = HasValidEntryStorage(tiff, ifds);
        byte[]? rewritten = null;
        foreach (IGrouping<TiffIfd, GeneratorField> group in fields.GroupBy(static field => field.Ifd)) {
            TiffIfd ifd = group.Key;
            bool valid = containerIsStructurallyValid && storageValid && ifd.TagsAreSorted;
            if (!valid && options.RequireStructurallyValidCarrier) continue;
            if (HasOverlappingIfdRewriteStorage(tiff, ifds, ifd, options.Limits.MaxContainerEntries)) continue;
            rewritten ??= OfficeProvenanceBinary.CloneForOutput(tiff, options.EffectiveMaxOutputBytes);
            var removed = new HashSet<TiffEntry>(group.Select(static field => field.Entry));
            foreach (GeneratorField field in group) {
                // Clear out-of-line text unless another entry or image data shares those bytes; inline values go with the entry.
                if (field.Length > field.Entry.InlineSize &&
                    !HasOverlappingValueStorage(tiff, ifds, field.Entry, field.Offset, field.Length, options.Limits.MaxContainerEntries)) {
                    Array.Clear(rewritten, field.Offset, field.Length);
                }
                changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.AiGeneratorMetadata, $"{location}/{field.Name}", field.Length));
            }
            RewriteIfd(rewritten, ifd, ifd.Entries.Where(entry => !removed.Contains(entry)).ToList());
        }
        if (rewritten == null) return false;
        output = rewritten;
        return true;
    }

    private static bool TryFindGeneratorFields(
        byte[] tiff,
        OfficeProvenanceOptions options,
        out List<TiffIfd> ifds,
        out List<GeneratorField> fields) {
        fields = new List<GeneratorField>();
        try {
            ifds = ReadIfds(tiff, options, followExifIfd: true);
        } catch (InvalidDataException exception) when (!OfficeProvenanceLimitException.Is(exception)) {
            ifds = new List<TiffIfd>();
            return false;
        }
        if (ifds.Count == 0) return false;
        TiffIfd primary = ifds[0];
        TiffEntry? exifPointer = primary.Entries.FirstOrDefault(static entry => entry.Tag == ExifIfdTag);
        TiffIfd? exif = exifPointer == null ? null : ifds.FirstOrDefault(ifd => (ulong)ifd.Offset == exifPointer.ValueOrOffset);
        foreach (TiffEntry entry in primary.Entries) {
            if (entry.Tag != SoftwareTag || entry.Type != AsciiType || !TryGetValueStorageRange(tiff, entry, out int offset, out int length)) continue;
            string value = System.Text.Encoding.UTF8.GetString(tiff, offset, length).TrimEnd('\0');
            if (OfficeProvenanceAiGenerators.IsAiGenerator(value)) fields.Add(new GeneratorField(primary, entry, "IFD0/Software", value, offset, length));
        }
        if (exif != null) {
            foreach (TiffEntry entry in exif.Entries) {
                if (entry.Tag != UserCommentTag || entry.Type != UndefinedType || !TryGetValueStorageRange(tiff, entry, out int offset, out int length)) continue;
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
