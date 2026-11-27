using OfficeIMO.Archives;

namespace OfficeIMO.Reader.Archives;

internal static partial class ArchiveReaderAdapter {
    /// <summary>Reads a 7z archive. Members are listed first so oversized or unwanted members are never decoded; the rest
    /// are decoded together, once per solid folder, and verified by OfficeIMO.Archives.</summary>
    private static IEnumerable<ReaderChunk> ReadSevenZip(Stream source, string archivePath, int nestedDepth, ArchiveOperationBudget budget,
        ReadContext context) {
        IReadOnlyList<SevenZipArchiveEntry> listing;
        var admitted = new HashSet<string>(StringComparer.Ordinal);
        Dictionary<string, byte[]> data;
        try {
            source.Position = 0;
            listing = ArchiveReader.ListSevenZip(source, ReadOptions(context, new ArchiveOperationBudget(), null), context.Token);
            foreach (SevenZipArchiveEntry entry in listing) {
                if (entry.Kind == ArchiveEntryKind.File && IsWanted(entry.RelativePath, context.Archive)
                    && Admit(entry.RelativePath, entry.Length, nestedDepth, context)) admitted.Add(entry.RelativePath);
            }
            source.Position = 0;
            data = admitted.Count == 0 ? new Dictionary<string, byte[]>()
                : ArchiveReader.ReadSevenZip(source, ReadOptions(context, budget, admitted.Contains), context.Token)
                    .ToDictionary(entry => entry.RelativePath, entry => entry.Data, StringComparer.Ordinal);
        } catch (OfficeIMO.OfficeInspectionResourceLimitException error) {
            throw new ReaderResourceLimitException(error.LimitName, error.Maximum, error);
        }

        foreach (SevenZipArchiveEntry entry in listing.Where(entry => entry.Kind == ArchiveEntryKind.File && IsWanted(entry.RelativePath, context.Archive))
                     .OrderBy(entry => entry.RelativePath, StringComparer.Ordinal)) {
            context.Token.ThrowIfCancellationRequested();
            string path = entry.RelativePath;
            foreach (var chunk in EmitMember(path, archivePath, entry.Length, entry.LastWriteTime,
                         () => Admitted(data.TryGetValue(path, out byte[]? bytes) ? bytes : null, path), nestedDepth, budget, context))
                yield return chunk;
        }
    }

    /// <summary>Reads a Microsoft Cabinet or Windows Installer package through <see cref="ArchiveInspector"/>, which
    /// verifies every member it returns. Members the inventory cannot read become warnings.</summary>
    private static IEnumerable<ReaderChunk> ReadInspected(Stream source, string archivePath, int nestedDepth, ArchiveOperationBudget budget,
        ReadContext context) {
        source.Position = 0;
        ArchiveInspection inspection;
        try {
            inspection = ArchiveInspector.Open(source, InspectionOptions(context.Traversal, budget), context.Token);
        } catch (OfficeIMO.OfficeInspectionResourceLimitException error) {
            throw new ReaderResourceLimitException(error.LimitName, error.Maximum, error);
        }
        using (inspection) {
            var names = inspection.Entries.Where(entry => !entry.IsDirectory).Select(entry => MemberName(entry)).ToList();
            var duplicates = new HashSet<string>(names.GroupBy(name => name, StringComparer.Ordinal).Where(group => group.Count() > 1)
                .Select(group => group.Key), StringComparer.Ordinal);
            foreach (ArchiveInspectionEntry entry in inspection.Entries) {
                context.Token.ThrowIfCancellationRequested();
                if (entry.IsDirectory) continue;
                string name = MemberName(entry);
                string entryArchivePath = duplicates.Contains(name) ? DuplicatePath(archivePath, entry.Ordinal) : archivePath;
                if (entry.Diagnostic != null) {
                    if (IsWanted(name, context.Archive)) {
                        yield return BuildWarningChunk(context.Source, entryArchivePath, name, context.Warnings.Next(),
                            "Skipped archive entry: " + entry.Diagnostic.Message, context.ReaderOptions.ComputeHashes,
                            sourceLengthBytes: entry.Length, sourceLastWriteTime: entry.LastWriteTime);
                    }
                    continue;
                }
                foreach (var chunk in EmitMember(name, entryArchivePath, entry.Length, entry.LastWriteTime,
                             () => ReadAllBytes(inspection, entry, context.Token), nestedDepth, inspection.OperationBudget, context))
                    yield return chunk;
            }
        }
    }

    /// <summary>The installed target path for MSI files, otherwise the authored member name, with forward slashes.</summary>
    private static string MemberName(ArchiveInspectionEntry entry) => (entry.AuthoredPath ?? entry.RawName).Replace('\\', '/');
}
