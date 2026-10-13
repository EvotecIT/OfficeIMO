using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.OneNote;

/// <summary>Adapts the shared Cabinet reader to OneNote entries and diagnostic codes.</summary>
internal static class OneNoteCabinetArchiveReader {
    internal static IReadOnlyList<OneNoteCabinetEntry> Read(Stream stream, long? maxInputBytes,
        long maxExpandedBytes, long maxEntryBytes, int maxEntries, CancellationToken cancellationToken = default) {
        try {
            return Adapt(OfficeCabinetArchiveReader.Read(stream, maxInputBytes, maxExpandedBytes, maxEntryBytes, maxEntries, cancellationToken));
        } catch (OfficeCabinetException exception) {
            throw Adapt(exception);
        }
    }

    internal static IReadOnlyList<OneNoteCabinetEntry> Read(byte[] data, long maxExpandedBytes,
        long maxEntryBytes, int maxEntries, CancellationToken cancellationToken = default) {
        try {
            return Adapt(OfficeCabinetArchiveReader.Read(data, maxExpandedBytes, maxEntryBytes, maxEntries, cancellationToken));
        } catch (OfficeCabinetException exception) {
            throw Adapt(exception);
        }
    }

    private static IReadOnlyList<OneNoteCabinetEntry> Adapt(IReadOnlyList<OfficeCabinetEntry> entries) =>
        entries.Select(entry => new OneNoteCabinetEntry(entry.Name, entry.Data)).ToList().AsReadOnly();

    private static OneNoteFormatException Adapt(OfficeCabinetException exception) =>
        new OneNoteFormatException("ONENOTE_" + exception.Code, exception.Message, exception.Offset, exception);
}

internal sealed class OneNoteCabinetEntry {
    internal OneNoteCabinetEntry(string name, byte[] data) {
        Name = name;
        Data = data;
    }

    internal string Name { get; }
    internal byte[] Data { get; }
}
