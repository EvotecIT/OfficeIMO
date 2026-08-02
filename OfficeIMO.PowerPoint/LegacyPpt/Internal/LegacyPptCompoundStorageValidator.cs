using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Validates embedded compound storages while retaining bounded VBA
    /// streams needed to verify the declared module graph.
    /// </summary>
    internal static class LegacyPptCompoundStorageValidator {
        internal static bool TryRead(byte[] bytes,
            LegacyPptImportOptions options, out OfficeCompoundFile? compound,
            out string? reason) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (options == null) throw new ArgumentNullException(nameof(options));
            int maxDirectoryEntries = Math.Max(1,
                Math.Min(65536, options.MaxRecordCount));
            int maxStreamCount = Math.Max(1,
                Math.Min(32768, options.MaxRecordCount));
            long physicalBytes = Math.Max(1L, bytes.LongLength);
            var readOptions = new OfficeCompoundReadOptions(
                maxDirectoryEntries, maxStreamCount,
                physicalBytes, physicalBytes);
            using var source = new MemoryStream(bytes, writable: false);
            return OfficeCompoundFileReader.TryReadSelective(source,
                readOptions, externalize: (name, _) =>
                    !name.StartsWith("VBA/",
                        StringComparison.OrdinalIgnoreCase),
                openExternalDestination: (_, _) => Stream.Null,
                out compound, out reason);
        }
    }
}
