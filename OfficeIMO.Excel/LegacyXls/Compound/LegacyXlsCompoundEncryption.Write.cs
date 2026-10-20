using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel.LegacyXls.Compound;

internal static partial class LegacyXlsCompoundEncryption {
    /// <summary>Rebuilds encrypted-only auxiliary entries and removes their clear directory counterparts.</summary>
    internal static byte[] ProtectSummaryEntries(OfficeCompoundFile source, IReadOnlyDictionary<string, bool> protectedEntries,
        OfficeBinaryRc4CryptoApiSession session, long maximumBytes) {
        var entries = new List<OfficeBinaryEncryptedSummary.Entry>();
        var removed = new List<string>();
        var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in protectedEntries) {
            byte[] bytes;
            if (item.Value) {
                if (!source.Streams.TryGetValue(item.Key, out bytes!))
                    throw new InvalidDataException("A protected XLS stream is missing: " + item.Key);
            } else {
                var storage = OfficeCompoundStorage.Extract(source, item.Key);
                bytes = OfficeCompoundFileWriter.Rewrite(storage, new Dictionary<string, byte[]>(), maxOutputBytes: maximumBytes);
            }
            entries.Add(new OfficeBinaryEncryptedSummary.Entry(item.Key, item.Value, bytes));
            if (item.Key.Equals("\u0005DocumentSummaryInformation", StringComparison.OrdinalIgnoreCase)) {
                replacements[item.Key] = OfficeOlePropertySetWriter.CreatePropertySet((OfficeOlePropertySetWriter.DocumentSummaryInformationFormatId,
                    OfficeOlePropertySetWriter.CreateSection(Array.Empty<OfficeOleProperty>())));
            } else removed.Add(item.Key);
        }
        replacements["encryption"] = OfficeBinaryEncryptedSummary.Encrypt(entries, session, maximumBytes);
        return OfficeCompoundFileWriter.Rewrite(source, replacements, removedPaths: removed,
            maxOutputBytes: maximumBytes);
    }
}
