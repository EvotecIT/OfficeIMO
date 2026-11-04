using System.IO.Compression;
using System.Xml;

namespace OfficeIMO.Excel.SpreadsheetXml;

/// <summary>Checks expanded package XML before the Open XML SDK can materialize a part.</summary>
internal static class SpreadsheetXmlSnapshotPreflight {
    internal static void Validate(byte[] snapshot, ExcelSpreadsheetXmlOptions options, long sourceMaxCharactersInPart) {
        long maximumExpandedBytes = options.MaxOutputBytes;
        long maximumPartCharacters = sourceMaxCharactersInPart > 0
            ? Math.Min(sourceMaxCharactersInPart, maximumExpandedBytes)
            : maximumExpandedBytes;
        long expandedBytes = 0;
        long elements = 0;
        using var source = new MemoryStream(snapshot, writable: false);
        using var archive = new ZipArchive(source, ZipArchiveMode.Read);
        foreach (ZipArchiveEntry entry in archive.Entries) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                && !entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
            if (entry.Length > maximumExpandedBytes - expandedBytes)
                throw new InvalidDataException("Workbook expanded XML exceeds MaxOutputBytes.");
            expandedBytes += entry.Length;
            using Stream part = entry.Open();
            using XmlReader reader = XmlReader.Create(part, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = maximumPartCharacters
            });
            try {
                while (reader.Read()) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (++elements > options.MaxXmlElements)
                        throw new InvalidDataException("Workbook expanded XML exceeds MaxXmlElements.");
                    if (reader.Depth + 1 > options.MaxXmlDepth)
                        throw new InvalidDataException("Workbook expanded XML exceeds MaxXmlDepth.");
                }
            } catch (XmlException exception) {
                throw new InvalidDataException("Workbook expanded XML failed admission limits or XML parsing.", exception);
            }
        }
    }
}
