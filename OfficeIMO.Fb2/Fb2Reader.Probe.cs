using System.Text;
using System.Xml;
using OfficeIMO.Archives;

namespace OfficeIMO.Fb2;

public static partial class Fb2Reader {
    internal static bool Probe(Stream source, string? name, long maximumInputBytes, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        try {
            if (name?.EndsWith(".fb2.zip", StringComparison.OrdinalIgnoreCase) == true) {
                using var archive = ArchiveInspector.Open(source, new ArchiveInspectionOptions {
                    MaxInputBytes = maximumInputBytes, MaxEntryCount = 128, MaxEntryBytes = MaximumInputBytes
                }, token);
                if (archive.Format != ArchiveFormat.Zip) return false;
                var books = archive.Entries.Where(entry => !entry.IsDirectory && entry.RawName.EndsWith(".fb2", StringComparison.OrdinalIgnoreCase))
                    .Take(2).ToArray();
                return books.Length == 1 && books[0].CanRead;
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using var xml = XmlReader.Create(source, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                MaxCharactersInDocument = Math.Min(maximumInputBytes, 64 * 1024)
            });
            return xml.MoveToContent() == XmlNodeType.Element &&
                xml.LocalName == "FictionBook" && xml.NamespaceURI == Fb.NamespaceName;
        } catch (Exception error) when (error is XmlException or InvalidDataException or NotSupportedException or IOException) {
            return false;
        }
    }
}
