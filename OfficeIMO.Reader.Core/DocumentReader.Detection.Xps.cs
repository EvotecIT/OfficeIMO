using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    // Identify the OPC fixed-representation relationship, without loading pages/resources.
    // Both compressed and expanded metadata are capped independently of declarations.
    private static async Task<DetectionCandidate?> ReadXpsRelationshipCandidate(Stream stream, long start,
        uint offset, ushort compression, uint compressedSize, uint expandedSize, long returnPosition,
        bool asynchronous, CancellationToken token) {
        if (compression is not 0 and not 8 || compressedSize is 0 or > 65536 || expandedSize is 0 or > 65536) return null;
        try {
            token.ThrowIfCancellationRequested(); var header = new byte[30]; long position = start + offset;
            if (position < start || position > stream.Length - header.Length) return null;
            stream.Position = position;
            bool read = asynchronous ? await ReadExactAsync(stream, header, 0, header.Length, token).ConfigureAwait(false)
                : ReadExact(stream, header, 0, header.Length);
            if (!read || ReadUInt32(header, 0) != ZipLocalHeaderSignature || ReadUInt16(header, 8) != compression || (ReadUInt16(header, 6) & 1) != 0) return null;
            int nameLength = ReadUInt16(header, 26), extraLength = ReadUInt16(header, 28);
            if (nameLength is 0 or > 4096 || nameLength > stream.Length - stream.Position) return null;
            var name = new byte[nameLength];
            read = asynchronous ? await ReadExactAsync(stream, name, 0, name.Length, token).ConfigureAwait(false)
                : ReadExact(stream, name, 0, name.Length);
            if (!read || NormalizeZipEntryName(name) != "_rels/.rels" || extraLength > stream.Length - stream.Position - compressedSize) return null;
            stream.Position += extraLength; var payload = new byte[(int)compressedSize];
            read = asynchronous ? await ReadExactAsync(stream, payload, 0, payload.Length, token).ConfigureAwait(false)
                : ReadExact(stream, payload, 0, payload.Length);
            if (!read) return null;
            var xml = InflateBoundedContainerEntry(payload, compression, expandedSize, token);
            if (xml == null) return null;
            using var input = new MemoryStream(xml, false);
            using var reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null, MaxCharactersInDocument = 65536 });
            var root = XElement.Load(reader); XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";
            if (root.Name != ns + "Relationships") return null;
            foreach (var relation in root.Elements(ns + "Relationship")) {
                token.ThrowIfCancellationRequested(); string? mode = (string?)relation.Attribute("TargetMode");
                if (mode != null && mode != "Internal" || string.IsNullOrWhiteSpace((string?)relation.Attribute("Target"))) continue;
                string? type = (string?)relation.Attribute("Type");
                if (type == "http://schemas.microsoft.com/xps/2005/06/fixedrepresentation")
                    return DetectionCandidate.High(ReaderInputKind.Xps, "application/vnd.ms-xpsdocument", "container:xps-fixed-representation");
                if (type == "http://schemas.openxps.org/oxps/v1.0/fixedrepresentation")
                    return DetectionCandidate.High(ReaderInputKind.Xps, "application/oxps", "container:openxps-fixed-representation");
            }
            return null;
        } catch (InvalidDataException) { return null; }
        catch (IOException) { return null; }
        catch (XmlException) { return null; }
        finally { stream.Position = returnPosition; }
    }
}
