using System.Text;
using System.Threading;
using System.Xml;
using OfficeIMO.Provenance;

namespace OfficeIMO.IWork.Internal;

/// <summary>Packages owned archives and resources through Core's deterministic bounded ZIP writer.</summary>
internal static class IWorkNativePackageWriter {
    internal static byte[] Write(byte[] document, byte[] metadata, string identifier, int maximumBytes,
        CancellationToken cancellationToken, IEnumerable<(string FileName, byte[] Bytes)>? resources = null) {
        var entries = new SortedDictionary<string, byte[]>(StringComparer.Ordinal) {
            ["Index/Document.iwa"] = document,
            ["Index/Metadata.iwa"] = metadata,
            ["Metadata/DocumentIdentifier"] = Encoding.ASCII.GetBytes(identifier),
            ["Metadata/Properties.plist"] = Properties(identifier),
            ["Metadata/BuildVersionHistory.plist"] = Plist(writer => {
                writer.WriteStartElement("array"); writer.WriteElementString("string", "OfficeIMO.IWork"); writer.WriteEndElement();
            })
        };
        if (resources != null) foreach (var resource in resources) {
            cancellationToken.ThrowIfCancellationRequested();
            entries.Add("Data/" + resource.FileName, resource.Bytes);
        }
        return WriteEntries(entries, maximumBytes, cancellationToken);
    }

    internal static byte[] WriteEntries(IReadOnlyDictionary<string, byte[]> entries, int maximumBytes,
        CancellationToken cancellationToken) {
        // Core writes local DOS timestamps. Local midnight keeps bytes independent of the host time zone.
        var timestamp = new DateTimeOffset(new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Local));
        var zipEntries = entries.OrderBy(entry => entry.Key, StringComparer.Ordinal).Select(entry => new OfficeProvenanceZipWriteEntry(entry.Key, entry.Value.Length,
            false, timestamp, 0, 0, Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<byte>(),
            () => new MemoryStream(entry.Value, writable: false))).ToArray();
        byte[] output = OfficeProvenanceZipWriter.Write(zipEntries, maximumBytes,
            maximumOutputBytes: maximumBytes, cancellationToken: cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return output;
    }

    internal static void Copy(byte[] bytes, Stream destination, CancellationToken cancellationToken) {
        for (int offset = 0; offset < bytes.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(64 * 1024, bytes.Length - offset);
            destination.Write(bytes, offset, count);
            offset += count;
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static byte[] Properties(string identifier) => Plist(writer => {
        writer.WriteStartElement("dict");
        var values = new SortedDictionary<string, string>(StringComparer.Ordinal) {
            ["documentUUID"] = identifier, ["stableDocumentUUID"] = identifier, ["versionUUID"] = identifier,
            ["privateUUID"] = identifier, ["shareUUID"] = identifier, ["revision"] = "0::" + identifier,
            ["fileFormatVersion"] = "14.4.1"
        };
        foreach (var item in values) { writer.WriteElementString("key", item.Key); writer.WriteElementString("string", item.Value); }
        writer.WriteElementString("key", "isMultiPage"); writer.WriteStartElement("true"); writer.WriteEndElement();
        writer.WriteElementString("key", "hasExternalReferenceOrMissingOrUnmaterializedRemoteData"); writer.WriteStartElement("false"); writer.WriteEndElement();
        writer.WriteEndElement();
    });

    private static byte[] Plist(Action<XmlWriter> content) {
        using var output = new MemoryStream();
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false })) {
            writer.WriteStartDocument();
            writer.WriteDocType("plist", "-//Apple//DTD PLIST 1.0//EN", "http://www.apple.com/DTDs/PropertyList-1.0.dtd", null);
            writer.WriteStartElement("plist"); writer.WriteAttributeString("version", "1.0");
            content(writer);
            writer.WriteEndElement(); writer.WriteEndDocument();
        }
        return output.ToArray();
    }
}
