using System.Xml;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const string OpenDocumentManifestNamespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
    private const int MaximumOpenDocumentManifestBytes = 256 * 1024;

    /// <summary>
    /// Reports whether an OpenDocument package's <c>META-INF/manifest.xml</c> declares encrypted entries. Only that one
    /// entry is read, with its compressed and expanded sizes capped; a manifest that cannot be read leaves the answer false.
    /// </summary>
    private static bool HasEncryptedOpenDocumentEntries(Stream stream, long start, int maxEntries, CancellationToken cancellationToken) {
        long originalPosition = stream.Position;
        try {
            if (!TryLocateZipCentralDirectory(stream, start, out long centralDirectoryOffset, out int entryCount)) return false;
            stream.Position = start + centralDirectoryOffset;
            var header = new byte[ZipCentralDirectoryHeaderLength];
            for (int entryIndex = 0; entryIndex < Math.Min(entryCount, maxEntries); entryIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!ReadExact(stream, header, 0, header.Length) ||
                    ReadUInt32(header, 0) != ZipCentralDirectoryHeaderSignature) return false;
                ushort nameLength = ReadUInt16(header, 28);
                if (nameLength == 0 || nameLength > 4096 || nameLength > stream.Length - stream.Position) return false;
                var name = new byte[nameLength];
                if (!ReadExact(stream, name, 0, name.Length)) return false;
                long next = stream.Position + ReadUInt16(header, 30) + ReadUInt16(header, 32);
                if (next > stream.Length) return false;
                if (NormalizeZipEntryName(name) == "meta-inf/manifest.xml") {
                    byte[]? manifest = ReadBoundedZipEntry(stream, start, ReadUInt32(header, 42), ReadUInt16(header, 10),
                        ReadUInt32(header, 20), ReadUInt32(header, 24), cancellationToken);
                    return manifest != null && ManifestDeclaresEncryption(manifest);
                }
                stream.Position = next;
            }
            return false;
        } catch (InvalidDataException) {
            return false;
        } catch (IOException) {
            return false;
        } catch (XmlException) {
            return false;
        } finally {
            stream.Position = originalPosition;
        }
    }

    private static byte[]? ReadBoundedZipEntry(Stream stream, long start, uint localHeaderOffset, ushort compression,
        uint compressedSize, uint expandedSize, CancellationToken cancellationToken) {
        if (compression is not 0 and not 8 ||
            compressedSize is 0 or > MaximumOpenDocumentManifestBytes ||
            expandedSize is 0 or > MaximumOpenDocumentManifestBytes) return null;
        var local = new byte[30];
        long position = start + localHeaderOffset;
        if (position < start || position > stream.Length - local.Length) return null;
        stream.Position = position;
        if (!ReadExact(stream, local, 0, local.Length) || ReadUInt32(local, 0) != ZipLocalHeaderSignature ||
            (ReadUInt16(local, 6) & ZipEncryptedEntryFlag) != 0) return null;
        long dataPosition = stream.Position + ReadUInt16(local, 26) + ReadUInt16(local, 28);
        if (dataPosition > stream.Length - compressedSize) return null;
        stream.Position = dataPosition;
        var payload = new byte[(int)compressedSize];
        return ReadExact(stream, payload, 0, payload.Length)
            ? InflateBoundedContainerEntry(payload, compression, expandedSize, cancellationToken)
            : null;
    }

    private static bool ManifestDeclaresEncryption(byte[] manifest) {
        using var input = new MemoryStream(manifest, writable: false);
        using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaximumOpenDocumentManifestBytes
        });
        while (reader.Read()) {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "encryption-data" &&
                reader.NamespaceURI == OpenDocumentManifestNamespace) return true;
        }
        return false;
    }
}
