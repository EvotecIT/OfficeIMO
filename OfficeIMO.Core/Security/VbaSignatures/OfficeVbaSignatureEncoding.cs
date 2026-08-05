using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Security;

internal static class OfficeVbaSignatureEncoding {
    internal const string AuthenticodeContentTypeOid = "1.3.6.1.4.1.311.2.1.4";
    internal const string Sha256Oid = "2.16.840.1.101.3.4.2.1";

    internal static byte[] CreateSignedContent(OfficeVbaSignatureProfile profile, byte[] contentHash) {
        if (profile != OfficeVbaSignatureProfile.Legacy) return CreateSpcIndirectDataContentV2(contentHash);
        byte[] data = Sequence(ObjectIdentifier("1.3.6.1.4.1.311.2.1.29"),
            Explicit(0, OctetString(new byte[] { 0x00 })));
        return Sequence(data, DigestInfo(contentHash));
    }

    internal static byte[] CreateDigSigInfoSerialized(byte[] cms, byte[] signerCertificate) {
        if (cms == null) throw new ArgumentNullException(nameof(cms));
        if (signerCertificate == null) throw new ArgumentNullException(nameof(signerCertificate));
        if (cms.Length == 0) throw new ArgumentException("CMS bytes are required.", nameof(cms));
        if (signerCertificate.Length == 0) {
            throw new ArgumentException("Signer certificate bytes are required.", nameof(signerCertificate));
        }
        const int infoHeader = 36;
        const int parentHeader = 8;
        int signatureOffset = parentHeader + infoHeader;
        byte[] certificateStore = CreateSerializedCertificateStore(signerCertificate);
        int certificateStoreOffset = checked(signatureOffset + cms.Length);
        int projectNameOffset = checked(certificateStoreOffset + certificateStore.Length);
        int timestampUrlOffset = checked(projectNameOffset + 2);
        var bytes = new List<byte>(checked(infoHeader + cms.Length + certificateStore.Length + 4));
        AppendUInt32(bytes, checked((uint)cms.Length));
        AppendUInt32(bytes, checked((uint)signatureOffset));
        AppendUInt32(bytes, checked((uint)certificateStore.Length));
        AppendUInt32(bytes, checked((uint)certificateStoreOffset));
        AppendUInt32(bytes, 0);
        AppendUInt32(bytes, checked((uint)projectNameOffset));
        AppendUInt32(bytes, 0);
        AppendUInt32(bytes, 0);
        AppendUInt32(bytes, checked((uint)timestampUrlOffset));
        bytes.AddRange(cms);
        bytes.AddRange(certificateStore);
        bytes.Add(0);
        bytes.Add(0);
        bytes.Add(0);
        bytes.Add(0);
        return bytes.ToArray();
    }

    private static byte[] CreateSerializedCertificateStore(byte[] certificate) {
        var bytes = new List<byte>(checked(32 + certificate.Length));
        AppendUInt32(bytes, 0);
        AppendUInt32(bytes, 0x54524543);
        AppendUInt32(bytes, 0x20);
        AppendUInt32(bytes, 1);
        AppendUInt32(bytes, checked((uint)certificate.Length));
        bytes.AddRange(certificate);
        AppendUInt32(bytes, 0);
        AppendUInt32(bytes, 0);
        AppendUInt32(bytes, 0);
        return bytes.ToArray();
    }

    private static byte[] CreateSpcIndirectDataContentV2(byte[] sourceHash) {
        var descriptorBytes = new List<byte>(12);
        AppendInt32(descriptorBytes, 12);
        AppendInt32(descriptorBytes, 1);
        AppendInt32(descriptorBytes, 1);
        byte[] descriptor = descriptorBytes.ToArray();
        byte[] data = Sequence(ObjectIdentifier("1.3.6.1.4.1.311.2.1.31"),
            Explicit(0, OctetString(descriptor)));

        byte[] algorithm = Encoding.ASCII.GetBytes(Sha256Oid + "\0");
        const int headerSize = 6 * 4;
        int sourceOffset = checked(headerSize + algorithm.Length);
        var serialized = new List<byte>(checked(sourceOffset + sourceHash.Length));
        AppendInt32(serialized, algorithm.Length);
        AppendInt32(serialized, 0);
        AppendInt32(serialized, sourceHash.Length);
        AppendInt32(serialized, headerSize);
        AppendInt32(serialized, 0);
        AppendInt32(serialized, sourceOffset);
        serialized.AddRange(algorithm);
        serialized.AddRange(sourceHash);
        return Sequence(data, DigestInfo(serialized.ToArray()));
    }

    internal static bool TryExtractV2SourceHash(byte[] serialized, string digestAlgorithmOid,
        out byte[] compiledHash, out byte[] sourceHash, out string detail) {
        const int headerSize = 6 * 4;
        compiledHash = Array.Empty<byte>();
        sourceHash = Array.Empty<byte>();
        detail = string.Empty;
        if (serialized == null || serialized.Length < headerSize) {
            detail = "The VBA V2 signature-data header is truncated.";
            return false;
        }
        int algorithmSize = ReadInt32LittleEndian(serialized, 0);
        int compiledHashSize = ReadInt32LittleEndian(serialized, 4);
        int sourceHashSize = ReadInt32LittleEndian(serialized, 8);
        int algorithmOffset = ReadInt32LittleEndian(serialized, 12);
        int compiledHashOffset = ReadInt32LittleEndian(serialized, 16);
        int sourceHashOffset = ReadInt32LittleEndian(serialized, 20);
        int algorithmEnd = algorithmOffset >= 0 && algorithmSize >= 0 &&
                           algorithmOffset <= serialized.Length - algorithmSize
            ? algorithmOffset + algorithmSize
            : -1;
        bool compiledLayoutValid = compiledHashSize == 0
            ? compiledHashOffset == 0 || compiledHashOffset == algorithmEnd
            : compiledHashSize > 0 && compiledHashSize <= 1024 &&
              compiledHashOffset == algorithmEnd &&
              compiledHashOffset <= serialized.Length - compiledHashSize &&
              sourceHashOffset == compiledHashOffset + compiledHashSize;
        if (algorithmSize <= 1 || algorithmSize > 128 || compiledHashSize < 0 ||
            sourceHashSize <= 0 || sourceHashSize > 1024 || algorithmOffset != headerSize ||
            algorithmEnd < 0 || !compiledLayoutValid || sourceHashOffset < algorithmEnd ||
            sourceHashOffset > serialized.Length - sourceHashSize ||
            sourceHashOffset + sourceHashSize != serialized.Length) {
            detail = "The VBA V2 signature-data offsets or lengths are invalid " +
                     "(algorithmSize=" + algorithmSize + ", compiledHashSize=" + compiledHashSize +
                     ", sourceHashSize=" + sourceHashSize + ", algorithmOffset=" + algorithmOffset +
                     ", compiledHashOffset=" + compiledHashOffset + ", sourceHashOffset=" +
                     sourceHashOffset + ", totalLength=" + serialized.Length + ").";
            return false;
        }
        if (serialized[algorithmOffset + algorithmSize - 1] != 0) {
            detail = "The VBA V2 signature-data algorithm identifier is not null-terminated.";
            return false;
        }
        string algorithm = Encoding.ASCII.GetString(serialized, algorithmOffset, algorithmSize - 1);
        if (!string.Equals(algorithm, digestAlgorithmOid, StringComparison.Ordinal)) {
            detail = "The VBA V2 source-hash algorithm does not match DigestInfo.";
            return false;
        }
        if (compiledHashSize > 0) {
            compiledHash = new byte[compiledHashSize];
            Buffer.BlockCopy(serialized, compiledHashOffset, compiledHash, 0, compiledHash.Length);
        }
        sourceHash = new byte[sourceHashSize];
        Buffer.BlockCopy(serialized, sourceHashOffset, sourceHash, 0, sourceHash.Length);
        return true;
    }

    private static byte[] DigestInfo(byte[] digest) => Sequence(
        Sequence(ObjectIdentifier(Sha256Oid), Null()), OctetString(digest));

    private static byte[] Sequence(params byte[][] values) => Tlv(0x30, Concat(values));
    private static byte[] Explicit(int tag, byte[] value) => Tlv((byte)(0xA0 + tag), value);
    private static byte[] OctetString(byte[] value) => Tlv(0x04, value);
    private static byte[] Null() => new byte[] { 0x05, 0x00 };

    private static byte[] Integer(int value) {
        if (value == 0) return new byte[] { 0x02, 0x01, 0x00 };
        var bytes = new List<byte>();
        uint remaining = checked((uint)value);
        while (remaining != 0) {
            bytes.Add((byte)remaining);
            remaining >>= 8;
        }
        bytes.Reverse();
        if ((bytes[0] & 0x80) != 0) bytes.Insert(0, 0);
        return Tlv(0x02, bytes.ToArray());
    }

    private static byte[] ObjectIdentifier(string oid) {
        string[] parts = oid.Split('.');
        if (parts.Length < 2) throw new ArgumentException("An ASN.1 object identifier needs at least two arcs.", nameof(oid));
        ulong first = ulong.Parse(parts[0], System.Globalization.CultureInfo.InvariantCulture);
        ulong second = ulong.Parse(parts[1], System.Globalization.CultureInfo.InvariantCulture);
        var bytes = new List<byte>();
        AppendBase128(bytes, checked(first * 40 + second));
        for (int index = 2; index < parts.Length; index++) {
            AppendBase128(bytes, ulong.Parse(parts[index], System.Globalization.CultureInfo.InvariantCulture));
        }
        return Tlv(0x06, bytes.ToArray());
    }

    private static void AppendBase128(ICollection<byte> output, ulong value) {
        var encoded = new List<byte> { (byte)(value & 0x7F) };
        while ((value >>= 7) != 0) encoded.Add((byte)((value & 0x7F) | 0x80));
        encoded.Reverse();
        foreach (byte item in encoded) output.Add(item);
    }

    private static byte[] Tlv(byte tag, byte[] value) => Concat(new[] { tag }, EncodeLength(value.Length), value);

    private static byte[] EncodeLength(int length) {
        if (length < 0x80) return new[] { (byte)length };
        var bytes = new List<byte>();
        int remaining = length;
        while (remaining != 0) {
            bytes.Add((byte)remaining);
            remaining >>= 8;
        }
        bytes.Reverse();
        bytes.Insert(0, (byte)(0x80 | bytes.Count));
        return bytes.ToArray();
    }

    private static byte[] Concat(params byte[][] values) {
        int length = values.Aggregate(0, (current, value) => checked(current + value.Length));
        var output = new byte[length];
        int offset = 0;
        foreach (byte[] value in values) {
            Buffer.BlockCopy(value, 0, output, offset, value.Length);
            offset += value.Length;
        }
        return output;
    }

    private static void AppendUInt32(ICollection<byte> output, uint value) {
        output.Add((byte)value);
        output.Add((byte)(value >> 8));
        output.Add((byte)(value >> 16));
        output.Add((byte)(value >> 24));
    }

    private static void AppendInt32(ICollection<byte> output, int value) =>
        AppendUInt32(output, unchecked((uint)value));

    private static int ReadInt32LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24;

}

internal static class OfficeVbaPackageSignatureWriter {
    private static readonly XNamespace ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
#if !NET8_0_OR_GREATER
    private static readonly OfficeVbaSignatureProfile[] SignatureProfiles = {
        OfficeVbaSignatureProfile.Legacy,
        OfficeVbaSignatureProfile.Agile,
        OfficeVbaSignatureProfile.V3
    };
#endif

    internal static void Write(string packagePath, string vbaPartUri,
        IReadOnlyDictionary<OfficeVbaSignatureProfile, byte[]> profileParts) {
        using ZipArchive archive = ZipFile.Open(packagePath, ZipArchiveMode.Update);
        string vbaEntryPath = Normalize(vbaPartUri);
        int slash = vbaEntryPath.LastIndexOf('/');
        string directory = slash < 0 ? string.Empty : vbaEntryPath.Substring(0, slash + 1);
        string relationshipPath = directory + "_rels/" + vbaEntryPath.Substring(slash + 1) + ".rels";

        XDocument relationships = ReadXml(archive, relationshipPath) ??
            new XDocument(new XElement(RelationshipsNamespace + "Relationships"));
        XElement root = relationships.Root ?? throw new InvalidDataException("The VBA relationship document has no root element.");
        if (root.Name != RelationshipsNamespace + "Relationships") {
            throw new InvalidDataException("The VBA relationship document has an invalid root element.");
        }
        foreach (XElement relationship in root.Elements(RelationshipsNamespace + "Relationship")
                     .Where(element => TryGetProfile((string?)element.Attribute("Type"))).ToArray()) {
            relationship.Remove();
        }

        XDocument contentTypes = ReadXml(archive, "[Content_Types].xml")
            ?? throw new InvalidDataException("The package has no [Content_Types].xml part.");
        XElement typesRoot = contentTypes.Root ?? throw new InvalidDataException("The package content-types document has no root element.");
        if (typesRoot.Name != ContentTypesNamespace + "Types") {
            throw new InvalidDataException("The package content-types document has an invalid root element.");
        }

        foreach (OfficeVbaSignatureProfile profile in EnumerateSignatureProfiles()) {
            string fileName = GetFileName(profile);
            string entryPath = directory + fileName;
            DeleteEntry(archive, entryPath);
            foreach (XElement existing in typesRoot.Elements(ContentTypesNamespace + "Override")
                         .Where(element => string.Equals((string?)element.Attribute("PartName"), "/" + entryPath,
                             StringComparison.OrdinalIgnoreCase)).ToArray()) existing.Remove();
            if (!profileParts.TryGetValue(profile, out byte[]? bytes)) continue;
            WriteEntry(archive, entryPath, bytes);
            typesRoot.Add(new XElement(ContentTypesNamespace + "Override",
                new XAttribute("PartName", "/" + entryPath),
                new XAttribute("ContentType", GetContentType(profile))));
            root.Add(new XElement(RelationshipsNamespace + "Relationship",
                new XAttribute("Id", UniqueRelationshipId(root, "rIdOfficeImoVba" + profile)),
                new XAttribute("Type", GetRelationshipType(profile)),
                new XAttribute("Target", fileName)));
        }
        WriteXml(archive, relationshipPath, relationships);
        WriteXml(archive, "[Content_Types].xml", contentTypes);
    }

    private static IEnumerable<OfficeVbaSignatureProfile> EnumerateSignatureProfiles() {
#if NET8_0_OR_GREATER
        return Enum.GetValues<OfficeVbaSignatureProfile>();
#else
        return SignatureProfiles;
#endif
    }

    private static string UniqueRelationshipId(XElement root, string prefix) {
        var used = new HashSet<string>(root.Elements(RelationshipsNamespace + "Relationship")
            .Select(element => (string?)element.Attribute("Id"))
            .Where(value => !string.IsNullOrWhiteSpace(value))!, StringComparer.Ordinal);
        string candidate = prefix;
        int suffix = 1;
        while (used.Contains(candidate)) candidate = prefix + suffix++;
        return candidate;
    }

    private static bool TryGetProfile(string? relationshipType) =>
        relationshipType == "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"
        || relationshipType == "http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile"
        || relationshipType == "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile"
        || relationshipType == "http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3";

    private static string GetFileName(OfficeVbaSignatureProfile profile) => profile switch {
        OfficeVbaSignatureProfile.Legacy => "vbaProjectSignature.bin",
        OfficeVbaSignatureProfile.Agile => "vbaProjectSignatureAgile.bin",
        OfficeVbaSignatureProfile.V3 => "vbaProjectSignatureV3.bin",
        _ => throw new ArgumentOutOfRangeException(nameof(profile))
    };

    private static string GetRelationshipType(OfficeVbaSignatureProfile profile) => profile switch {
        OfficeVbaSignatureProfile.Legacy => "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature",
        OfficeVbaSignatureProfile.Agile => "http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile",
        OfficeVbaSignatureProfile.V3 => "http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3",
        _ => throw new ArgumentOutOfRangeException(nameof(profile))
    };

    private static string GetContentType(OfficeVbaSignatureProfile profile) => profile switch {
        OfficeVbaSignatureProfile.Legacy => "application/vnd.ms-office.vbaProjectSignature",
        OfficeVbaSignatureProfile.Agile => "application/vnd.ms-office.vbaProjectSignatureAgile",
        OfficeVbaSignatureProfile.V3 => "application/vnd.ms-office.vbaProjectSignatureV3",
        _ => throw new ArgumentOutOfRangeException(nameof(profile))
    };

    private static XDocument? ReadXml(ZipArchive archive, string path) {
        ZipArchiveEntry? entry = FindEntry(archive, path);
        if (entry == null) return null;
        using Stream input = entry.Open();
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
        using XmlReader reader = XmlReader.Create(input, settings);
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    private static void WriteXml(ZipArchive archive, string path, XDocument document) {
        DeleteEntry(archive, path);
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using Stream output = entry.Open();
        using XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
            Encoding = new UTF8Encoding(false), Indent = false, CloseOutput = false
        });
        document.Save(writer);
    }

    private static void WriteEntry(ZipArchive archive, string path, byte[] bytes) {
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using Stream output = entry.Open();
        output.Write(bytes, 0, bytes.Length);
    }

    private static void DeleteEntry(ZipArchive archive, string path) => FindEntry(archive, path)?.Delete();

    private static ZipArchiveEntry? FindEntry(ZipArchive archive, string path) =>
        archive.Entries.FirstOrDefault(entry => string.Equals(entry.FullName, Normalize(path), StringComparison.OrdinalIgnoreCase));

    private static string Normalize(string path) => path.Replace('\\', '/').TrimStart('/');
}
