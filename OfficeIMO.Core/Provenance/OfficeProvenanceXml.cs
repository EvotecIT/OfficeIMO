using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceXml {
    private const int MaximumDepth = 256;

    internal static bool TryLoadDocument(
        byte[] data,
        OfficeProvenanceOptions options,
        out XDocument? document) {
        document = null;
        if (data.LongLength > options.MaxAssetBytes) return false;
        try {
            ValidateMaterializedNodeBudget(data, options, "XMP");
            using var stream = new MemoryStream(data, writable: false);
            using XmlReader reader = XmlReader.Create(stream, CreateReaderSettings(options));
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            return document.Root != null;
        } catch (XmlException) {
            document = null;
            return false;
        }
    }

    internal static XmlReaderSettings CreateReaderSettings(OfficeProvenanceOptions options) =>
        new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = options.MaxAssetBytes,
            MaxCharactersFromEntities = 0,
            IgnoreWhitespace = false
        };

    internal static Encoding ResolveTextEncoding(
        string filePath,
        long maximumBytes,
        CancellationToken cancellationToken) {
        string fullPath = Path.GetFullPath(filePath);
        using var stream = new FileStream(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            4096,
            FileOptions.SequentialScan);
        if (stream.Length > maximumBytes) {
            throw new InvalidDataException("The XML document exceeds the configured text-integrity byte limit.");
        }
        cancellationToken.ThrowIfCancellationRequested();
        using var reader = new XmlTextReader(stream) {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null
        };
        reader.Read();
        cancellationToken.ThrowIfCancellationRequested();
        return reader.Encoding ?? Encoding.UTF8;
    }

    internal static void ValidateMaterializedNodeBudget(
        byte[] data,
        OfficeProvenanceOptions options,
        string formatName) {
        using var stream = new MemoryStream(data, writable: false);
        ValidateMaterializedNodeBudget(stream, options, formatName);
    }

    internal static void ValidateMaterializedNodeBudget(
        Stream stream,
        OfficeProvenanceOptions options,
        string formatName) {
        using XmlReader reader = XmlReader.Create(stream, CreateReaderSettings(options));
        int materializedNodes = 0;
        while (reader.Read()) {
            if (reader.Depth > MaximumDepth) {
                throw OfficeProvenanceLimitException.Create($"{formatName} exceeds the configured XML depth limit.");
            }
            int nodes = reader.NodeType == XmlNodeType.Element
                ? 1 + reader.AttributeCount
                : IsMaterializedNode(reader.NodeType) ? 1 : 0;
            if (nodes > 0 && materializedNodes > options.MaxContainerEntries - nodes) {
                throw OfficeProvenanceLimitException.Create($"{formatName} exceeds the configured XML node limit.");
            }
            materializedNodes += nodes;
        }
    }

    private static bool IsMaterializedNode(XmlNodeType type) =>
        type is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.ProcessingInstruction or
            XmlNodeType.Comment or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace;
}
