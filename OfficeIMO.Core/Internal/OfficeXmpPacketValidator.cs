using System;
using System.IO;
using System.Text;
using System.Xml;

namespace OfficeIMO.Core.Internal;

/// <summary>Validates bounded UTF-8 XMP packets shared by image container readers.</summary>
internal static class OfficeXmpPacketValidator {
    internal const int MaximumPacketBytes = 1024 * 1024;
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    internal static bool TryValidate(byte[] data, int offset, int length) {
        if (data == null || offset < 0 || length <= 0 || length > MaximumPacketBytes ||
            offset > data.Length - length) {
            return false;
        }

        try {
            string xml = StrictUtf8.GetString(data, offset, length);
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaximumPacketBytes
            };
            using var input = new StringReader(xml);
            using XmlReader reader = XmlReader.Create(input, settings);
            bool foundXmpRoot = false;
            while (reader.Read()) {
                if (reader.NodeType != XmlNodeType.Element || reader.Depth != 0) continue;
                foundXmpRoot =
                    (reader.LocalName == "xmpmeta" && reader.NamespaceURI == "adobe:ns:meta/") ||
                    (reader.LocalName == "RDF" &&
                     reader.NamespaceURI == "http://www.w3.org/1999/02/22-rdf-syntax-ns#");
                if (!foundXmpRoot) return false;
            }
            return foundXmpRoot;
        } catch (Exception exception) when (exception is DecoderFallbackException || exception is XmlException) {
            return false;
        }
    }
}
