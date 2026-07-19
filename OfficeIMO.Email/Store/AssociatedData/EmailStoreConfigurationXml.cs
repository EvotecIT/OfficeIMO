using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Email.Store;

internal static class EmailStoreConfigurationXml {
    internal static XDocument Parse(byte[] bytes, int maximumBytes, string description) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        if (bytes.Length > maximumBytes) {
            throw new EmailStoreLimitExceededException(nameof(maximumBytes), bytes.Length, maximumBytes);
        }
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = maximumBytes
        };
        try {
            using (var stream = new MemoryStream(bytes, writable: false))
            using (XmlReader reader = XmlReader.Create(stream, settings)) {
                return XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
            }
        } catch (XmlException exception) {
            throw new InvalidDataException(string.Concat(description, " is not valid bounded XML."), exception);
        }
    }

    internal static byte[] Serialize(XDocument document) {
        var settings = new XmlWriterSettings {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration == null,
            NewLineHandling = NewLineHandling.None
        };
        using (var stream = new MemoryStream()) {
            using (XmlWriter writer = XmlWriter.Create(stream, settings)) document.Save(writer);
            return stream.ToArray();
        }
    }

    internal static DateTimeOffset? ParseDate(string? value) {
        if (value == null) return null;
        return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out DateTimeOffset parsed)
            ? parsed
            : (DateTimeOffset?)null;
    }

    internal static string FormatDate(DateTimeOffset value) =>
        value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture);
}
