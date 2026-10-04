using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.PowerPoint;

internal static class PowerPointXmlReader {
    internal const long MaximumPackageXmlCharacters = 16L * 1024L * 1024L;

    private static readonly XmlReaderSettings PackageXmlReaderSettings = new() {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        MaxCharactersInDocument = MaximumPackageXmlCharacters
    };

    internal static XDocument LoadPackagePartXml(Stream stream, LoadOptions options = LoadOptions.None) {
        using XmlReader reader = XmlReader.Create(stream, PackageXmlReaderSettings);
        return XDocument.Load(reader, options);
    }

    internal static bool? ReadSlideShow(Stream stream, long maximumCharactersInPart) {
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = maximumCharactersInPart > 0
                ? Math.Min(MaximumPackageXmlCharacters, maximumCharactersInPart)
                : MaximumPackageXmlCharacters
        };
        using XmlReader reader = XmlReader.Create(stream, settings);
        reader.MoveToContent();
        string? value = reader.GetAttribute("show");
        while (reader.Read()) { }
        return value == null ? null : XmlConvert.ToBoolean(value);
    }
}
