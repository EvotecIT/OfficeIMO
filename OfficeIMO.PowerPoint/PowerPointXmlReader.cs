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

    internal static bool? ReadSlideShow(Stream stream) {
        using XmlReader reader = XmlReader.Create(stream, PackageXmlReaderSettings);
        reader.MoveToContent();
        string? value = reader.GetAttribute("show");
        return value == null ? null : XmlConvert.ToBoolean(value);
    }
}
