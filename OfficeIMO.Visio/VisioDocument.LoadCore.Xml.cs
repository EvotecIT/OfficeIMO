using System.IO;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {
        internal const long MaxPackageXmlPartBytes = 32_000_000;
        internal const long MaxPackageXmlCharacters = 30_000_000;

        private static readonly XmlReaderSettings PackageXmlReaderSettings = new() {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaxPackageXmlCharacters,
            MaxCharactersFromEntities = 0,
        };

        private static XDocument LoadPackageXml(PackagePart part, string description) {
            using Stream stream = part.GetStream();
            using Stream boundedStream = new BoundedReadStream(stream, MaxPackageXmlPartBytes, description);
            using XmlReader reader = XmlReader.Create(boundedStream, PackageXmlReaderSettings);
            return XDocument.Load(reader, LoadOptions.None);
        }
    }
}
