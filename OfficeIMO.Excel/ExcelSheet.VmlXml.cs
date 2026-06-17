using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal const long MaxVmlXmlPartBytes = 12_000_000;
        internal const long MaxVmlXmlCharacters = 10_000_000;

        private static readonly XmlReaderSettings VmlXmlReaderSettings = new() {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaxVmlXmlCharacters,
            MaxCharactersFromEntities = 0,
        };

        /// <summary>
        /// Loads workbook VML XML parts with external entities disabled and bounded document size.
        /// </summary>
        private static XDocument LoadVmlXDocument(Stream stream) {
            if (stream.CanSeek && stream.Length > MaxVmlXmlPartBytes) {
                throw new InvalidDataException($"Excel VML XML part exceeds {MaxVmlXmlPartBytes} bytes.");
            }

            using XmlReader reader = XmlReader.Create(stream, VmlXmlReaderSettings);
            return XDocument.Load(reader, LoadOptions.None);
        }
    }
}
