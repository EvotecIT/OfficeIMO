using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Text;
using System.Xml;
using System.Globalization;

namespace OfficeIMO.Excel.Utilities {
    internal static class ExcelXmlPartWriter {
        /// <summary>Writes text-bearing package parts without XML normalizing carriage returns.</summary>
        internal static void SavePreservingLineEndings(OpenXmlPart part, OpenXmlElement root) {
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            WritePreservingLineEndings(stream, root);
        }

        internal static void WritePreservingLineEndings(Stream stream, OpenXmlElement root) {
            using var writer = XmlWriter.Create(stream, new XmlWriterSettings {
                Encoding = new UTF8Encoding(false),
                CloseOutput = false,
                NewLineHandling = NewLineHandling.Entitize
            });
            writer.WriteStartDocument();
            root.WriteTo(writer);
        }

        internal static string SerializePreservingLineEndings(OpenXmlElement element) {
            var builder = new StringBuilder();
            using var textWriter = new StringWriter(builder, CultureInfo.InvariantCulture);
            using (var writer = XmlWriter.Create(textWriter, new XmlWriterSettings {
                ConformanceLevel = ConformanceLevel.Fragment,
                OmitXmlDeclaration = true,
                NewLineHandling = NewLineHandling.Entitize
            })) {
                element.WriteTo(writer);
            }
            return builder.ToString();
        }
    }
}
