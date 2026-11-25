using System.Text;
using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Import;

internal static class NotebookImportImage {
    private static readonly XmlWriterSettings SvgWriterSettings = new() {
        Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        OmitXmlDeclaration = true
    };

    internal static bool TryReadBase64DataUri(string? source, long maximumBytes, out string payload, out OfficeImageInfo info) {
        payload = string.Empty;
        if (!OfficeImageReader.TryReadBase64DataUri(source, maximumBytes, out byte[] bytes, out info)) return false;
        if (info.Format != OfficeImageFormat.Svg) {
            payload = Convert.ToBase64String(bytes);
            return true;
        }

        // Store the validated SVG root as Unicode XML without its byte encoding declaration.
        using var stream = new MemoryStream(bytes, writable: false);
        try {
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Ignore,
                XmlResolver = null
            });
            if (reader.MoveToContent() != XmlNodeType.Element) return false;
            using var output = new OfficeBoundedMemoryStream(maximumBytes);
            using (XmlWriter writer = XmlWriter.Create(output, SvgWriterSettings)) {
                writer.WriteNode(reader, defattr: false);
            }
            payload = Encoding.UTF8.GetString(output.GetBuffer(), 0, (int)output.Length);
            return true;
        } catch (Exception error) when (error is XmlException || OfficeOutputLimit.Is(error)) {
            return false;
        }
    }
}
