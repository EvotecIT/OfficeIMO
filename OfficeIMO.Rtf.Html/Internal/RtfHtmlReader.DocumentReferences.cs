using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyFileReferences(Dictionary<string, string> values) {
            var files = new List<RtfFileReference>();
            for (int index = 0; ; index++) {
                string prefix = "file." + index.ToString(CultureInfo.InvariantCulture);
                int? id = ReadInt(values, prefix + ".id");
                string? path = ReadString(values, prefix + ".path");
                if (!id.HasValue || string.IsNullOrWhiteSpace(path)) {
                    break;
                }

                var file = new RtfFileReference(id.Value, path!) {
                    RelativePathStart = ReadInt(values, prefix + ".relativePathStart"),
                    OperatingSystemNumber = ReadInt(values, prefix + ".operatingSystemNumber"),
                    Sources = ReadEnum(values, prefix + ".sources", RtfFileSource.None)
                };
                files.Add(file);
            }

            if (files.Count > 0) {
                _document.ReplaceFileReferences(files);
            }
        }

        private void ApplyXmlNamespaces(Dictionary<string, string> values) {
            var namespaces = new List<RtfXmlNamespace>();
            for (int index = 0; ; index++) {
                string prefix = "namespace." + index.ToString(CultureInfo.InvariantCulture);
                int? id = ReadInt(values, prefix + ".id");
                string? uri = ReadString(values, prefix + ".uri");
                if (!id.HasValue || string.IsNullOrWhiteSpace(uri)) {
                    break;
                }

                namespaces.Add(new RtfXmlNamespace(id.Value, uri!));
            }

            if (namespaces.Count > 0) {
                _document.ReplaceXmlNamespaces(namespaces);
            }
        }
    }
}
