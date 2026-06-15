using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyRevisionTables(Dictionary<string, string> values) {
            var authors = new List<RtfRevisionAuthor>();
            for (int index = 0; ; index++) {
                string key = "author." + index.ToString(CultureInfo.InvariantCulture) + ".name";
                if (!values.TryGetValue(key, out string? name)) {
                    break;
                }

                authors.Add(new RtfRevisionAuthor(name));
            }

            if (authors.Count > 0) {
                _document.ReplaceRevisionAuthors(authors);
            }

            _document.SetRevisionRootSaveId(ReadInt(values, "rsid.root"));
            var saveIds = new List<int>();
            for (int index = 0; ; index++) {
                string key = "rsid." + index.ToString(CultureInfo.InvariantCulture);
                int? id = ReadInt(values, key);
                if (!id.HasValue) {
                    break;
                }

                saveIds.Add(id.Value);
            }

            if (saveIds.Count > 0) {
                _document.ReplaceRevisionSaveIds(saveIds);
            }
        }
    }
}
