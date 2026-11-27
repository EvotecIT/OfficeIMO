using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Archives;
using OfficeIMO.Epub;

namespace OfficeIMO.Comic;

public static partial class ComicReader {
    private static XElement? ReadMetadata(IReadOnlyList<ArchiveEntryData> entries, string format, List<OfficeDiagnostic> diagnostics,
        CancellationToken token) {
        ArchiveEntryData? entry = entries.FirstOrDefault(item => IsComicInfo(item.RelativePath));
        if (entry == null) return null;
        const int maximumMetadataBytes = 256 * 1024;
        if (entry.Length > maximumMetadataBytes) {
            diagnostics.Add(Loss(format + "_COMICINFO_LIMIT", "ComicInfo exceeds the metadata limit.", entry.RelativePath));
            return null;
        }

        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = maximumMetadataBytes
            };
            using var stream = new MemoryStream(entry.Data, writable: false);
            using (XmlReader scan = XmlReader.Create(stream, settings)) {
                while (scan.Read()) {
                    token.ThrowIfCancellationRequested();
                    if (scan.Depth > 16) throw new XmlException("ComicInfo nesting exceeds 16 levels.");
                }
            }
            stream.Position = 0;
            using XmlReader reader = XmlReader.Create(stream, settings);
            XElement root = XElement.Load(reader);
            if (root.Name != "ComicInfo") throw new XmlException("ComicInfo root element is missing.");
            return root;
        } catch (XmlException exception) {
            diagnostics.Add(Loss(format + "_COMICINFO_INVALID", exception.Message, entry.RelativePath));
            return null;
        }
    }

    private static void ApplyMetadata(EpubPublication book, XElement? metadata, string format, List<OfficeDiagnostic> diagnostics) {
        foreach (XElement element in metadata?.Elements() ?? Enumerable.Empty<XElement>()) {
            string key = element.Name.LocalName;
            string value = element.Value.Trim();
            if (value.Length == 0) continue;
            string? field = key switch {
                "Writer" => "creator",
                "Penciller" or "Inker" or "Colorist" or "Letterer" or "CoverArtist" or "Editor" or "Translator" => "contributor",
                "Summary" => "description", "Publisher" => "publisher", "Year" => "date",
                "Genre" or "Tags" => "subject", "Web" => "source",
                "GTIN" => "identifier",
                _ => null
            };
            if (field != null) {
                book.AddDublinCoreMetadata(field, value);
            } else if (key == "Manga") {
                if (value == "YesAndRightToLeft") book.PageProgressionDirection = "rtl";
            } else if (key is not ("Title" or "LanguageISO" or "Series" or "Number" or "Pages" or "PageCount")) {
                diagnostics.Add(Loss(format + "_COMICINFO_FIELD_OMITTED", "The ComicInfo field has no mapped publication metadata.", key));
            }
        }
        if (Text(metadata, "Series") is { } series) {
            var collection = new EpubCollectionMetadata { Name = series };
            if (Text(metadata, "Number") is { } number) {
                string[] parts = number.Split('.');
                var position = new List<uint>();
                foreach (string part in parts) {
                    if (uint.TryParse(part, NumberStyles.None, CultureInfo.InvariantCulture, out uint value)) position.Add(value);
                    else break;
                }
                if (position.Count == parts.Length) collection.Position = position;
                else diagnostics.Add(Loss(format + "_COMICINFO_NUMBER_OMITTED",
                    "The issue number is outside the EPUB collection-position profile.", number));
            }
            book.AddCollection("comic-series", collection);
        }
    }

    private static Dictionary<int, (string? Bookmark, bool Cover)> ReadPages(XElement? metadata, int count, string format,
        List<OfficeDiagnostic> diagnostics) {
        var pages = new Dictionary<int, (string?, bool)>();
        foreach (XElement page in metadata?.Element("Pages")?.Elements("Page") ?? Enumerable.Empty<XElement>()) {
            if (!int.TryParse((string?)page.Attribute("Image"), out int index) || index < 0 || index >= count || pages.ContainsKey(index)) {
                diagnostics.Add(Loss(format + "_COMICINFO_PAGE_UNAVAILABLE", "The ComicInfo page index is invalid or duplicated.",
                    (string?)page.Attribute("Image")));
                continue;
            }
            string? bookmark = ((string?)page.Attribute("Bookmark"))?.Trim();
            string? type = (string?)page.Attribute("Type");
            pages.Add(index, (string.IsNullOrEmpty(bookmark) ? null : bookmark, type == "FrontCover"));
            if ((string?)page.Attribute("DoublePage") is "true" or "1" || type is not (null or "FrontCover" or "Story"))
                diagnostics.Add(Loss(format + "_COMICINFO_PAGE_ROLE_OMITTED",
                    "The page image is preserved, but its ComicInfo role or spread is not carried over.", (string?)page.Attribute("Image")));
        }
        return pages;
    }
}
