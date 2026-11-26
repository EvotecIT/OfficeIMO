using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Archives;
using OfficeIMO.Core.Internal;
using OfficeIMO.Epub;

namespace OfficeIMO.Fb2;

/// <summary>Reads FictionBook 2 documents and single-book FB2.ZIP archives.</summary>
public static partial class Fb2Reader {
    /// <summary>Input and XML payload ceiling for this reader.</summary>
    public const long MaximumInputBytes = 32L * 1024 * 1024;
    private static readonly XNamespace Fb = "http://www.gribuser.ru/xml/fictionbook/2.0";
    private static readonly XNamespace Link = "http://www.w3.org/1999/xlink";
    private const string Component = "OfficeIMO.Fb2";

    /// <summary>Projects the current stream into a writable EPUB publication and reports unsupported source content.</summary>
    public static OfficeConversionResult<EpubPublication> Import(Stream source, string sourceName = "book.fb2",
        long maximumInputBytes = MaximumInputBytes, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceName);
        if (maximumInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumInputBytes));
        maximumInputBytes = Math.Min(MaximumInputBytes, maximumInputBytes);
        byte[] payload;
        if (sourceName.EndsWith(".fb2.zip", StringComparison.OrdinalIgnoreCase)) {
            IReadOnlyList<ArchiveEntryData> entries = ArchiveReader.ReadZip(source, new ArchiveReadOptions {
                MaxInputBytes = maximumInputBytes,
                MaxEntryCount = 128,
                MaxExpandedBytes = MaximumInputBytes,
                MaxEntryBytes = MaximumInputBytes,
                MaxPathDepth = 16,
                IncludeFile = path => path.EndsWith(".fb2", StringComparison.OrdinalIgnoreCase)
            }, cancellationToken);
            if (entries.Count != 1) throw new InvalidDataException("An FB2.ZIP archive must contain exactly one FictionBook (.fb2) member.");
            payload = entries[0].Data;
        } else {
            payload = OfficeStreamReader.ReadForwardBytes(source, cancellationToken, maximumInputBytes);
        }

        XDocument document = ReadXml(payload, cancellationToken);
        XElement root = document.Root ?? throw new InvalidDataException("The FictionBook document has no root element.");
        if (root.Name != Fb + "FictionBook") throw new InvalidDataException("The document root is not FictionBook 2.");
        return new Projection(sourceName, cancellationToken).Create(root);
    }

    private static XDocument ReadXml(byte[] payload, CancellationToken token) {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = MaximumInputBytes
        };
        using var input = new MemoryStream(payload, writable: false);
        using (XmlReader scan = XmlReader.Create(input, settings)) {
            int nodes = 0;
            while (scan.Read()) {
                token.ThrowIfCancellationRequested();
                nodes += 1 + scan.AttributeCount;
                if (scan.Depth > 64 || nodes > 250_000)
                    throw new InvalidDataException("The FictionBook XML exceeds its depth or node limit.");
            }
        }
        input.Position = 0;
        using XmlReader reader = XmlReader.Create(input, settings);
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    private static string? Value(XElement? parent, string name) {
        XElement? element = parent?.Element(Fb + name);
        string? value = element == null ? null : name == "title"
            ? string.Join(" ", element.Elements(Fb + "p").Select(item => item.Value.Trim())).Trim()
            : element.Value.Trim();
        return string.IsNullOrEmpty(value) ? null : value;
    }

    private static string? DateValue(XElement? date) {
        string? value = ((string?)date?.Attribute("value"))?.Trim();
        if (string.IsNullOrEmpty(value)) value = date?.Value.Trim();
        return string.IsNullOrEmpty(value) ? null : value;
    }

    private static string Ordinal(int number) => number.ToString(CultureInfo.InvariantCulture);

    private sealed partial class Projection(string sourceName, CancellationToken token) {
        private readonly List<OfficeDiagnostic> _diagnostics = new();
        private readonly HashSet<string> _reported = new(StringComparer.Ordinal);

        internal OfficeConversionResult<EpubPublication> Create(XElement root) {
            XElement? description = root.Element(Fb + "description");
            XElement? titleInfo = description?.Element(Fb + "title-info");
            string title = Value(titleInfo, "book-title") ?? Path.GetFileNameWithoutExtension(sourceName);
            string language = Value(titleInfo, "lang") ?? "und";
            string? identifier = Value(description?.Element(Fb + "document-info"), "id");
            EpubPublication book;
            try { book = EpubImportedPublication.Create(title, language, identifier); }
            catch (ArgumentException) when (language != "und") {
                Report("FB2_LANGUAGE", "An invalid FictionBook language was replaced with an undetermined language.", language);
                book = EpubImportedPublication.Create(title, "und", identifier);
            }
            AddMetadata(book, description, titleInfo);
            AddImages(book, root);
            var navigation = new List<EpubNavigationEntry>();
            string? cover = (string?)titleInfo?.Element(Fb + "coverpage")?.Element(Fb + "image")?.Attribute(Link + "href");
            if (GetLocalFragment(cover) is { } coverId && _images.TryGetValue(coverId, out EpubManifestItem? image)) {
                book.SetCoverImage(image.Id);
                book.AddChapter("cover", "EPUB/cover.xhtml", "Cover",
                    new XElement("div", new XElement("img", new XAttribute("src", image.Href), new XAttribute("alt", title + " cover")))
                        .ToString(SaveOptions.DisableFormatting));
            } else if (cover != null) {
                Report("FB2_COVER_UNAVAILABLE", "The declared cover does not resolve to a supported embedded raster image.", cover);
            }
            if (root.Elements(Fb + "stylesheet").Any())
                Report("FB2_STYLE_OMITTED", "FictionBook stylesheets are omitted. Semantic formatting is retained.");
            XElement[] bodies = root.Elements(Fb + "body").Take(129).ToArray();
            if (bodies.Length is 0 or > 128) throw new InvalidDataException("A FictionBook document must contain between 1 and 128 body elements.");
            for (int i = 0; i < bodies.Length; i++) IndexAnchors(bodies[i], i);
            for (int i = 0; i < bodies.Length; i++) {
                token.ThrowIfCancellationRequested();
                XElement body = bodies[i];
                string path = "EPUB/" + BodyName(i);
                string label = Value(body, "title") ?? ((string?)body.Attribute("name") is { Length: > 0 } name ? name : title);
                var content = new XElement("div", body.Nodes().Select(node => Render(node, i, 1)));
                if (body.Attribute(XNamespace.Xml + "lang") is { } lang) content.Add(new XAttribute(lang));
                if (_elementIds.TryGetValue(body, out string? anchor)) content.SetAttributeValue("id", anchor);
                ReportDroppedAttributes(body, content);
                book.AddChapter("body-" + Ordinal(i + 1), path, label, content.ToString(SaveOptions.DisableFormatting));
                EpubNavigationEntry[] sections = ReadNavigation(body, path).ToArray();
                if (bodies.Length == 1 && sections.Length != 0) navigation.AddRange(sections);
                else navigation.Add(new EpubNavigationEntry(label, path, sections));
            }
            book.SetNavigation(navigation);
            return new OfficeConversionResult<EpubPublication>(book,
                new OfficeConversionReport(sourceName, "EPUB", diagnostics: _diagnostics));
        }

        private void AddMetadata(EpubPublication book, XElement? description, XElement? info) {
            XElement? print = description?.Element(Fb + "publish-info");
            foreach (XElement node in info?.Elements() ?? Enumerable.Empty<XElement>()) {
                token.ThrowIfCancellationRequested();
                string? text = node.Name == Fb + "date" ? DateValue(node) : node.Value.Trim();
                if (string.IsNullOrEmpty(text)) continue;
                switch (node.Name.LocalName) {
                    case "author":
                    case "translator":
                        string person = string.Join(" ", new[] { Value(node, "first-name"), Value(node, "middle-name"),
                            Value(node, "last-name") }.OfType<string>());
                        if (person.Length == 0) person = Value(node, "nickname") ?? text;
                        book.AddDublinCoreMetadata(node.Name.LocalName == "author" ? "creator" : "contributor", person);
                        break;
                    case "genre":
                    case "keywords": book.AddDublinCoreMetadata("subject", text); break;
                    case "annotation": book.AddDublinCoreMetadata("description", text); break;
                    case "date": book.AddDublinCoreMetadata("date", text); break;
                }
            }
            if (Value(print, "publisher") is { } publisher) book.AddDublinCoreMetadata("publisher", publisher);
            if (Value(print, "isbn") is { } isbn) book.AddDublinCoreMetadata("identifier", isbn);
            if (Value(print, "year") is { } year && DateValue(info?.Element(Fb + "date")) == null) book.AddDublinCoreMetadata("date", year);
            if (description?.Elements().Any(node => node.Name.LocalName is "custom-info" or "output" or "src-title-info") == true ||
                info?.Elements(Fb + "sequence").Any() == true)
                Report("FB2_METADATA_OMITTED", "Custom, original-title, distribution and series metadata are outside this projection.");
        }

        private void Report(string code, string message, string? location = null) {
            if (_diagnostics.Count >= 1_000 || !_reported.Add(code + ":" + (location ?? string.Empty))) return;
            _diagnostics.Add(new OfficeDiagnostic(code, OfficeDiagnosticSeverity.Warning, message, location,
                Component, OfficeConversionLossKind.Omission));
        }
    }
}
