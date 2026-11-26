using System.Text;
using System.Globalization;
using System.Net;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Epub;

namespace OfficeIMO.Mobi;

/// <summary>Reads unencrypted PalmDOC, Mobipocket and Kindle Format 8 ebooks.</summary>
public static partial class MobiReader {
    /// <summary>Maximum Palm database input size.</summary>
    public const long MaximumInputBytes = 64L * 1024 * 1024;
    private const string Component = "OfficeIMO.Mobi";

    /// <summary>Projects a Palm database stream into a writable EPUB publication with reported source losses.</summary>
    /// <exception cref="NotSupportedException">The book is encrypted or uses a profile outside the supported reader.</exception>
    public static OfficeConversionResult<EpubPublication> Import(Stream source, string sourceName = "book.mobi",
        long maximumInputBytes = MaximumInputBytes, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceName);
        if (maximumInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumInputBytes));
        byte[] payload = OfficeStreamReader.ReadForwardBytes(source, cancellationToken, Math.Min(MaximumInputBytes, maximumInputBytes));
        var database = new PalmDatabase(payload, cancellationToken);
        var first = new MobiHeader(database, 0, cancellationToken);
        MobiHeader selected = first;
        if (!database.IsPalmDoc && !first.IsKf8 && first.Integer(121) is { } boundary) {
            if (boundary <= first.TextRecords || !PalmDatabase.HasSignature(database.Record(boundary - 1), "BOUNDARY"))
                throw new InvalidDataException("The dual-format Kindle book has no valid KF8 boundary.");
            selected = new MobiHeader(database, boundary, cancellationToken);
            if (!selected.IsKf8) throw new InvalidDataException("The dual-format Kindle book does not contain its declared KF8 section.");
        }
        byte[] text = PalmDocText.Read(database, selected, cancellationToken);
        return new Projection(database, first, selected, sourceName, cancellationToken).Create(text);
    }

    internal static bool Probe(Stream source, string? name, long maximumInputBytes, CancellationToken token) {
        var prefix = new byte[68];
        if (maximumInputBytes < prefix.Length) return false;
        int read = 0;
        while (read < prefix.Length) {
            token.ThrowIfCancellationRequested();
            int count = source.Read(prefix, read, prefix.Length - read);
            if (count == 0) return false;
            read += count;
        }
        return PalmDatabase.HasSignature(prefix.AsSpan(60), "TEXtREAd") ||
            PalmDatabase.HasSignature(prefix.AsSpan(60), "BOOKMOBI");
    }

    private sealed partial class Projection(PalmDatabase database, MobiHeader first, MobiHeader selected,
        string sourceName, CancellationToken token) {
        private readonly List<OfficeDiagnostic> _diagnostics = new();
        private readonly HashSet<string> _reported = new(StringComparer.Ordinal);
        private readonly Dictionary<int, EpubManifestItem> _images = new();
        private readonly HashSet<int> _usedImages = new();

        internal OfficeConversionResult<EpubPublication> Create(byte[] text) {
            string title = TextMetadata(503).FirstOrDefault() ?? selected.HeaderTitle ?? first.HeaderTitle ??
                MobiHeader.Decode(selected.Encoding, database.Name);
            title = XmlText(WebUtility.HtmlDecode(title));
            if (string.IsNullOrWhiteSpace(title)) title = Path.GetFileNameWithoutExtension(sourceName);

            string language = TextMetadata(524).FirstOrDefault() ?? LanguageFromLocale(selected.Locale);
            string? identifier = TextMetadata(104).FirstOrDefault() ?? TextMetadata(113).FirstOrDefault();
            EpubPublication book;
            try { book = EpubPublication.Create(title, language, identifier); }
            catch (ArgumentException) when (language != "und") {
                Warn("MOBI_LANGUAGE", "The Kindle language was replaced with an undetermined language.", language);
                book = EpubPublication.Create(title, "und", identifier);
            }

            if (database.IsPalmDoc) {
                string body = selected.Encoding.GetString(text);
                var literal = new XElement("div");
                using var lines = new StringReader(body);
                int count = 0;
                while (lines.ReadLine() is { } line) {
                    token.ThrowIfCancellationRequested();
                    if (++count > 150_000) throw new InvalidDataException("PalmDOC text exceeds the paragraph limit.");
                    literal.Add(new XElement("p", XmlText(line)));
                }
                book.AddChapter("text", "EPUB/book.xhtml", title, literal.ToString(SaveOptions.DisableFormatting));
                if (database.Count - 1 > selected.TextRecords)
                    Warn("PALMDOC_BOOKMARKS_OMITTED", "PalmDOC records following the text are outside this reader profile.");
            } else {
                ApplyMetadata(book);
                AddImages(book);
                AddMarkup(book, text);
            }
            return new OfficeConversionResult<EpubPublication>(book, new OfficeConversionReport(sourceName, "EPUB", diagnostics: _diagnostics));
        }

        private IEnumerable<string> TextMetadata(int id) {
            string[] preferred = selected.Text(id).ToArray();
            return preferred.Length != 0 || ReferenceEquals(first, selected) ? preferred : first.Text(id);
        }

        private void ApplyMetadata(EpubPublication book) {
            foreach ((int id, string name) in new[] { (100, "creator"), (101, "publisher"), (103, "description"),
                (104, "identifier"), (105, "subject"), (106, "date"), (108, "contributor"), (109, "rights"),
                (112, "source"), (113, "identifier") }) {
                foreach (string value in TextMetadata(id)) book.AddDublinCoreMetadata(name, XmlText(WebUtility.HtmlDecode(value)));
            }
            if (TextMetadata(122).Any(value => value.Equals("true", StringComparison.OrdinalIgnoreCase)))
                Warn("MOBI_FIXED_LAYOUT_OMITTED", "The Kindle book was imported as reflowable content. Fixed-layout positioning was omitted.");
        }

        private void AddImages(EpubPublication book) {
            int? start = selected.Record != 0 ? first.FirstImage : selected.FirstImage;
            if (!start.HasValue || start.Value == 0) return;
            int end = selected.Record != 0 ? selected.Record - 1 : database.Count;
            int count = (selected.Record == 0 ? selected : first).Integer(125) ?? (end - start.Value);
            if (start.Value >= end || count < 0 || count > end - start.Value)
                throw new InvalidDataException("The Kindle resource record range is out of bounds.");
            if (count > 2_000) throw new InvalidDataException("The Kindle resource count exceeds 2,000.");
            for (int i = 0; i < count; i++) {
                token.ThrowIfCancellationRequested();
                ReadOnlySpan<byte> record = database.Record(start.Value + i);
                if (record.Length < 4 || record.Length > 8 * 1024 * 1024) continue;
                byte[] bytes = record.ToArray();
                if (!EpubImportedPublication.TryReadRaster(bytes, null, token, out OfficeImageInfo image)) continue;
                string id = "image-" + Number(i + 1);
                string extension = OfficeImageInfo.GetDefaultExtension(image.Format);
                string path = "EPUB/" + id + (extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension);
                _images.Add(i + 1, book.AddResource(id, path, image.MimeType, bytes));
            }
        }

        private void Warn(string code, string message, string? location = null, OfficeConversionLossKind loss = OfficeConversionLossKind.Omission) {
            if (_diagnostics.Count >= 1_000 || !_reported.Add(code + ":" + location)) return;
            _diagnostics.Add(new OfficeDiagnostic(code, OfficeDiagnosticSeverity.Warning, message, location, Component, loss));
        }
    }

    private static string Number(int value) => value.ToString(CultureInfo.InvariantCulture);

    private static string XmlText(string value) {
        StringBuilder? clean = null;
        for (int i = 0; i < value.Length; i++) {
            if (char.IsHighSurrogate(value[i]) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1])) {
                if (clean != null) clean.Append(value, i, 2);
                i++;
            } else if (System.Xml.XmlConvert.IsXmlChar(value[i])) clean?.Append(value[i]);
            else {
                clean ??= new StringBuilder(value.Substring(0, i));
                clean.Append('\uFFFD');
            }
        }
        return clean?.ToString() ?? value;
    }

    private static string LanguageFromLocale(uint value) {
        if (value is 0 or > int.MaxValue) return "und";
        try { return CultureInfo.GetCultureInfo((int)value).Name; }
        catch (CultureNotFoundException) { return "und"; }
    }
}
