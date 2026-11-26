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
            if (boundary <= first.TextRecords + 1 || !PalmDatabase.HasSignature(database.Record(boundary - 1), "BOUNDARY"))
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
            if (selected.HeaderTitleInvalidEncoding || first.HeaderTitleInvalidEncoding)
                ReportMetadataEncodingLoss("header title");
            if (selected.HeaderTitleInvalidRange || first.HeaderTitleInvalidRange)
                Warn("MOBI_HEADER_TITLE_RANGE_OMITTED", "The optional Mobipocket title is outside its header record.", "header title");
            string title = TextMetadata(503).FirstOrDefault() ?? NormalizeMetadata(selected.HeaderTitle) ??
                NormalizeMetadata(first.HeaderTitle) ?? ReadDatabaseName() ??
                NormalizeMetadata(DisplayStem(sourceName)) ?? "Untitled";

            string language = TextMetadata(524).FirstOrDefault() ?? LanguageFromLocale(selected.Locale);
            string? identifier = TextMetadata(104).FirstOrDefault() ?? TextMetadata(113).FirstOrDefault();
            EpubPublication book;
            try { book = EpubImportedPublication.Create(title, language, identifier); }
            catch (ArgumentException) when (language != "und") {
                Warn("MOBI_LANGUAGE", "The Kindle language was replaced with an undetermined language.", language);
                book = EpubImportedPublication.Create(title, "und", identifier);
            }

            if (database.IsPalmDoc) {
                AddLiteralText(book, DecodeBody(text));
                if (database.Count - 1 > selected.TextRecords)
                    Warn("PALMDOC_BOOKMARKS_OMITTED", "PalmDOC records following the text are outside this reader profile.");
            } else {
                ApplyMetadata(book, identifier);
                AddImages(book);
                AddMarkup(book, text);
            }
            return new OfficeConversionResult<EpubPublication>(book, new OfficeConversionReport(sourceName, "EPUB", diagnostics: _diagnostics));
        }

        private string? ReadDatabaseName() {
            if (MobiHeader.TryDecode(selected.Encoding, database.Name, out string name) ||
                (selected.Encoding.CodePage == 65001 &&
                 MobiHeader.TryDecode(MobiHeader.GetEncoding(1252), database.Name, out name)))
                return NormalizeMetadata(name);
            Warn("MOBI_DATABASE_NAME_OMITTED", "The Palm database name could not be decoded.");
            return null;
        }

        private static string DisplayStem(string name) {
            int start = Math.Max(name.LastIndexOf('/'), name.LastIndexOf('\\')) + 1;
            int extension = name.LastIndexOf('.');
            return name.Substring(start, (extension >= start ? extension : name.Length) - start);
        }

        private string DecodeBody(byte[] bytes) {
            try {
                return selected.Encoding.GetString(bytes);
            } catch (DecoderFallbackException error) {
                throw new InvalidDataException("The Mobipocket body text is invalid for the declared encoding.", error);
            }
        }

        private void AddLiteralText(EpubPublication book, string body) {
            var literal = new XElement("div");
            using var lines = new StringReader(body);
            long limit = new EpubPublicationLoadOptions().MaxEntryBytes;
            // Include the body envelope AddChapter parses around the fragment.
            long bytes = "<body xmlns='' xmlns:epub=''><div></div></body>".Length +
                EpubPublication.Html.NamespaceName.Length + EpubPublication.Ops.NamespaceName.Length;
            int count = 0;
            while (lines.ReadLine() is { } line) {
                token.ThrowIfCancellationRequested();
                if (++count > 150_000) throw new InvalidDataException("The Mobipocket plain text exceeds the paragraph limit.");
                string xmlText = XmlText(line);
                bytes += "<p></p>".Length + Encoding.UTF8.GetByteCount(xmlText);
                for (int i = 0; i < xmlText.Length; i++) {
                    if ((i & 4095) == 0) token.ThrowIfCancellationRequested();
                    bytes += xmlText[i] switch { '&' => 4, '<' or '>' => 3, _ => 0 };
                    if (bytes > limit) throw LiteralChapterLimit(limit);
                }
                if (bytes > limit) throw LiteralChapterLimit(limit);
                literal.Add(new XElement("p", xmlText));
            }
            try {
                book.AddChapter("text", "EPUB/book.xhtml", book.Title, literal.ToString(SaveOptions.DisableFormatting));
            } catch (InvalidDataException error) when (OfficeOutputLimit.Is(error)) {
                throw LiteralChapterLimit(limit, error);
            }
        }

        private static InvalidDataException LiteralChapterLimit(long limit, Exception? inner = null) =>
            new InvalidDataException($"The Mobipocket plain text exceeds the escaped XHTML chapter limit of {limit} bytes.", inner);

        private IEnumerable<string> TextMetadata(int id) {
            string? NormalizeField(string value) => NormalizeMetadata(id == 103 ? WebUtility.HtmlDecode(value) : value);
            void ReportLoss() => ReportMetadataEncodingLoss("EXTH " + Number(id));
            string[] preferred = selected.Text(id, ReportLoss).Select(NormalizeField).OfType<string>().ToArray();
            return preferred.Length != 0 || ReferenceEquals(first, selected)
                ? preferred
                : first.Text(id, ReportLoss).Select(NormalizeField).OfType<string>();
        }

        private void ReportMetadataEncodingLoss(string field) =>
            Warn("MOBI_METADATA_ENCODING_OMITTED", "An optional Mobipocket metadata value could not be decoded.", field);

        private static string? NormalizeMetadata(string? value) {
            if (value == null) return null;
            string normalized = XmlText(value).Trim();
            return normalized.Length == 0 ? null : normalized;
        }

        private void ApplyMetadata(EpubPublication book, string? primaryIdentifier) {
            foreach ((int id, string name) in new[] { (100, "creator"), (101, "publisher"), (103, "description"),
                (104, "identifier"), (105, "subject"), (106, "date"), (108, "contributor"), (109, "rights"),
                (112, "source"), (113, "identifier") }) {
                foreach (string value in TextMetadata(id)) {
                    if (name != "identifier" || value != primaryIdentifier) book.AddDublinCoreMetadata(name, value);
                }
            }
            string? progression = TextMetadata(527).FirstOrDefault()?.ToLowerInvariant();
            if (progression is "ltr" or "rtl" or "default")
                book.PageProgressionDirection = progression;
            else if (progression != null)
                Warn("MOBI_PAGE_PROGRESSION_OMITTED", "The Kindle page progression direction is not recognized.", "EXTH 527");
            if (TextMetadata(122).Any(value => value.Equals("true", StringComparison.OrdinalIgnoreCase)))
                Warn("MOBI_FIXED_LAYOUT_OMITTED", "The Kindle book was imported as reflowable content. Fixed-layout positioning was omitted.");
        }

        private void AddImages(EpubPublication book) {
            int end = selected.Record != 0 ? selected.Record - 1 : database.Count;
            int? declaredCount = (selected.Record == 0 ? selected : first).Integer(125);
            if (declaredCount == 0) return;
            int? start = selected.Record != 0 ? first.FirstImage : selected.FirstImage;
            if (start is null or 0) {
                if (!selected.IsKf8 || selected.Record != 0 || declaredCount is not (>= 1 and <= 2_000)) return;
                var inferred = FindStandaloneRasterRange(declaredCount.Value, end);
                if (inferred == null) return;
                Warn("MOBI_RESOURCE_RANGE_INFERRED", "The Kindle raster range was inferred from its declared resource count.",
                    loss: OfficeConversionLossKind.Approximation);
                for (int i = 0; i < inferred.Count; i++) {
                    token.ThrowIfCancellationRequested();
                    AddRasterResource(book, i + 1, inferred[i].Bytes, inferred[i].Image);
                }
                return;
            }
            int count = declaredCount ?? (end - start.Value);
            if (start.Value > end || count < 0 || count > end - start.Value)
                throw new InvalidDataException("The Kindle resource record range is out of bounds.");
            if (count > 2_000) throw new InvalidDataException("The Kindle resource count exceeds 2,000.");
            for (int i = 0; i < count; i++) {
                token.ThrowIfCancellationRequested();
                if (!TryReadRasterRecord(start.Value + i, out byte[] bytes, out OfficeImageInfo image)) continue;
                AddRasterResource(book, i + 1, bytes, image);
            }
        }

        private void AddRasterResource(EpubPublication book, int index, byte[] bytes, OfficeImageInfo image) {
            string id = "image-" + Number(index);
            string extension = OfficeImageInfo.GetDefaultExtension(image.Format);
            string path = "EPUB/" + id + (extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension);
            _images.Add(index, book.AddResource(id, path, image.MimeType, bytes));
        }

        private List<(byte[] Bytes, OfficeImageInfo Image)>? FindStandaloneRasterRange(int count, int end) {
            // Infer a range only when one contiguous raster block matches the number declared by the book.
            List<(byte[] Bytes, OfficeImageInfo Image)>? found = null;
            var run = new List<(byte[] Bytes, OfficeImageInfo Image)>();
            for (int index = selected.TextRecords + 1; index <= end; index++) {
                token.ThrowIfCancellationRequested();
                if (index < end && TryReadRasterRecord(index, out byte[] bytes, out OfficeImageInfo image)) {
                    // One extra entry marks a block that exceeds the declared count.
                    if (run.Count <= count) run.Add((bytes, image));
                    continue;
                }
                if (run.Count == count) {
                    if (found != null) return null;
                    found = run;
                    run = new List<(byte[] Bytes, OfficeImageInfo Image)>();
                } else {
                    run.Clear();
                }
            }
            return found;
        }

        private bool TryReadRasterRecord(int index, out byte[] data, out OfficeImageInfo image) {
            ReadOnlySpan<byte> record = database.Record(index);
            if (record.Length is >= 4 and <= 8 * 1024 * 1024) {
                data = record.ToArray();
                return EpubImportedPublication.TryReadRaster(data, null, token, out image);
            }
            data = Array.Empty<byte>();
            image = null!;
            return false;
        }

        private void Warn(string code, string message, string? location = null, OfficeConversionLossKind loss = OfficeConversionLossKind.Omission) {
            if (_diagnostics.Count >= 1_000 || !_reported.Add(code + ":" + location)) return;
            if (_diagnostics.Count == 999) {
                _diagnostics.Add(new OfficeDiagnostic("MOBI_DIAGNOSTICS_TRUNCATED", OfficeDiagnosticSeverity.Warning,
                    "Additional distinct Mobipocket diagnostics were omitted after reaching the report limit.",
                    null, Component, OfficeConversionLossKind.Omission));
                return;
            }
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
