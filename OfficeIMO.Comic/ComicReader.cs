using OfficeIMO.Drawing;
using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Archives;
using OfficeIMO.Epub;

namespace OfficeIMO.Comic;

/// <summary>Reads CBZ, CB7 and stored RAR5 comic archives as ordered raster pages with optional ComicInfo metadata.</summary>
public static partial class ComicReader {
    /// <summary>Maximum compressed comic archive input size.</summary>
    public const long MaximumInputBytes = 128L * 1024 * 1024;
    private const string Component = "OfficeIMO.Comic";

    /// <summary>Projects the current CBZ, CB7 or stored RAR5 CBR stream into a writable fixed-layout EPUB publication with fidelity evidence.</summary>
    /// <param name="source">The readable comic archive stream.</param>
    /// <param name="sourceName">The original filename, including the .cb7 or .cbr extension for a 7z or RAR5 comic.</param>
    /// <param name="maximumInputBytes">The caller's compressed input limit, capped by <see cref="MaximumInputBytes"/>.</param>
    /// <param name="cancellationToken">Cancels archive reading and publication projection.</param>
    public static OfficeConversionResult<EpubPublication> Import(Stream source, string sourceName = "comic.cbz",
        long maximumInputBytes = MaximumInputBytes, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceName);
        if (maximumInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumInputBytes));
        var diagnostics = new List<OfficeDiagnostic>();
        var omitted = new List<string>();
        string format = "CBZ";
        if (IsSevenZip(sourceName)) format = "CB7";
        else if (IsRar(sourceName)) format = "CBR";
        var options = new ArchiveReadOptions {
            MaxInputBytes = Math.Min(MaximumInputBytes, maximumInputBytes),
            MaxEntryCount = 2_000,
            MaxEntryBytes = 16L * 1024 * 1024,
            MaxExpandedBytes = 120L * 1024 * 1024,
            MaxPathDepth = 16,
            IncludeFile = path => {
                if (path.StartsWith("__MACOSX/", StringComparison.OrdinalIgnoreCase) ||
                    Path.GetFileName(path).StartsWith(".", StringComparison.Ordinal)) return false;
                if (IsComicInfo(path) || OfficeImageReader.IsKnownImageExtension(path)) return true;
                omitted.Add(path);
                return false;
            }
        };
        var entries = format switch {
            "CB7" => ArchiveReader.ReadSevenZip(source, options, cancellationToken),
            "CBR" => ArchiveReader.ReadRar5(source, options, cancellationToken),
            _ => ArchiveReader.ReadZip(source, options, cancellationToken)
        };
        foreach (string path in omitted)
            diagnostics.Add(Loss(format + "_MEMBER_OMITTED", "The archive member is not a supported comic page or ComicInfo.", path));

        XElement? metadata = ReadMetadata(entries, format, diagnostics, cancellationToken);
        var pages = entries.Where(entry => !IsComicInfo(entry.RelativePath))
            .OrderBy(entry => entry.RelativePath, ComicPageOrder.Instance).ToArray();
        if (pages.Length is 0 or > 1_000) throw new InvalidDataException($"A {format} archive must contain between 1 and 1,000 page images.");

        string title = Text(metadata, "Title") ?? Path.GetFileNameWithoutExtension(sourceName);
        string language = Text(metadata, "LanguageISO") ?? "und";
        EpubPublication book;
        try { book = EpubPublication.Create(title, language); }
        catch (ArgumentException) when (language != "und") {
            diagnostics.Add(Loss(format + "_LANGUAGE", "The invalid ComicInfo language was replaced with an undetermined language.", language));
            book = EpubPublication.Create(title, "und");
        }
        ApplyMetadata(book, metadata, format, diagnostics);
        book.SetRenditionLayout(EpubRenditionLayout.PrePaginated);
        var annotations = ReadPages(metadata, pages.Length, format, diagnostics);
        string? firstImage = null;
        string? coverImage = null;
        for (int i = 0; i < pages.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            ArchiveEntryData entry = pages[i];
            if (!EpubImportedPublication.TryReadRaster(entry.Data, entry.RelativePath, cancellationToken, out OfficeImageInfo image)) {
                diagnostics.Add(Loss(format + "_PAGE_UNAVAILABLE", "A page did not contain a supported raster image within the dimension limit.", entry.RelativePath));
                continue;
            }
            var annotation = annotations.TryGetValue(i, out var selected) ? selected : default;
            string number = (i + 1).ToString(CultureInfo.InvariantCulture);
            string label = annotation.Bookmark ?? "Page " + number;
            string imageId = "image-" + number;
            string imageName = imageId + OfficeImageInfo.GetDefaultExtension(image.Format);
            book.AddResource(imageId, "EPUB/" + imageName, image.MimeType, entry.Data);
            string pageId = "page-" + number;
            var markup = new XElement("div",
                new XElement("img", new XAttribute("src", imageName), new XAttribute("alt", label),
                    new XAttribute("width", image.Width), new XAttribute("height", image.Height)));
            book.AddChapter(pageId, "EPUB/" + pageId + ".xhtml", label, markup.ToString(SaveOptions.DisableFormatting));
            book.SetFixedLayoutPage(pageId, new EpubFixedLayoutPage(image.Width, image.Height), cancellationToken);
            firstImage ??= imageId;
            if (annotation.Cover) coverImage ??= imageId;
        }
        if (firstImage == null) throw new InvalidDataException($"The {format} archive contains no readable comic pages.");
        book.SetCoverImage(coverImage ?? firstImage);
        return new OfficeConversionResult<EpubPublication>(book, new OfficeConversionReport(sourceName, "EPUB", diagnostics: diagnostics));
    }

    internal static bool Probe(Stream source, string? name, long maximumInputBytes, CancellationToken token) {
        if (IsSevenZip(name)) return ProbeContainer(source, maximumInputBytes, false, token);
        if (IsRar(name)) return ProbeContainer(source, maximumInputBytes, true, token);
        try {
            using var archive = ArchiveInspector.Open(source, new ArchiveInspectionOptions {
                MaxInputBytes = maximumInputBytes, MaxEntryCount = 2_000, MaxEntryBytes = 16L * 1024 * 1024
            }, token);
            return archive.Format == ArchiveFormat.Zip && archive.Entries.Any(entry =>
                !entry.IsDirectory && entry.CanRead && OfficeImageReader.IsKnownImageExtension(entry.RawName));
        } catch (Exception error) when (error is InvalidDataException or NotSupportedException or IOException) {
            return false;
        }
    }

    private static bool ProbeContainer(Stream source, long maximumInputBytes, bool rar, CancellationToken token) {
        long? position = source.CanSeek ? source.Position : null;
        try {
            var options = new ArchiveReadOptions {
                MaxInputBytes = maximumInputBytes, MaxEntryCount = 2_000, MaxEntryBytes = 16L * 1024 * 1024,
                MaxExpandedBytes = 120L * 1024 * 1024, MaxPathDepth = 16
            };
            if (rar) {
                // Recognize a comic by its primary file table. Import diagnoses unsupported selected entries.
                return ArchiveReader.ListRar5(source, options, token).Any(entry => entry.Kind == ArchiveEntryKind.File &&
                    OfficeImageReader.IsKnownImageExtension(entry.RelativePath));
            }
            return ArchiveReader.ListSevenZip(source, options, token).Any(entry => entry.Kind == ArchiveEntryKind.File &&
                entry.Compression != ArchiveCompression.Unknown && OfficeImageReader.IsKnownImageExtension(entry.RelativePath));
        } catch (Exception error) when (error is not Rar5EncryptionNotSupportedException &&
            (error is InvalidDataException or NotSupportedException or IOException)) {
            return false;
        } finally {
            if (position.HasValue) source.Position = position.Value;
        }
    }

    private static bool IsSevenZip(string? sourceName) =>
        sourceName?.EndsWith(".cb7", StringComparison.OrdinalIgnoreCase) == true;

    private static bool IsRar(string? name) => name?.EndsWith(".cbr", StringComparison.OrdinalIgnoreCase) == true;

    private static OfficeDiagnostic Loss(string code, string message, string? path = null) =>
        new(code, OfficeDiagnosticSeverity.Warning, message, path, Component, OfficeConversionLossKind.Omission);

    private static bool IsComicInfo(string path) => path.Equals("ComicInfo.xml", StringComparison.OrdinalIgnoreCase);

    private static string? Text(XElement? metadata, string key) {
        string? value = metadata?.Element(key)?.Value.Trim();
        return string.IsNullOrEmpty(value) ? null : value;
    }
}
