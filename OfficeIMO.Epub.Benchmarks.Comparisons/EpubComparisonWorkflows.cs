using System.Globalization;
using System.Text;
using HtmlAgilityPack;
using OfficeIMO.Epub;
using VersOneBook = VersOne.Epub.EpubBook;

namespace OfficeIMO.Epub.Benchmarks.Comparisons;

internal static class EpubComparisonWorkflows {
    private static readonly EpubReadOptions OfficeIMOOptions = new() {
        IncludeRawHtml = true,
        IncludeResourceData = false,
        PreferSpineOrder = true,
        DeterministicOrder = true
    };

    internal static long ReadOfficeIMO(byte[] package) {
        EpubDocument document = LoadOfficeIMO(package);
        long checksum = MetadataChecksum(document.Title, document.Creator, document.Language);
        foreach (EpubChapter chapter in document.Chapters) {
            checksum = checked(checksum + chapter.Path.Length + (chapter.Html?.Length ?? 0) + chapter.Text.Length);
        }

        return checksum;
    }

    internal static long ReadVersOne(byte[] package) {
        VersOneBook book = LoadVersOne(package);
        string? language = book.Schema.Package.Metadata.Languages.FirstOrDefault()?.Language;
        long checksum = MetadataChecksum(book.Title, book.Author, language);
        var extractedText = new List<string>(book.ReadingOrder.Count);
        foreach (global::VersOne.Epub.EpubLocalTextContentFile chapter in book.ReadingOrder) {
            string text = ExtractVersOneText(chapter.Content);
            extractedText.Add(text);
            checksum = checked(checksum + chapter.FilePath.Length + chapter.Content.Length + text.Length);
        }

        return checksum;
    }

    internal static EpubReadEvidence InspectOfficeIMO(byte[] package) {
        EpubDocument document = LoadOfficeIMO(package);
        return Inspect(
            "OfficeIMO",
            package.LongLength,
            document.Title,
            document.Creator,
            document.Language,
            document.Chapters.Select(chapter =>
                new ChapterContent(chapter.Path, chapter.Html ?? string.Empty, chapter.Text)));
    }

    internal static EpubReadEvidence InspectVersOne(byte[] package) {
        VersOneBook book = LoadVersOne(package);
        return Inspect(
            "VersOne.Epub",
            package.LongLength,
            book.Title,
            book.Author,
            book.Schema.Package.Metadata.Languages.FirstOrDefault()?.Language,
            book.ReadingOrder.Select(chapter =>
                new ChapterContent(chapter.FilePath, chapter.Content, ExtractVersOneText(chapter.Content))));
    }

    private static EpubDocument LoadOfficeIMO(byte[] package) {
        using var stream = new MemoryStream(package, writable: false);
        return EpubDocument.Load(stream, OfficeIMOOptions);
    }

    private static VersOneBook LoadVersOne(byte[] package) {
        using var stream = new MemoryStream(package, writable: false);
        return global::VersOne.Epub.EpubReader.ReadBook(stream);
    }

    private static EpubReadEvidence Inspect(
        string implementation,
        long inputBytes,
        string? title,
        string? creator,
        string? language,
        IEnumerable<ChapterContent> chapters) {
        var chapterCount = 0;
        long contentCharacters = 0;
        long textCharacters = 0;
        ulong pathHash = FnvOffset;
        ulong contentHash = FnvOffset;
        ulong textHash = FnvOffset;
        string? firstPath = null;
        string? lastPath = null;
        foreach (ChapterContent chapter in chapters) {
            chapterCount++;
            contentCharacters = checked(contentCharacters + chapter.Content.Length);
            textCharacters = checked(textCharacters + chapter.Text.Length);
            firstPath ??= chapter.Path;
            lastPath = chapter.Path;
            pathHash = Hash(pathHash, chapter.Path);
            contentHash = Hash(contentHash, chapter.Content);
            textHash = Hash(textHash, chapter.Text);
        }

        return new EpubReadEvidence(
            implementation,
            inputBytes,
            title,
            creator,
            language,
            chapterCount,
            contentCharacters,
            textCharacters,
            firstPath,
            lastPath,
            pathHash.ToString("X16", CultureInfo.InvariantCulture),
            contentHash.ToString("X16", CultureInfo.InvariantCulture),
            textHash.ToString("X16", CultureInfo.InvariantCulture));
    }

    private static string ExtractVersOneText(string html) {
        var document = new HtmlDocument();
        document.LoadHtml(html);
        HtmlNode scope = document.DocumentNode.SelectSingleNode("//body") ?? document.DocumentNode;
        HtmlNodeCollection? textNodes = scope.SelectNodes(".//text()");
        if (textNodes == null || textNodes.Count == 0) return string.Empty;

        var text = new StringBuilder(html.Length);
        foreach (HtmlNode node in textNodes) {
            string value = HtmlEntity.DeEntitize(node.InnerText);
            if (string.IsNullOrWhiteSpace(value)) continue;
            text.Append(value);
            text.Append(' ');
        }
        return NormalizeWhitespace(text.ToString());
    }

    private static string NormalizeWhitespace(string value) {
        var normalized = new StringBuilder(value.Length);
        bool pendingSpace = false;
        foreach (char character in value) {
            if (char.IsWhiteSpace(character)) {
                pendingSpace = normalized.Length > 0;
            } else {
                if (pendingSpace) normalized.Append(' ');
                normalized.Append(character);
                pendingSpace = false;
            }
        }
        return normalized.ToString();
    }

    private static long MetadataChecksum(string? title, string? creator, string? language) =>
        (title?.Length ?? 0) + (creator?.Length ?? 0) + (language?.Length ?? 0);

    private static ulong Hash(ulong hash, string value) {
        foreach (char character in value) {
            hash ^= character;
            hash *= FnvPrime;
        }

        return hash;
    }

    private sealed record ChapterContent(string Path, string Content, string Text);

    private const ulong FnvOffset = 14695981039346656037UL;
    private const ulong FnvPrime = 1099511628211UL;
}
