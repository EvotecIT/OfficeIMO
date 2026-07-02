using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private readonly struct MarkdownImageSyntaxRanges(
        int? altStart,
        int? altLength,
        int? sourceStart,
        int? sourceLength,
        int? titleStart,
        int? titleLength,
        int? linkTargetStart,
        int? linkTargetLength,
        int? linkTitleStart,
        int? linkTitleLength) {
        public int? AltStart { get; } = altStart;
        public int? AltLength { get; } = altLength;
        public int? SourceStart { get; } = sourceStart;
        public int? SourceLength { get; } = sourceLength;
        public int? TitleStart { get; } = titleStart;
        public int? TitleLength { get; } = titleLength;
        public int? LinkTargetStart { get; } = linkTargetStart;
        public int? LinkTargetLength { get; } = linkTargetLength;
        public int? LinkTitleStart { get; } = linkTitleStart;
        public int? LinkTitleLength { get; } = linkTitleLength;
    }

    private static bool IsImageLine(string line) {
        ImageBlock image;
        string? sizeSpec;
        MarkdownImageSyntaxRanges ranges;
        return TryParseImage(line, out image, out sizeSpec, out ranges);
    }
    private static bool TryParseImage(string line, out ImageBlock image, out string? sizeSpec) =>
        TryParseImage(line, new MarkdownReaderOptions(), new MarkdownReaderState(), out image, out sizeSpec, out _);

    private static bool TryParseImage(string line, out ImageBlock image, out string? sizeSpec, out MarkdownImageSyntaxRanges ranges) =>
        TryParseImage(line, new MarkdownReaderOptions(), new MarkdownReaderState(), out image, out sizeSpec, out ranges);

    private static bool TryParseImage(string line, MarkdownReaderOptions options, MarkdownReaderState state, out ImageBlock image, out string? sizeSpec) =>
        TryParseImage(line, options, state, out image, out sizeSpec, out _);

    private static bool TryParseImage(string line, MarkdownReaderOptions options, MarkdownReaderState state, out ImageBlock image, out string? sizeSpec, out MarkdownImageSyntaxRanges ranges) {
        image = null!;
        sizeSpec = null;
        ranges = default;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.Trim();
        if (!t.StartsWith("![", StringComparison.Ordinal)) return false;
        int altEnd = FindMatchingBracket(t, 1);
        if (altEnd < 2) return false;
        if (altEnd + 1 >= t.Length || t[altEnd + 1] != '(') return false;
        int parenClose = FindMatchingParen(t, altEnd + 1);
        if (parenClose <= altEnd + 2) return false;
        string alt = t.Substring(2, altEnd - 2);
        string inner = t.Substring(altEnd + 2, parenClose - (altEnd + 2));
        if (!TrySplitUrlAndOptionalTitle(inner, out var src, out var title, out int srcStart, out int srcLength, out int? titleStart, out int? titleLength)) {
            if (!TryParseTrimmedLiteralDestination(inner, out src, out srcStart, out srcLength)) return false;
            title = null;
            titleStart = null;
            titleLength = null;
        }
        string plainAlt = ExtractImageAltPlainText(alt, options, state);
        image = new ImageBlock(src, alt, title, plainAlt: plainAlt);
        ranges = new MarkdownImageSyntaxRanges(
            altStart: 2,
            altLength: altEnd - 2,
            sourceStart: altEnd + 2 + srcStart,
            sourceLength: srcLength,
            titleStart: titleStart.HasValue ? altEnd + 2 + titleStart.Value : null,
            titleLength: titleLength,
            linkTargetStart: null,
            linkTargetLength: null,
            linkTitleStart: null,
            linkTitleLength: null);
        ConsumeImageTrailingBlocks(t.Substring(parenClose + 1), image, options, ref sizeSpec);
        return true;
    }

    private static bool TryParseLinkedImageBlock(string line, out ImageBlock image, out string? sizeSpec) =>
        TryParseLinkedImageBlock(line, new MarkdownReaderOptions(), new MarkdownReaderState(), out image, out sizeSpec, out _);

    private static bool TryParseLinkedImageBlock(string line, out ImageBlock image, out string? sizeSpec, out MarkdownImageSyntaxRanges ranges) =>
        TryParseLinkedImageBlock(line, new MarkdownReaderOptions(), new MarkdownReaderState(), out image, out sizeSpec, out ranges);

    private static bool TryParseLinkedImageBlock(string line, MarkdownReaderOptions options, MarkdownReaderState state, out ImageBlock image, out string? sizeSpec) =>
        TryParseLinkedImageBlock(line, options, state, out image, out sizeSpec, out _);

    private static bool TryParseLinkedImageBlock(string line, MarkdownReaderOptions options, MarkdownReaderState state, out ImageBlock image, out string? sizeSpec, out MarkdownImageSyntaxRanges ranges) {
        image = null!;
        sizeSpec = null;
        ranges = default;
        if (string.IsNullOrEmpty(line)) {
            return false;
        }

        var t = line.Trim();
        if (!TryParseImageLink(
            t,
            0,
            null,
            out int consumed,
            out var alt,
            out var src,
            out var title,
            out var href,
            out var hrefTitle,
            out int altStart,
            out int altLength,
            out int srcStart,
            out int srcLength,
            out int? titleStart,
            out int? titleLength,
            out int hrefStart,
            out int hrefLength,
            out int? hrefTitleStart,
            out int? hrefTitleLength) || consumed <= 0) {
            return false;
        }

        string plainAlt = ExtractImageAltPlainText(alt, options, state);
        image = new ImageBlock(src, alt, title, plainAlt: plainAlt) {
            LinkUrl = href,
            LinkTitle = hrefTitle
        };
        ranges = new MarkdownImageSyntaxRanges(
            altStart,
            altLength,
            srcStart,
            srcLength,
            titleStart,
            titleLength,
            hrefStart,
            hrefLength,
            hrefTitleStart,
            hrefTitleLength);

        ConsumeImageTrailingBlocks(t.Substring(consumed), image, options, ref sizeSpec);

        return true;
    }

    private static void ConsumeImageTrailingBlocks(string? suffix, ImageBlock image, MarkdownReaderOptions options, ref string? sizeSpec) {
        var rest = suffix?.Trim();
        while (!string.IsNullOrEmpty(rest) && rest![0] == '{') {
            int close = MarkdownGenericAttributeParser.FindMatchingClosingBrace(rest, 0);
            if (close <= 0) {
                break;
            }

            var block = rest.Substring(1, close - 1).Trim();
            if (TryApplyImageSizeSpec(block, image)) {
                sizeSpec = block;
            } else if (options.GenericAttributes && MarkdownGenericAttributeParser.TryParseAttributeBlock(block, out var attributes)) {
                image.SetAttributes(attributes);
            } else {
                break;
            }

            rest = rest.Substring(close + 1).TrimStart();
        }
    }

    private static bool TryApplyImageSizeSpec(string? block, ImageBlock image) {
        bool applied = false;
        foreach (var part in (block ?? string.Empty).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
            int eq = part.IndexOf('=');
            if (eq <= 0) {
                continue;
            }

            var key = part.Substring(0, eq).Trim();
            var val = part.Substring(eq + 1).Trim();
            if (!double.TryParse(val, NumberStyles.Number, CultureInfo.InvariantCulture, out var num)) {
                continue;
            }

            if (string.Equals(key, "width", StringComparison.OrdinalIgnoreCase)) {
                image.Width = num;
                applied = true;
            } else if (string.Equals(key, "height", StringComparison.OrdinalIgnoreCase)) {
                image.Height = num;
                applied = true;
            }
        }

        return applied;
    }
}
