namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class ImageParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Images) return false;
            var line = lines[i] ?? string.Empty;
            bool parsed = TryParseImage(line, options, state, out var img, out var sizeSpec, out var ranges);
            if (!parsed) {
                int captionIndex = i + 1;
                if (captionIndex >= lines.Length || !TryParseCaption(lines[captionIndex], out _)) {
                    return false;
                }

                if (!TryParseLinkedImageBlock(line, options, state, out img, out sizeSpec, out ranges)) {
                    return false;
                }
            }
            int absoluteLine = state.SourceLineOffset + i + 1;
            int startColumn = CountLeadingIndentColumns(line) + 1;
            img.SetMarkdownSyntaxMetadataSpans(
                CreateMetadataSpan(ranges.AltStart, ranges.AltLength),
                CreateMetadataSpan(ranges.SourceStart, ranges.SourceLength),
                CreateMetadataSpan(ranges.TitleStart, ranges.TitleLength),
                CreateMetadataSpan(ranges.LinkTargetStart, ranges.LinkTargetLength),
                CreateMetadataSpan(ranges.LinkTitleStart, ranges.LinkTitleLength));
            if (!string.IsNullOrWhiteSpace(sizeSpec))
            {
                foreach (var part in sizeSpec!.Split(new[]{' '}, StringSplitOptions.RemoveEmptyEntries))
                {
                    var kv = part.Split(new[]{'='}, 2);
                    if (kv.Length == 2)
                    {
                        var key = kv[0].Trim().ToLowerInvariant();
                        var val = kv[1].Trim();
                        if (key == "width" && double.TryParse(val, out var w)) img.Width = w;
                        if (key == "height" && double.TryParse(val, out var h)) img.Height = h;
                    }
                }
            }
            var resolvedPath = ResolveUrl(img.Path, options);
            if (string.IsNullOrWhiteSpace(resolvedPath)) {
                // Unsafe or invalid URL: let paragraph parsing treat this line as plain text.
                return false;
            }
            string? resolvedLink = null;
            if (!string.IsNullOrWhiteSpace(img.LinkUrl)) {
                string originalLink = img.LinkUrl!;
                resolvedLink = ResolveUrl(originalLink, options);
                if (string.IsNullOrWhiteSpace(resolvedLink)) {
                    return false;
                }
            }

            if (!string.Equals(resolvedPath, img.Path, StringComparison.Ordinal)
                || !string.Equals(resolvedLink, img.LinkUrl, StringComparison.Ordinal)) {
                var originalImage = img;
                var pictureSources = img.PictureSources
                    .Select(source => new ImagePictureSource(source.Path, source.Media, source.Type, source.Sizes, source.SrcSet))
                    .ToList();
                string? pictureFallbackPath = img.PictureFallbackPath;
                img = new ImageBlock(resolvedPath!, img.Alt, img.Title, img.Width, img.Height, resolvedLink, img.LinkTitle, img.LinkTarget, img.LinkRel, img.PlainAlt) {
                    Caption = img.Caption,
                    PictureFallbackPath = pictureFallbackPath
                };
                img.CopyMarkdownSyntaxMetadataSpansFrom(originalImage);
                foreach (var pictureSource in pictureSources) {
                    img.PictureSources.Add(pictureSource);
                }
            }
            int j = i + 1;
            if (j < lines.Length && TryParseCaption(lines[j], out var cap)) { img.Caption = cap; j++; }
            doc.Add(img); i = j; return true;

            MarkdownSourceSpan? CreateMetadataSpan(int? relativeStart, int? length) {
                if (!relativeStart.HasValue || !length.HasValue || length.Value <= 0) {
                    return null;
                }

                return CreateSpan(
                    state,
                    absoluteLine,
                    startColumn + relativeStart.Value,
                    absoluteLine,
                    startColumn + relativeStart.Value + length.Value - 1);
            }
        }
    }
}
