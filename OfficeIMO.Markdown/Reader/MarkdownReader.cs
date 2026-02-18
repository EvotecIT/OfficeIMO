using System.IO;
using System.Text;
// Intentionally avoid heavy regex use; simple scanning is used for resilience and speed.

namespace OfficeIMO.Markdown;

/// <summary>
/// Parses Markdown text into OfficeIMO.Markdown's typed object model (<see cref="MarkdownDoc"/>, blocks, and inlines).
///
/// Scope: intentionally focused on the syntax that <see cref="MarkdownDoc"/> currently emits so we can
/// round-trip what we generate. Reader behavior is controlled via <see cref="MarkdownReaderOptions"/>.
/// </summary>
public static partial class MarkdownReader {
    /// <summary>
    /// Parses Markdown text into a <see cref="MarkdownDoc"/> with typed blocks and basic inlines.
    /// </summary>
    public static MarkdownDoc Parse(string markdown, MarkdownReaderOptions? options = null) {
        options ??= new MarkdownReaderOptions();
        var state = new MarkdownReaderState();
        return ParseInternal(markdown, options, state, allowFrontMatter: true);
    }

    /// <summary>Parses a Markdown file path into a <see cref="MarkdownDoc"/>.</summary>
    public static MarkdownDoc ParseFile(string path, MarkdownReaderOptions? options = null) {
        string text = File.ReadAllText(path, Encoding.UTF8);
        return Parse(text, options);
    }

    private static MarkdownDoc ParseInternal(string markdown, MarkdownReaderOptions options, MarkdownReaderState state, bool allowFrontMatter) {
        var doc = MarkdownDoc.Create();
        if (string.IsNullOrEmpty(markdown)) return doc;

        // Normalize BOM (U+FEFF) at the very beginning to avoid blocking heading/html detection
        if (markdown[0] == '\uFEFF') {
            markdown = markdown.Substring(1);
        }

        var preParseNormalization = CreatePreParseNormalizationOptions(options.InputNormalization);
        if (preParseNormalization != null) {
            markdown = MarkdownInputNormalizer.Normalize(markdown, preParseNormalization);
        }

        // Normalize line endings and split. Keep empty lines significant for block boundaries.
        var text = markdown.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = text.Split('\n');
        int i = 0;

        // Front matter (YAML) only if it's the very first thing in the file
        if (allowFrontMatter && options.FrontMatter && i < lines.Length && lines[i].Trim() == "---") {
            int start = i + 1;
            int end = -1;
            for (int j = start; j < lines.Length; j++) { if (lines[j].Trim() == "---") { end = j; break; } }
            if (end > start) {
                var dict = ParseFrontMatter(lines, start, end - 1);
                if (dict.Count > 0) doc.Add(FrontMatterBlock.FromObject(dict));
                i = end + 1;
                // optional blank line after front matter
                if (i < lines.Length && string.IsNullOrWhiteSpace(lines[i])) i++;
            }
        }

        var pipeline = MarkdownReaderPipeline.Default(options);
        // Pre-scan for reference-style link definitions so inline refs in earlier paragraphs can resolve
        PreScanReferenceLinkDefinitions(lines, state, options);
        while (i < lines.Length) {
            if (string.IsNullOrWhiteSpace(lines[i])) { i++; continue; }
            bool matched = false;
            var parsers = pipeline.Parsers;
            for (int p = 0; p < parsers.Count; p++) {
                if (parsers[p].TryParse(lines, ref i, options, doc, state)) { matched = true; break; }
            }
            if (!matched) i++; // defensive: avoid infinite loop
        }

        return doc;
    }

    private static void PreScanReferenceLinkDefinitions(string[] lines, MarkdownReaderState state) {
        PreScanReferenceLinkDefinitions(lines, state, new MarkdownReaderOptions());
    }

    private static void PreScanReferenceLinkDefinitions(string[] lines, MarkdownReaderState state, MarkdownReaderOptions options) {
        bool inFence = false;
        char fenceChar = '\0';
        int fenceLen = 0;

        for (int idx = 0; idx < lines.Length; idx++) {
            var line = lines[idx]; if (string.IsNullOrWhiteSpace(line)) continue;

            // Ignore anything inside fenced code blocks.
            if (!inFence) {
                if (IsCodeFenceOpen(line, out _, out fenceChar, out fenceLen)) {
                    inFence = true;
                    continue;
                }
            } else {
                if (IsCodeFenceClose(line, fenceChar, fenceLen)) {
                    inFence = false;
                }
                continue;
            }

            // Ignore indented code blocks (4+ leading spaces or a tab). Reference definitions are only valid
            // up to 3 leading spaces in typical Markdown implementations.
            int leading = 0;
            while (leading < line.Length && line[leading] == ' ') leading++;
            if (leading >= 4) continue;
            if (leading < line.Length && line[leading] == '\t') continue;

            var t = line.Trim(); if (t.Length < 5 || t[0] != '[') continue;
            if (t.Length > 1 && t[1] == '^') continue; // footnote definition, not a link ref
            int rb = t.IndexOf(']'); if (rb <= 1) continue;
            if (rb + 1 >= t.Length || t[rb + 1] != ':') continue;
            string label = NormalizeReferenceLabel(t.Substring(1, rb - 1));
            string rest = t.Substring(rb + 2).Trim(); if (string.IsNullOrEmpty(rest)) continue;
            if (!TrySplitUrlAndOptionalTitle(rest, out var url, out var title)) continue;
            if (!string.IsNullOrEmpty(label) && !string.IsNullOrEmpty(url)) {
                var resolved = ResolveUrl(url, options);
                if (!string.IsNullOrEmpty(resolved)) state.LinkRefs[label] = (resolved!, title);
            }
        }
    }

    private static string NormalizeReferenceLabel(string? label) {
        if (string.IsNullOrWhiteSpace(label)) return string.Empty;
        var t = label!.Trim();
        var sb = new System.Text.StringBuilder(t.Length);
        bool prevSpace = false;
        for (int i = 0; i < t.Length; i++) {
            char c = t[i];
            if (char.IsWhiteSpace(c)) {
                if (!prevSpace) sb.Append(' ');
                prevSpace = true;
            } else {
                sb.Append(c);
                prevSpace = false;
            }
        }
        return sb.ToString();
    }

    private static MarkdownReaderOptions CloneOptionsWithoutFrontMatter(MarkdownReaderOptions source) {
        return new MarkdownReaderOptions {
            FrontMatter = false,
            Callouts = source.Callouts,
            Headings = source.Headings,
            FencedCode = source.FencedCode,
            IndentedCodeBlocks = source.IndentedCodeBlocks,
            Images = source.Images,
            UnorderedLists = source.UnorderedLists,
            OrderedLists = source.OrderedLists,
            Tables = source.Tables,
            DefinitionLists = source.DefinitionLists,
            HtmlBlocks = source.HtmlBlocks,
            Paragraphs = source.Paragraphs,
            AutolinkUrls = source.AutolinkUrls,
            AutolinkWwwUrls = source.AutolinkWwwUrls,
            AutolinkWwwScheme = source.AutolinkWwwScheme,
            AutolinkEmails = source.AutolinkEmails,
            BackslashHardBreaks = source.BackslashHardBreaks,
            InlineHtml = source.InlineHtml,
            BaseUri = source.BaseUri,
            DisallowScriptUrls = source.DisallowScriptUrls,
            DisallowFileUrls = source.DisallowFileUrls,
            AllowMailtoUrls = source.AllowMailtoUrls,
            AllowDataUrls = source.AllowDataUrls,
            AllowProtocolRelativeUrls = source.AllowProtocolRelativeUrls,
            RestrictUrlSchemes = source.RestrictUrlSchemes,
            AllowedUrlSchemes = source.AllowedUrlSchemes,
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeSoftWrappedStrongSpans = source.InputNormalization?.NormalizeSoftWrappedStrongSpans ?? false,
                NormalizeInlineCodeSpanLineBreaks = source.InputNormalization?.NormalizeInlineCodeSpanLineBreaks ?? false,
                NormalizeEscapedInlineCodeSpans = source.InputNormalization?.NormalizeEscapedInlineCodeSpans ?? false,
                NormalizeTightStrongBoundaries = source.InputNormalization?.NormalizeTightStrongBoundaries ?? false,
                NormalizeLooseStrongDelimiters = source.InputNormalization?.NormalizeLooseStrongDelimiters ?? false,
                NormalizeOrderedListMarkerSpacing = source.InputNormalization?.NormalizeOrderedListMarkerSpacing ?? false,
                NormalizeOrderedListParenMarkers = source.InputNormalization?.NormalizeOrderedListParenMarkers ?? false,
                NormalizeOrderedListCaretArtifacts = source.InputNormalization?.NormalizeOrderedListCaretArtifacts ?? false,
                NormalizeTightParentheticalSpacing = source.InputNormalization?.NormalizeTightParentheticalSpacing ?? false,
                NormalizeNestedStrongDelimiters = source.InputNormalization?.NormalizeNestedStrongDelimiters ?? false
            }
        };
    }

    private static MarkdownInputNormalizationOptions? CreatePreParseNormalizationOptions(MarkdownInputNormalizationOptions? source) {
        bool normalizeSoftWrappedStrong = source?.NormalizeSoftWrappedStrongSpans ?? false;
        bool normalizeInlineCodeLineBreaks = source?.NormalizeInlineCodeSpanLineBreaks ?? false;
        bool normalizeLooseStrongDelimiters = source?.NormalizeLooseStrongDelimiters ?? false;
        bool normalizeOrderedListMarkerSpacing = source?.NormalizeOrderedListMarkerSpacing ?? false;
        bool normalizeOrderedListParenMarkers = source?.NormalizeOrderedListParenMarkers ?? false;
        bool normalizeOrderedListCaretArtifacts = source?.NormalizeOrderedListCaretArtifacts ?? false;
        bool normalizeTightParentheticalSpacing = source?.NormalizeTightParentheticalSpacing ?? false;
        bool normalizeNestedStrongDelimiters = source?.NormalizeNestedStrongDelimiters ?? false;

        if (!normalizeSoftWrappedStrong
            && !normalizeInlineCodeLineBreaks
            && !normalizeLooseStrongDelimiters
            && !normalizeOrderedListMarkerSpacing
            && !normalizeOrderedListParenMarkers
            && !normalizeOrderedListCaretArtifacts
            && !normalizeTightParentheticalSpacing
            && !normalizeNestedStrongDelimiters) {
            return null;
        }

        return new MarkdownInputNormalizationOptions {
            NormalizeSoftWrappedStrongSpans = normalizeSoftWrappedStrong,
            NormalizeInlineCodeSpanLineBreaks = normalizeInlineCodeLineBreaks,
            NormalizeLooseStrongDelimiters = normalizeLooseStrongDelimiters,
            NormalizeOrderedListMarkerSpacing = normalizeOrderedListMarkerSpacing,
            NormalizeOrderedListParenMarkers = normalizeOrderedListParenMarkers,
            NormalizeOrderedListCaretArtifacts = normalizeOrderedListCaretArtifacts,
            NormalizeTightParentheticalSpacing = normalizeTightParentheticalSpacing,
            NormalizeNestedStrongDelimiters = normalizeNestedStrongDelimiters
        };
    }

    private static MarkdownReaderState CloneState(MarkdownReaderState state) {
        var clone = new MarkdownReaderState();
        foreach (var kvp in state.LinkRefs) clone.LinkRefs[kvp.Key] = kvp.Value;
        return clone;
    }
}
