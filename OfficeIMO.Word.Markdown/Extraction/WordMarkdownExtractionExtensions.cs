using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading;

namespace OfficeIMO.Word.Markdown;

/// <summary>
/// Chunked extraction helpers intended for AI ingestion.
/// </summary>
public static class WordMarkdownExtractionExtensions {
    /// <summary>
    /// Extracts the document as block-aligned Markdown chunks with basic location metadata.
    /// </summary>
    /// <remarks>
    /// This method is dependency-free and deterministic.
    /// Chunks are best-effort capped by <see cref="WordMarkdownChunkingOptions.MaxChars"/> and only split between blocks.
    /// </remarks>
    public static IEnumerable<WordMarkdownChunk> ExtractMarkdownChunks(
        this WordDocument document,
        WordToMarkdownOptions? markdownOptions = null,
        WordMarkdownChunkingOptions? chunking = null,
        string? sourcePath = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        markdownOptions ??= new WordToMarkdownOptions();
        chunking ??= new WordMarkdownChunkingOptions();

        if (chunking.MaxChars < 256) chunking.MaxChars = 256;

        var converter = new WordToMarkdownConverter();
        var headingStack = new List<(int Level, string Text)>();

        var current = new StringBuilder(capacity: Math.Min(chunking.MaxChars, 16_384));
        int chunkIndex = 0;
        int? firstBlockIndex = null;
        string? firstHeadingPath = null;
        var warnings = new List<string>(capacity: 2);

        int blockIndex = 0;

        foreach (var section in DocumentTraversal.EnumerateSections(document)) {
            cancellationToken.ThrowIfCancellationRequested();

            var elements = section.Elements;
            if (elements == null || elements.Count == 0) {
                elements = new List<WordElement>(section.Paragraphs.Count + section.Tables.Count);
                elements.AddRange(section.Paragraphs);
                elements.AddRange(section.Tables);
            }

            for (int i = 0; i < elements.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();

                var el = elements[i];
                string? blockMarkdown = null;

                if (el is WordParagraph p) {
                    bool hasRuns = false;
                    try {
                        hasRuns = p.GetRuns().Any();
                    } catch (InvalidOperationException ex) {
                        Debug.WriteLine($"GetRuns() failed for paragraph during chunk extraction: {ex.Message}");
                        hasRuns = false;
                    }

                    // Only render once per underlying OpenXml paragraph (see converter logic).
                    // Also look ahead to detect checkbox state across wrappers.
                    bool paraHasCheckbox = p.IsCheckBox;
                    bool paraCheckboxChecked = p.CheckBox?.IsChecked == true;
                    int j = i + 1;
                    while (j < elements.Count && elements[j] is WordParagraph p2 && p2.Equals(p)) {
                        if (p2.IsCheckBox) { paraHasCheckbox = true; paraCheckboxChecked = p2.CheckBox?.IsChecked == true; }
                        j++;
                    }
                    if (hasRuns && !p.IsFirstRun) continue;

                    // Track heading path for better chunk metadata (best-effort).
                    int? headingLevel = p.Style.HasValue
                        ? HeadingStyleMapper.GetLevelForHeadingStyle(p.Style.Value)
                        : (int?)null;
                    if (headingLevel.HasValue && headingLevel.Value > 0) {
                        var headingText = SafePlainTextFromParagraph(p, markdownOptions, converter);
                        UpdateHeadingStack(headingStack, headingLevel.Value, headingText);
                    }

                    blockMarkdown = converter.ConvertParagraph(p, markdownOptions, paraHasCheckbox, paraCheckboxChecked);
                } else if (el is WordTable t) {
                    blockMarkdown = converter.ConvertTable(t, markdownOptions);
                } else if (el is WordEmbeddedDocument ed) {
                    var html = ed.GetHtml();
                    if (!string.IsNullOrWhiteSpace(html)) blockMarkdown = html!.TrimEnd();
                }

                if (string.IsNullOrWhiteSpace(blockMarkdown)) continue;

                var headingPath = BuildHeadingPath(headingStack);
                AppendBlockOrFlush(
                    chunking,
                    sourcePath,
                    headingPath,
                    blockMarkdown!,
                    blockIndex,
                    ref chunkIndex,
                    ref firstBlockIndex,
                    ref firstHeadingPath,
                    current,
                    warnings,
                    out var flushed);

                if (flushed != null) yield return flushed;
                blockIndex++;
            }
        }

        // Footnotes as a final block (same as converter, but chunk-aware).
        if (chunking.IncludeFootnotes && document.FootNotes.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var sb = new StringBuilder();
            sb.AppendLine("## Footnotes");
            sb.AppendLine();
            foreach (var footnote in document.FootNotes.OrderBy(fn => fn.ReferenceId)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (footnote.ReferenceId.HasValue) {
                    sb.Append("[^").Append(footnote.ReferenceId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append("]: ");
                    sb.AppendLine(converter.RenderFootnote(footnote, markdownOptions));
                }
            }

            var headingPath = BuildHeadingPath(headingStack);
            AppendBlockOrFlush(
                chunking,
                sourcePath,
                headingPath,
                sb.ToString().TrimEnd(),
                blockIndex,
                ref chunkIndex,
                ref firstBlockIndex,
                ref firstHeadingPath,
                current,
                warnings,
                out var flushed);
            if (flushed != null) yield return flushed;
            blockIndex++;
        }

        // Final flush.
        if (current.Length > 0) {
            yield return BuildChunk(
                sourcePath,
                chunkIndex,
                firstBlockIndex,
                firstHeadingPath,
                current.ToString().TrimEnd(),
                warnings.Count > 0 ? warnings.ToArray() : null);
        }
    }

    private static void AppendBlockOrFlush(
        WordMarkdownChunkingOptions chunking,
        string? sourcePath,
        string? headingPath,
        string blockMarkdown,
        int blockIndex,
        ref int chunkIndex,
        ref int? firstBlockIndex,
        ref string? firstHeadingPath,
        StringBuilder current,
        List<string> warnings,
        out WordMarkdownChunk? flushed) {
        flushed = null;

        // Add a blank line between blocks.
        string block = blockMarkdown.TrimEnd();
        if (block.Length > chunking.MaxChars) {
            // Hard-cap a single pathological block so callers don't accidentally ingest megabytes in one chunk.
            block = block.Substring(0, chunking.MaxChars) + "\n\n<!-- truncated -->";
            warnings.Add("A single block exceeded MaxChars and was truncated.");
        }
        int extra = (current.Length == 0 ? 0 : 2) + block.Length;

        if (current.Length > 0 && (current.Length + extra) > chunking.MaxChars) {
            flushed = BuildChunk(
                sourcePath,
                chunkIndex,
                firstBlockIndex,
                firstHeadingPath,
                current.ToString().TrimEnd(),
                warnings.Count > 0 ? warnings.ToArray() : null);
            chunkIndex++;
            current.Clear();
            warnings.Clear();
            firstBlockIndex = null;
            firstHeadingPath = null;
        }

        if (firstBlockIndex == null) {
            firstBlockIndex = blockIndex;
            firstHeadingPath = headingPath;
        }

        if (current.Length > 0) current.AppendLine().AppendLine();
        current.Append(block);
    }

    private static WordMarkdownChunk BuildChunk(
        string? sourcePath,
        int chunkIndex,
        int? firstBlockIndex,
        string? headingPath,
        string markdown,
        string[]? warnings) {
        var id = BuildStableId("word-md", sourcePath, chunkIndex, firstBlockIndex);
        return new WordMarkdownChunk {
            Id = id,
            Location = new WordMarkdownLocation {
                Path = sourcePath,
                BlockIndex = firstBlockIndex,
                HeadingPath = headingPath
            },
            Text = markdown,
            Markdown = markdown,
            Warnings = warnings
        };
    }

    private static string? BuildHeadingPath(List<(int Level, string Text)> stack) {
        if (stack.Count == 0) return null;
        var sb = new StringBuilder();
        for (int i = 0; i < stack.Count; i++) {
            if (i > 0) sb.Append(" > ");
            sb.Append(stack[i].Text);
        }
        var s = sb.ToString().Trim();
        return s.Length == 0 ? null : s;
    }

    private static void UpdateHeadingStack(List<(int Level, string Text)> stack, int level, string text) {
        if (level < 1) return;
        if (string.IsNullOrWhiteSpace(text)) text = $"Heading {level}";

        // Remove any existing headings at this level or deeper, then push.
        for (int i = stack.Count - 1; i >= 0; i--) {
            if (stack[i].Level >= level) stack.RemoveAt(i);
        }
        stack.Add((level, text.Trim()));
    }

    private static string SafePlainTextFromParagraph(WordParagraph paragraph, WordToMarkdownOptions opt, WordToMarkdownConverter converter) {
        try {
            // Reuse run rendering (may include basic markup, but is good enough for a heading path label).
            var s = converter.RenderRuns(paragraph, opt);
            return string.IsNullOrWhiteSpace(s) ? "Heading" : CollapseWhitespace(s);
        } catch {
            return "Heading";
        }
    }

    private static string CollapseWhitespace(string text) {
        var sb = new StringBuilder(text.Length);
        bool prevWs = false;
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            bool ws = char.IsWhiteSpace(c);
            if (ws) {
                if (!prevWs) sb.Append(' ');
                prevWs = true;
            } else {
                sb.Append(c);
                prevWs = false;
            }
        }
        return sb.ToString().Trim();
    }

    private static string BuildStableId(string kind, string? path, int chunkIndex, int? blockIndex) {
        // Keep IDs short, stable and ASCII-only; avoid leaking full paths when caller doesn't provide one.
        var safe = string.IsNullOrWhiteSpace(path) ? "memory" : System.IO.Path.GetFileName(path!.Trim());
        var b = blockIndex.HasValue ? blockIndex.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "na";
        return $"{kind}:{safe}:c{chunkIndex}:b{b}";
    }
}
