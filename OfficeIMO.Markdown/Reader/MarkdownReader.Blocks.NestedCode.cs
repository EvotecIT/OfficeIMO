using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    private static bool TryParseNestedFencedCodeBlock(
        string[] lines,
        ref int index,
        int continuationIndent,
        MarkdownLoadOptions options,
        MarkdownReaderState state,
        out IMarkdownBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.FencedCode) return false;

        var attributeSourceText = string.Empty;
        MarkdownAttributeSet attributeSet = MarkdownAttributeSet.Empty;
        MarkdownSourceSpan? attributeSourceSpan = null;
        if (TryConsumeNestedStandaloneGenericAttributeLine(
                lines,
                index,
                continuationIndent,
                options,
                state,
                NestedStandaloneGenericAttributeTarget.FencedCode,
                out attributeSet,
                out attributeSourceText,
                out attributeSourceSpan)) {
            index++;
            if (index >= lines.Length) {
                return false;
            }
        }

        string line = lines[index] ?? string.Empty;
        int indent = CountLeadingIndentColumns(line);
        if (indent < continuationIndent) return false;

        string first = StripLeadingIndentColumns(line, continuationIndent);
        int openingFenceIndent = CountLeadingIndentColumns(first);
        if (openingFenceIndent > 3) return false;

        if (!IsCodeFenceOpen(first, out string language, out char fenceChar, out int fenceLen, out int fenceIndentColumns, out int infoPaddingColumns, out int infoPaddingCharacters)) return false;

        int j = index + 1;
        int lineLimit = lines.Length - (state.HasSyntheticTerminalLine ? 1 : 0);
        var code = new StringBuilder();
        bool hasClosingFence = false;
        int closingFenceIndentColumns = 0;
        int closingFenceLength = fenceLen;
        while (j < lineLimit) {
            string raw = lines[j] ?? string.Empty;
            int ind = CountLeadingIndentColumns(raw);
            string sliced = ind >= continuationIndent ? StripLeadingIndentColumns(raw, continuationIndent) : raw.TrimStart();
            if (TryGetCodeFenceCloseInfo(sliced, fenceChar, fenceLen, out closingFenceIndentColumns, out closingFenceLength)) {
                hasClosingFence = true;
                j++;
                break;
            }
            int contentIndentToStrip = Math.Min(openingFenceIndent, CountLeadingIndentColumns(sliced));
            if (contentIndentToStrip > 0) {
                sliced = StripLeadingIndentColumns(sliced, contentIndentToStrip);
            }
            code.AppendLine(sliced);
            j++;
        }

        var contentLineCount = Math.Max(0, j - index - (hasClosingFence ? 2 : 1));
        var content = RemoveSingleTrailingLineEnding(code.ToString());
        string? caption = null;
        // Optional caption line (indented like other nested content)
        if (j < lineLimit) {
            var capLine = lines[j] ?? string.Empty;
            if (CountLeadingIndentColumns(capLine) >= continuationIndent) {
                var capSlice = StripLeadingIndentColumns(capLine, continuationIndent);
                if (TryParseCaption(capSlice, out var cap)) { caption = cap; j++; }
            }
        }

        block = CreateParsedFencedBlock(
            language,
            content,
            isFenced: true,
            caption,
            options,
            fenceIndentColumns,
            fenceLen,
            infoPaddingColumns,
            infoPaddingCharacters,
            fenceChar,
            hasClosingFence,
            openingFenceIndent + closingFenceIndentColumns,
            closingFenceLength,
            contentLineCount);
        if (!attributeSet.IsEmpty && block is MarkdownObject markdownObject && markdownObject.Attributes.IsEmpty) {
            markdownObject.SetAttributes(attributeSet);
            MarkdownGenericAttributeSourceSpans.Set(markdownObject, attributeSourceText, attributeSourceSpan);
        }

        index = j;
        return true;
    }

    private enum NestedStandaloneGenericAttributeTarget {
        Any,
        FencedCode,
        List,
        Paragraph,
        CustomContainer,
        Table
    }

    private static bool TryParseNestedIndentedCodeBlock(string[] lines, ref int index, int continuationIndent, MarkdownLoadOptions options, out CodeBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.IndentedCodeBlocks) return false;

        string line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) return false;

        int spaces = CountLeadingIndentColumns(line);
        int required = continuationIndent + 4;
        if (spaces < required) return false;

        int j = index;
        var sb = new StringBuilder();
        while (j < lines.Length) {
            string cur = lines[j] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cur)) {
                if (!HasIndentedCodeContinuationAfterBlankLines(lines, j, required)) break;
                sb.AppendLine(StripLeadingIndentColumns(cur, required));
                j++;
                continue;
            }

            int curSpaces = CountLeadingIndentColumns(cur);
            if (curSpaces < required) break;
            sb.AppendLine(StripLeadingIndentColumns(cur, required));
            j++;
        }

        block = new CodeBlock(string.Empty, RemoveSingleTrailingLineEnding(sb.ToString()), isFenced: false);
        index = j;
        return true;
    }

}
