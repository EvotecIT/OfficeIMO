using System;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public sealed partial class TableBlock {
    private IReadOnlyList<IMarkdownBlock>? TryParseStructuredCellBlocks(string? cell) {
        if (string.IsNullOrEmpty(cell)) {
            return null;
        }

        var normalized = NormalizeBreakMarkers(cell ?? string.Empty);
        if (!LooksLikeStructuredMarkdownCell(normalized)) {
            return null;
        }

        var options = InlineRenderOptions == null
            ? new MarkdownReaderOptions()
            : CloneOptionsWithoutTables(InlineRenderOptions);
        var state = InlineRenderState == null
            ? new MarkdownReaderState()
            : CloneState(InlineRenderState);
        var blocks = MarkdownReader.ParseBlockFragment(normalized, options, state);
        if (blocks.Count == 0) {
            return null;
        }

        if (ContainsUnsafeRawHtmlTableCellBlocks(blocks)) {
            return null;
        }

        if (blocks.Count == 1 && blocks[0] is ParagraphBlock) {
            return null;
        }

        return blocks;
    }

    internal static bool LooksLikeStructuredMarkdownCell(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var normalized = value!;
        if (normalized.IndexOf('\n') >= 0) {
            return true;
        }

        var trimmed = normalized.TrimStart();
        if (trimmed.Length == 0) {
            return false;
        }

        if (trimmed.StartsWith("```", StringComparison.Ordinal)
            || trimmed.StartsWith("~~~", StringComparison.Ordinal)
            || trimmed.StartsWith(">", StringComparison.Ordinal)
            || trimmed.StartsWith("<", StringComparison.Ordinal)) {
            return true;
        }

        if (trimmed[0] == '#') {
            int run = 1;
            while (run < trimmed.Length && trimmed[run] == '#') {
                run++;
            }

            if (run <= 6 && run < trimmed.Length && char.IsWhiteSpace(trimmed[run])) {
                return true;
            }
        }

        if (trimmed.Length >= 2
            && (trimmed[0] == '-' || trimmed[0] == '*' || trimmed[0] == '+')
            && char.IsWhiteSpace(trimmed[1])) {
            return true;
        }

        int digitIndex = 0;
        while (digitIndex < trimmed.Length && char.IsDigit(trimmed[digitIndex])) {
            digitIndex++;
        }

        if (digitIndex > 0
            && digitIndex + 1 < trimmed.Length
            && (trimmed[digitIndex] == '.' || trimmed[digitIndex] == ')')
            && char.IsWhiteSpace(trimmed[digitIndex + 1])) {
            return true;
        }

        return false;
    }

    internal static bool ContainsUnsafeRawHtmlTableCellBlocks(IReadOnlyList<IMarkdownBlock> blocks) {
        for (int i = 0; i < blocks.Count; i++) {
            if (ContainsUnsafeRawHtmlTableCellBlock(blocks[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsUnsafeRawHtmlTableCellBlock(IMarkdownBlock? block) {
        if (block == null) {
            return false;
        }

        if (block is HtmlRawBlock or HtmlCommentBlock) {
            return true;
        }

        if (block is not IChildMarkdownBlockContainer container || container.ChildBlocks.Count == 0) {
            return false;
        }

        for (int i = 0; i < container.ChildBlocks.Count; i++) {
            if (ContainsUnsafeRawHtmlTableCellBlock(container.ChildBlocks[i])) {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<string> PrepareStructuredRowMarkdown(
        IReadOnlyList<TableCell>? structuredRow,
        IReadOnlyList<string>? fallbackRow,
        int expectedCount) {
        if (structuredRow == null || structuredRow.Count == 0) {
            return PrepareRowCells(fallbackRow, expectedCount);
        }

        var cells = PrepareStructuredRowCells(structuredRow, expectedCount);
        var markdown = new string[cells.Count];
        for (int i = 0; i < cells.Count; i++) {
            markdown[i] = cells[i]?.Markdown ?? string.Empty;
        }
        return markdown;
    }

    private static List<TableCell> CloneStructuredRow(IReadOnlyList<TableCell> row) {
        var cloned = new List<TableCell>(row.Count);
        for (int i = 0; i < row.Count; i++) {
            cloned.Add(CloneStructuredCell(row[i]));
        }
        return cloned;
    }

    private static TableCell CloneStructuredCell(TableCell? cell) {
        if (cell == null) {
            return new TableCell();
        }

        return new TableCell(cell.Blocks) {
            IsHeader = cell.IsHeader,
            RowIndex = cell.RowIndex,
            ColumnIndex = cell.ColumnIndex,
            Alignment = cell.Alignment,
            BackgroundColor = cell.BackgroundColor,
            TextColor = cell.TextColor,
            Bold = cell.Bold,
            Italic = cell.Italic,
            Underline = cell.Underline,
            Strikethrough = cell.Strikethrough,
            ColumnSpan = cell.ColumnSpan,
            RowSpan = cell.RowSpan,
            SourceSpan = cell.SourceSpan,
            SyntaxChildren = cell.SyntaxChildren
        };
    }

    private static MarkdownReaderOptions CloneOptionsWithoutTables(MarkdownReaderOptions source) {
        var clone = MarkdownReaderOptions.CreateProfile(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO);
        clone.FrontMatter = false;
        clone.Callouts = source.Callouts;
        clone.Headings = source.Headings;
        clone.FencedCode = source.FencedCode;
        clone.IndentedCodeBlocks = source.IndentedCodeBlocks;
        clone.Images = source.Images;
        clone.UnorderedLists = source.UnorderedLists;
        clone.TaskLists = source.TaskLists;
        clone.OrderedLists = source.OrderedLists;
        clone.Tables = false;
        clone.DefinitionLists = source.DefinitionLists;
        clone.TocPlaceholders = source.TocPlaceholders;
        clone.Footnotes = source.Footnotes;
        clone.SingleTildeStrikethrough = source.SingleTildeStrikethrough;
        clone.Subscript = source.Subscript;
        clone.PreferNarrativeSingleLineDefinitions = source.PreferNarrativeSingleLineDefinitions;
        clone.HtmlBlocks = source.HtmlBlocks;
        clone.Paragraphs = source.Paragraphs;
        clone.AutolinkUrls = source.AutolinkUrls;
        clone.AutolinkAllowDomainWithoutPeriod = source.AutolinkAllowDomainWithoutPeriod;
        clone.AutolinkAllowQueryAndFragmentSpecialCharacters = source.AutolinkAllowQueryAndFragmentSpecialCharacters;
        clone.AutolinkAllowBalancedParenthesesWithTrailingPunctuation = source.AutolinkAllowBalancedParenthesesWithTrailingPunctuation;
        clone.AutolinkAllowTrailingPunctuationBeforeClosingParenthesis = source.AutolinkAllowTrailingPunctuationBeforeClosingParenthesis;
        clone.AutolinkTrimSingleTrailingPunctuationOrUnderscore = source.AutolinkTrimSingleTrailingPunctuationOrUnderscore;
        clone.AutolinkRequireLowercaseWwwPrefix = source.AutolinkRequireLowercaseWwwPrefix;
        clone.AutolinkRejectUnderscoreInWwwHost = source.AutolinkRejectUnderscoreInWwwHost;
        clone.AutolinkRequireLowercaseBareSchemePrefix = source.AutolinkRequireLowercaseBareSchemePrefix;
        clone.AutolinkBareMailtoDisplayAddressOnly = source.AutolinkBareMailtoDisplayAddressOnly;
        clone.AutolinkValidPreviousCharacters = source.AutolinkValidPreviousCharacters;
        clone.AutolinkBareSchemeUrls = source.AutolinkBareSchemeUrls;
        clone.AutolinkBareSchemePrefixes = source.AutolinkBareSchemePrefixes == null
            ? null
            : (string[])source.AutolinkBareSchemePrefixes.Clone();
        clone.AutolinkWwwUrls = source.AutolinkWwwUrls;
        clone.AutolinkWwwScheme = source.AutolinkWwwScheme;
        clone.AutolinkEmails = source.AutolinkEmails;
        clone.BackslashHardBreaks = source.BackslashHardBreaks;
        clone.SoftLineBreaksAsHardLineBreaks = source.SoftLineBreaksAsHardLineBreaks;
        clone.InlineHtml = source.InlineHtml;
        clone.BaseUri = source.BaseUri;
        clone.DisallowScriptUrls = source.DisallowScriptUrls;
        clone.DisallowFileUrls = source.DisallowFileUrls;
        clone.AllowMailtoUrls = source.AllowMailtoUrls;
        clone.AllowDataUrls = source.AllowDataUrls;
        clone.AllowProtocolRelativeUrls = source.AllowProtocolRelativeUrls;
        clone.RestrictUrlSchemes = source.RestrictUrlSchemes;
        clone.AllowedUrlSchemes = source.AllowedUrlSchemes;
        clone.MaxInputCharacters = source.MaxInputCharacters;
        clone.InputNormalization = source.InputNormalization == null
            ? new MarkdownInputNormalizationOptions()
            : source.InputNormalization;
        clone.FencedBlockExtensions.Clear();
        for (int i = 0; i < source.FencedBlockExtensions.Count; i++) {
            if (source.FencedBlockExtensions[i] != null) {
                clone.FencedBlockExtensions.Add(source.FencedBlockExtensions[i]);
            }
        }

        clone.BlockParserExtensions.Clear();
        for (int i = 0; i < source.BlockParserExtensions.Count; i++) {
            if (source.BlockParserExtensions[i] != null) {
                clone.BlockParserExtensions.Add(source.BlockParserExtensions[i]);
            }
        }

        clone.InlineParserExtensions.Clear();
        for (int i = 0; i < source.InlineParserExtensions.Count; i++) {
            if (source.InlineParserExtensions[i] != null) {
                clone.InlineParserExtensions.Add(source.InlineParserExtensions[i]);
            }
        }

        for (int i = 0; i < source.DocumentTransforms.Count; i++) {
            if (source.DocumentTransforms[i] != null) {
                clone.DocumentTransforms.Add(source.DocumentTransforms[i]);
            }
        }

        return clone;
    }

    private static MarkdownReaderState CloneState(MarkdownReaderState state) {
        var clone = new MarkdownReaderState();
        foreach (var kvp in state.LinkRefs) {
            clone.LinkRefs[kvp.Key] = kvp.Value;
        }

        clone.SourceLineOffset = state.SourceLineOffset;
        clone.SourceTextMap = state.SourceTextMap;
        return clone;
    }
}
