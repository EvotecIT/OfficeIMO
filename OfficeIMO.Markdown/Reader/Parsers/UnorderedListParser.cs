namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class UnorderedListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.UnorderedLists) return false;
            if (!IsUnorderedListLine(lines[i], out int level0Abs, out var isTask, out var done, out var firstContent)) return false;
            if (isTask && !options.TaskLists) {
                isTask = false;
                done = false;
                firstContent = GetUnorderedListItemContent(lines[i]);
            }
            var ul = new UnorderedListBlock();
            int firstContinuationIndent = GetListContinuationIndent(lines[i]);

            int j = i + 1;
            var firstSourceLines = new List<MarkdownSourceLineSlice>();
            var firstLines = ConsumeListContinuationLines(
                lines,
                ref j,
                firstContinuationIndent,
                firstContent,
                options,
                breakOnAnyOrderedListLine: false,
                sourceLines: firstSourceLines,
                absoluteLineOffset: state.SourceLineOffset,
                initialLineIndex: i,
                initialStartColumn: GetListLeadContentStartColumn(lines[i], isTask));
            var first = CreateListItemFromLeadLines(firstLines, isTask, done, options, state, firstSourceLines);
            first.Level = 0;
            AddListItemLeadSyntaxNodes(first, firstLines, i, options, state, firstSourceLines);
            ul.Items.Add(first);

            // Allow both same-type and mixed nested lists under the current item.
            ConsumeNestedBlocksForListItem(lines, ref j, level0Abs, firstContinuationIndent, options, state, first, allowNestedOrdered: true, allowNestedUnordered: true);

            while (j < lines.Length && IsUnorderedListLine(lines[j], out var lvlAbs, out var isTask2, out var done2, out var content2) && lvlAbs >= level0Abs) {
                if (isTask2 && !options.TaskLists) {
                    isTask2 = false;
                    done2 = false;
                    content2 = GetUnorderedListItemContent(lines[j]);
                }
                int continuationIndent = GetListContinuationIndent(lines[j]);
                int next = j + 1;
                var itemSourceLines = new List<MarkdownSourceLineSlice>();
                var itemLines = ConsumeListContinuationLines(
                    lines,
                    ref next,
                    continuationIndent,
                    content2,
                    options,
                    breakOnAnyOrderedListLine: false,
                    sourceLines: itemSourceLines,
                    absoluteLineOffset: state.SourceLineOffset,
                    initialLineIndex: j,
                    initialStartColumn: GetListLeadContentStartColumn(lines[j], isTask2));
                var li = CreateListItemFromLeadLines(itemLines, isTask2, done2, options, state, itemSourceLines);
                li.Level = lvlAbs - level0Abs;
                AddListItemLeadSyntaxNodes(li, itemLines, j, options, state, itemSourceLines);
                ul.Items.Add(li);
                j = next;

                ConsumeNestedBlocksForListItem(lines, ref j, lvlAbs, continuationIndent, options, state, li, allowNestedOrdered: true, allowNestedUnordered: true);
            }
            doc.Add(ul); i = j; return true;
        }
    }
}
