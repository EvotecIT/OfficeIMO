namespace OfficeIMO.Latex.Markdown;

internal static partial class LatexToMarkdownConverter {
    // A block inside a command argument belongs to that enclosing command. If
    // native block ranges fragment its syntax, preserve the complete command
    // before projecting any child block. Paragraphs on either side stay visible.
    private static BlockCandidate[] WithCommandFallbacks(LatexProjectionContext context, BlockCandidate[] candidates) {
        if (context.Document.Body == null || context.Document.Profile == LatexDocumentProfile.PreserveOnly) return candidates;
        var dominant = new List<BlockCandidate>();
        int end = context.Document.Body.ContentSpan.Start.Offset;
        foreach (BlockCandidate candidate in candidates) {
            context.CheckCancellation();
            if (candidate.Span.Start.Offset < end) continue;
            dominant.Add(candidate);
            end = candidate.Span.End.Offset;
        }
        var fallbacks = new List<BlockCandidate>();
        int blockIndex = 0, fallbackEnd = 0;
        int bodyStart = context.Document.Body.ContentSpan.Start.Offset;
        int bodyEnd = context.Document.Body.ContentSpan.End.Offset;
        foreach (LatexInlineCandidate inline in context.InlineCandidates(bodyStart, bodyEnd)) {
            context.CheckCancellation();
            if (inline.Command is not LatexCommand command) continue;
            LatexSourceSpan span = command.Syntax.Span;
            if (!IsInside(span, bodyStart, bodyEnd) || span.Start.Offset < fallbackEnd || !LatexSemanticBuilder.IsActiveSyntax(command.Syntax)) continue;
            while (blockIndex < dominant.Count && dominant[blockIndex].Span.End.Offset <= span.Start.Offset) blockIndex++;
            if (blockIndex >= dominant.Count) break;
            LatexSourceSpan block = dominant[blockIndex].Span;
            if (block.Start.Offset >= span.End.Offset || IsInside(span, block.Start.Offset, block.End.Offset)) continue;
            fallbacks.Add(new BlockCandidate(span, command));
            fallbackEnd = span.End.Offset;
        }
        if (fallbacks.Count == 0) return dominant.ToArray();
        var result = new List<BlockCandidate>(dominant.Count + fallbacks.Count);
        int fallbackIndex = 0;
        foreach (BlockCandidate candidate in dominant) {
            context.CheckCancellation();
            while (fallbackIndex < fallbacks.Count && fallbacks[fallbackIndex].Span.End.Offset <= candidate.Span.Start.Offset) fallbackIndex++;
            if (candidate.Value is LatexParagraph) {
                int cursor = candidate.Span.Start.Offset;
                for (int index = fallbackIndex; index < fallbacks.Count && fallbacks[index].Span.Start.Offset < candidate.Span.End.Offset; index++) {
                    context.CheckCancellation();
                    if (fallbacks[index].Span.Start.Offset > cursor)
                        result.Add(new BlockCandidate(context.Document.Source.CreateSpan(cursor, fallbacks[index].Span.Start.Offset), candidate.Value));
                    cursor = Math.Max(cursor, fallbacks[index].Span.End.Offset);
                }
                if (cursor < candidate.Span.End.Offset) result.Add(new BlockCandidate(context.Document.Source.CreateSpan(cursor, candidate.Span.End.Offset), candidate.Value));
            } else if (fallbackIndex >= fallbacks.Count || !IsInside(candidate.Span, fallbacks[fallbackIndex].Span.Start.Offset, fallbacks[fallbackIndex].Span.End.Offset)) {
                result.Add(candidate);
            }
        }
        result.AddRange(fallbacks);
        return result.OrderBy(static candidate => candidate.Span.Start.Offset).ThenByDescending(static candidate => candidate.Span.Length).ToArray();
    }
}
