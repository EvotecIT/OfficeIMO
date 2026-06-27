namespace OfficeIMO.Markdown;

/// <summary>
/// Writes markdown from parse results while preserving original source bytes when that is provably safe.
/// </summary>
public static class MarkdownRoundtripWriter {
    private const string PreserveTriviaRequiredId = "roundtrip.preserve-trivia-required";
    private const string DocumentTransformedId = "roundtrip.document-transformed";
    private const string OriginalSourceSliceUnavailableId = "roundtrip.original-source-slice-unavailable";
    private const string OverlappingEditsId = "roundtrip.overlapping-edits";

    /// <summary>
    /// Returns the captured original markdown for unchanged parse results, otherwise falls back to semantic markdown generation.
    /// </summary>
    /// <remarks>
    /// This is an intentionally conservative first lossless primitive. It preserves bytes only when the parse
    /// result retained the original input and no document transforms reported changes.
    /// </remarks>
    public static MarkdownRoundtripResult WriteUnchanged(MarkdownParseResult result) {
        if (result == null) {
            throw new ArgumentNullException(nameof(result));
        }

        if (!result.PreservesOriginalMarkdown) {
            return GeneratedWithDiagnostic(
                result,
                PreserveTriviaRequiredId,
                "The parse result does not contain original reader input. Parse with PreserveTrivia enabled before requesting a lossless unchanged roundtrip.");
        }

        if (result.TransformDiagnostics.Count > 0) {
            return GeneratedWithDiagnostic(
                result,
                DocumentTransformedId,
                "The parsed document was changed by one or more document transforms. Generated markdown was emitted instead of claiming a byte-preserving roundtrip.");
        }

        return new MarkdownRoundtripResult(result.OriginalMarkdown);
    }

    /// <summary>
    /// Applies one source edit and preserves original source bytes around it when the parse result can be mapped safely.
    /// </summary>
    public static MarkdownRoundtripResult WriteWithSourceEdit(MarkdownParseResult result, MarkdownNativeSourceEdit edit) {
        if (edit == null) {
            throw new ArgumentNullException(nameof(edit));
        }

        return WriteWithSourceEdits(result, new[] { edit });
    }

    /// <summary>
    /// Applies source edits and preserves original source bytes around them when the parse result can be mapped safely.
    /// </summary>
    public static MarkdownRoundtripResult WriteWithSourceEdits(
        MarkdownParseResult result,
        IEnumerable<MarkdownNativeSourceEdit> edits) {
        if (result == null) {
            throw new ArgumentNullException(nameof(result));
        }

        if (edits == null) {
            throw new ArgumentNullException(nameof(edits));
        }

        var editList = edits.Where(edit => edit != null).ToArray();
        if (editList.Length == 0) {
            return WriteUnchanged(result);
        }

        var diagnostics = new List<MarkdownRoundtripDiagnostic>();
        if (result.TransformDiagnostics.Count > 0) {
            diagnostics.Add(new MarkdownRoundtripDiagnostic(
                DocumentTransformedId,
                "The parsed document was changed by one or more document transforms. Source edits were applied to normalized markdown instead of claiming a byte-preserving original-source roundtrip."));
        } else if (!result.PreservesOriginalMarkdown) {
            diagnostics.Add(new MarkdownRoundtripDiagnostic(
                PreserveTriviaRequiredId,
                "The parse result does not contain original reader input. Source edits were applied to normalized markdown. Parse with PreserveTrivia enabled before requesting a lossless source-edit roundtrip."));
        } else if (TryCreateOriginalReplacements(result, editList, diagnostics, out var originalReplacements)) {
            return ApplyReplacements(result.OriginalMarkdown, originalReplacements, diagnostics);
        }

        var normalizedReplacements = editList
            .Select(edit => new SourceReplacement(edit.StartOffset, edit.EndOffsetInclusive, edit.ReplacementMarkdown, edit.SourceSpan))
            .ToArray();
        return ApplyReplacements(result.SourceMarkdown, normalizedReplacements, diagnostics);
    }

    private static MarkdownRoundtripResult GeneratedWithDiagnostic(
        MarkdownParseResult result,
        string id,
        string message) {
        return new MarkdownRoundtripResult(
            result.Document.ToMarkdown(),
            new[] { new MarkdownRoundtripDiagnostic(id, message) });
    }

    private static bool TryCreateOriginalReplacements(
        MarkdownParseResult result,
        IReadOnlyList<MarkdownNativeSourceEdit> edits,
        ICollection<MarkdownRoundtripDiagnostic> diagnostics,
        out SourceReplacement[] replacements) {
        var replacementList = new List<SourceReplacement>(edits.Count);
        for (var i = 0; i < edits.Count; i++) {
            var edit = edits[i];
            if (!result.TryCreateOriginalSourceSlice(edit.SourceSpan, out var slice)) {
                diagnostics.Add(new MarkdownRoundtripDiagnostic(
                    OriginalSourceSliceUnavailableId,
                    "At least one source edit could not be mapped back to original reader input. Source edits were applied to normalized markdown instead.",
                    edit.SourceSpan));
                replacements = Array.Empty<SourceReplacement>();
                return false;
            }

            replacementList.Add(new SourceReplacement(slice.StartOffset, slice.EndOffsetInclusive, edit.ReplacementMarkdown, edit.SourceSpan));
        }

        replacements = replacementList.ToArray();
        return true;
    }

    private static MarkdownRoundtripResult ApplyReplacements(
        string sourceMarkdown,
        IReadOnlyList<SourceReplacement> replacements,
        ICollection<MarkdownRoundtripDiagnostic> diagnostics) {
        if (!TryOrderReplacements(replacements, diagnostics, out var orderedReplacements)) {
            return new MarkdownRoundtripResult(sourceMarkdown ?? string.Empty, diagnostics.ToArray());
        }

        sourceMarkdown ??= string.Empty;
        var markdown = sourceMarkdown;
        for (var i = 0; i < orderedReplacements.Count; i++) {
            var replacement = orderedReplacements[i];
            var startOffset = Math.Max(0, Math.Min(markdown.Length, replacement.StartOffset));
            var endExclusive = Math.Max(startOffset, Math.Min(markdown.Length, replacement.EndOffsetInclusive + 1));
            markdown = markdown.Substring(0, startOffset)
                       + replacement.ReplacementMarkdown
                       + markdown.Substring(endExclusive);
        }

        return new MarkdownRoundtripResult(markdown, diagnostics.ToArray());
    }

    private static bool TryOrderReplacements(
        IReadOnlyList<SourceReplacement> replacements,
        ICollection<MarkdownRoundtripDiagnostic> diagnostics,
        out IReadOnlyList<SourceReplacement> orderedReplacements) {
        var ascending = replacements
            .OrderBy(replacement => replacement.StartOffset)
            .ThenBy(replacement => replacement.EndOffsetInclusive)
            .ToArray();

        for (var i = 1; i < ascending.Length; i++) {
            if (ascending[i].StartOffset <= ascending[i - 1].EndOffsetInclusive) {
                diagnostics.Add(new MarkdownRoundtripDiagnostic(
                    OverlappingEditsId,
                    "Source edits overlap and cannot be applied deterministically.",
                    ascending[i].SourceSpan));
                orderedReplacements = Array.Empty<SourceReplacement>();
                return false;
            }
        }

        orderedReplacements = ascending
            .OrderByDescending(replacement => replacement.StartOffset)
            .ToArray();
        return true;
    }

    private readonly struct SourceReplacement {
        public SourceReplacement(
            int startOffset,
            int endOffsetInclusive,
            string replacementMarkdown,
            MarkdownSourceSpan? sourceSpan) {
            StartOffset = startOffset;
            EndOffsetInclusive = endOffsetInclusive;
            ReplacementMarkdown = replacementMarkdown ?? string.Empty;
            SourceSpan = sourceSpan;
        }

        public int StartOffset { get; }

        public int EndOffsetInclusive { get; }

        public string ReplacementMarkdown { get; }

        public MarkdownSourceSpan? SourceSpan { get; }
    }
}
