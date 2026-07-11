using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Reader;

public static partial class ReaderHierarchicalChunker {
    private static void SplitSourceChunk(ReaderChunk source, int inputIndex, ChunkingState state) {
        ReaderHierarchicalChunkingOptions options = state.Options;
        IReaderTokenCounter counter = options.TokenCounter;
        bool markdownSelected = options.PreferMarkdown && !string.IsNullOrWhiteSpace(source.Markdown);
        string content = markdownSelected ? source.Markdown! : source.Text ?? string.Empty;
        string? context = BuildContext(source.Location, state);
        string prefix = BuildContextPrefix(context, content, state);
        state.SourceTokenCount += CountTokens(counter, content);

        if (content.Length == 0) {
            EmitSegment(source, inputIndex, markdownSelected, content, prefix, context, 0, 0, 0, 0, state);
            return;
        }

        int start = 0;
        int previousEnd = 0;
        int segmentIndex = 0;
        while (start < content.Length) {
            state.CancellationToken.ThrowIfCancellationRequested();
            if (state.Chunks.Count >= options.MaxOutputChunks) {
                state.OutputLimitReached = true;
                state.AddLimitDiagnostic("hierarchical-output-chunk-limit", options.MaxOutputChunks, "output chunks");
                return;
            }

            int end = FindSegmentEnd(content, start, prefix, options.MaxTokens, counter);
            int overlapCharacters = segmentIndex == 0 ? 0 : Math.Max(0, previousEnd - start);
            int overlapTokens = overlapCharacters == 0
                ? 0
                : CountTokens(counter, content.Substring(start, overlapCharacters));
            EmitSegment(
                source,
                inputIndex,
                markdownSelected,
                content,
                prefix,
                context,
                segmentIndex,
                start,
                end,
                overlapTokens,
                state);
            state.Segments[state.Segments.Count - 1].OverlapCharacterCount = overlapCharacters;

            if (end >= content.Length) return;
            previousEnd = end;
            start = FindNextStart(content, start, end, options.OverlapTokens, counter);
            segmentIndex++;
        }
    }

    private static string BuildContextPrefix(string? context, string content, ChunkingState state) {
        if (!state.Options.IncludeContextInText || string.IsNullOrWhiteSpace(context)) return string.Empty;
        string prefix = context + "\n\n";
        int prefixTokens = CountTokens(state.Options.TokenCounter, prefix);
        if (content.Length > 0 && prefixTokens >= state.Options.MaxTokens) {
            state.AddDiagnostic(
                "hierarchical-context-omitted",
                "Hierarchy context exceeded the chunk token budget and was retained only as segment metadata.");
            return string.Empty;
        }
        if (content.Length > 0 &&
            CountOutputTokens(
                state.Options.TokenCounter,
                prefix,
                content,
                0,
                NextCharacterBoundary(content, 0)) > state.Options.MaxTokens) {
            state.AddDiagnostic(
                "hierarchical-context-omitted",
                "Hierarchy context left no room for source content and was retained only as segment metadata.");
            return string.Empty;
        }
        if (content.Length == 0 && prefixTokens > state.Options.MaxTokens) {
            state.AddDiagnostic(
                "hierarchical-context-omitted",
                "Hierarchy context exceeded the chunk token budget and was retained only as segment metadata.");
            return string.Empty;
        }
        return prefix;
    }

    private static int FindSegmentEnd(
        string content,
        int start,
        string prefix,
        int maxTokens,
        IReaderTokenCounter counter) {
        if (CountOutputTokens(counter, prefix, content, start, content.Length) <= maxTokens) return content.Length;

        int firstEnd = NextCharacterBoundary(content, start);
        if (CountOutputTokens(counter, prefix, content, start, firstEnd) > maxTokens) {
            throw new InvalidOperationException(
                $"Token counter '{counter.Id}' cannot fit one source character within MaxTokens={maxTokens.ToString(CultureInfo.InvariantCulture)}.");
        }

        int low = firstEnd;
        int high = content.Length;
        int best = firstEnd;
        while (low <= high) {
            int middle = low + ((high - low) / 2);
            middle = ClampToCharacterBoundary(content, middle, low, high);
            int count = CountOutputTokens(counter, prefix, content, start, middle);
            if (count <= maxTokens) {
                best = middle;
                if (middle >= content.Length) break;
                low = NextCharacterBoundary(content, middle);
            } else {
                high = PreviousCharacterBoundary(content, middle);
            }
        }

        int preferred = FindPreferredEnd(content, start, best);
        if (preferred > start && CountOutputTokens(counter, prefix, content, start, preferred) <= maxTokens) {
            best = preferred;
        }
        while (best > start && CountOutputTokens(counter, prefix, content, start, best) > maxTokens) {
            best = PreviousCharacterBoundary(content, best);
        }
        if (best <= start) throw new InvalidOperationException("Token-aware splitting could not make forward progress.");
        return best;
    }

    private static int FindNextStart(
        string content,
        int segmentStart,
        int segmentEnd,
        int overlapTokens,
        IReaderTokenCounter counter) {
        if (overlapTokens <= 0) return segmentEnd;
        int minimum = NextCharacterBoundary(content, segmentStart);
        int low = minimum;
        int high = segmentEnd;
        int best = segmentEnd;
        while (low <= high) {
            int middle = ClampToCharacterBoundary(content, low + ((high - low) / 2), minimum, segmentEnd);
            int count = CountTokens(counter, content.Substring(middle, segmentEnd - middle));
            if (count <= overlapTokens) {
                best = middle;
                high = PreviousCharacterBoundary(content, middle);
            } else {
                if (middle >= segmentEnd) break;
                low = NextCharacterBoundary(content, middle);
            }
        }

        int preferred = FindPreferredStart(content, best, segmentEnd);
        return preferred > segmentStart && preferred <= segmentEnd ? preferred : best;
    }

    private static int FindPreferredEnd(string text, int start, int maximumEnd) {
        int minimum = start + Math.Max(1, (maximumEnd - start) / 2);
        int boundary = FindBackward(text, minimum, maximumEnd, character => character == '\n');
        if (boundary > start) return boundary;
        boundary = FindBackward(text, minimum, maximumEnd, IsSentenceBoundary);
        if (boundary > start) return boundary;
        boundary = FindBackward(text, minimum, maximumEnd, char.IsWhiteSpace);
        return boundary > start ? boundary : maximumEnd;
    }

    private static int FindPreferredStart(string text, int minimumStart, int end) {
        int searchEnd = Math.Min(end, minimumStart + Math.Max(8, (end - minimumStart) / 3));
        for (int index = minimumStart; index < searchEnd; index++) {
            if (char.IsWhiteSpace(text[index])) return index + 1;
        }
        return minimumStart;
    }

    private static int FindBackward(string text, int minimum, int maximum, Func<char, bool> predicate) {
        for (int index = maximum - 1; index >= minimum; index--) {
            if (predicate(text[index])) return index + 1;
        }
        return -1;
    }

    private static bool IsSentenceBoundary(char character) =>
        character == '.' || character == '!' || character == '?' || character == ';';

    private static int NextCharacterBoundary(string text, int index) {
        if (index >= text.Length) return text.Length;
        return char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])
            ? index + 2
            : index + 1;
    }

    private static int PreviousCharacterBoundary(string text, int index) {
        if (index <= 0) return 0;
        int previous = index - 1;
        return previous > 0 && char.IsLowSurrogate(text[previous]) && char.IsHighSurrogate(text[previous - 1])
            ? previous - 1
            : previous;
    }

    private static int ClampToCharacterBoundary(string text, int index, int minimum, int maximum) {
        int clamped = Math.Max(minimum, Math.Min(maximum, index));
        if (clamped > minimum && clamped < text.Length && char.IsLowSurrogate(text[clamped]) && char.IsHighSurrogate(text[clamped - 1])) {
            clamped--;
        } else if (clamped < maximum && clamped < text.Length && char.IsLowSurrogate(text[clamped]) && clamped > 0 && char.IsHighSurrogate(text[clamped - 1])) {
            clamped++;
        }
        return Math.Max(minimum, clamped);
    }

    private static void EmitSegment(
        ReaderChunk source,
        int inputIndex,
        bool markdownSelected,
        string sourceContent,
        string prefix,
        string? context,
        int segmentIndex,
        int start,
        int end,
        int overlapTokens,
        ChunkingState state) {
        if (state.Chunks.Count >= state.Options.MaxOutputChunks) {
            state.OutputLimitReached = true;
            state.AddLimitDiagnostic("hierarchical-output-chunk-limit", state.Options.MaxOutputChunks, "output chunks");
            return;
        }

        string segmentContent = sourceContent.Substring(start, end - start);
        string output = prefix + segmentContent;
        int outputTokens = CountTokens(state.Options.TokenCounter, output);
        if (outputTokens > state.Options.MaxTokens) {
            throw new InvalidOperationException("Token-aware splitting emitted a chunk beyond the configured token budget.");
        }
        int contentTokens = CountTokens(state.Options.TokenCounter, segmentContent);
        int contextTokens = prefix.Length == 0 ? 0 : Math.Max(0, outputTokens - contentTokens);
        string chunkId = BuildSegmentId(source, inputIndex, segmentIndex, start, end);
        var chunk = new ReaderChunk {
            Id = chunkId,
            Kind = source.Kind,
            Location = CloneLocation(source.Location),
            SourceId = source.SourceId,
            SourceHash = source.SourceHash,
            ChunkHash = ComputeSegmentHash(chunkId, output),
            SourceLastWriteUtc = source.SourceLastWriteUtc,
            SourceLengthBytes = source.SourceLengthBytes,
            TokenEstimate = outputTokens,
            Text = output,
            Markdown = markdownSelected ? output : null,
            Tables = segmentIndex == 0 ? source.Tables : null,
            Visuals = segmentIndex == 0 ? source.Visuals : null,
            FormFields = segmentIndex == 0 ? source.FormFields : null,
            Actions = segmentIndex == 0 ? source.Actions : null,
            Diagnostics = segmentIndex == 0 ? source.Diagnostics : null,
            Warnings = segmentIndex == 0 ? source.Warnings : null
        };
        state.Chunks.Add(chunk);
        state.Segments.Add(new ReaderChunkSegment {
            ChunkId = chunkId,
            SourceChunkId = source.Id ?? string.Empty,
            SegmentIndex = segmentIndex,
            StartCharacter = start,
            EndCharacter = end,
            OverlapTokenCount = overlapTokens,
            ContentTokenCount = contentTokens,
            ContextTokenCount = contextTokens,
            TokenCount = outputTokens,
            Context = context
        });
        state.OutputTokenCount += outputTokens;
        state.OverlapTokenCount += overlapTokens;
        state.ContextTokenCount += contextTokens;
    }

    private static int CountOutputTokens(
        IReaderTokenCounter counter,
        string prefix,
        string content,
        int start,
        int end) {
        return CountTokens(counter, prefix + content.Substring(start, end - start));
    }

    private static int CountTokens(IReaderTokenCounter counter, string value) {
        int count = counter.CountTokens(value ?? string.Empty);
        if (count < 0) {
            throw new InvalidOperationException($"Token counter '{counter.Id}' returned a negative count.");
        }
        return count;
    }

    private static string? BuildContext(ReaderLocation? location, ChunkingState state) {
        if (location == null) return null;
        var builder = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(location.Sheet)) AppendContext(builder, "Sheet: " + location.Sheet!.Trim());
        else if (location.Slide.HasValue) AppendContext(builder, "Slide " + location.Slide.Value.ToString(CultureInfo.InvariantCulture));
        else if (location.Page.HasValue) AppendContext(builder, "Page " + location.Page.Value.ToString(CultureInfo.InvariantCulture));
        string? hierarchyPath = ReaderHeadingPath.GetValidatedHierarchyPath(location);
        string? headingDisplay = hierarchyPath == null
            ? location.HeadingPath
            : ReaderHeadingPath.ToDisplayString(hierarchyPath);
        if (!string.IsNullOrWhiteSpace(headingDisplay)) AppendContext(builder, headingDisplay!);
        if (builder.Length == 0) return null;
        string context = builder.ToString();
        if (context.Length <= state.Options.MaxContextCharacters) return context;
        state.AddLimitDiagnostic(
            "hierarchical-context-character-limit",
            state.Options.MaxContextCharacters,
            "context characters");
        return TruncateAtCharacterBoundary(context, state.Options.MaxContextCharacters);
    }

    private static void AppendContext(StringBuilder builder, string value) {
        if (builder.Length > 0) builder.Append(" > ");
        builder.Append(value);
    }

    private static ReaderLocation CloneLocation(ReaderLocation? source) {
        if (source == null) return new ReaderLocation();
        return new ReaderLocation {
            Path = source.Path,
            BlockIndex = source.BlockIndex,
            SourceBlockIndex = source.SourceBlockIndex,
            StartLine = source.StartLine,
            EndLine = source.EndLine,
            NormalizedStartLine = source.NormalizedStartLine,
            NormalizedEndLine = source.NormalizedEndLine,
            HeadingPath = source.HeadingPath,
            HierarchyHeadingPath = source.HierarchyHeadingPath,
            HierarchyHeadingDisplayPath = source.HierarchyHeadingDisplayPath,
            HeadingSlug = source.HeadingSlug,
            SourceBlockKind = source.SourceBlockKind,
            BlockAnchor = source.BlockAnchor,
            Sheet = source.Sheet,
            A1Range = source.A1Range,
            Slide = source.Slide,
            Page = source.Page,
            TableIndex = source.TableIndex
        };
    }

    private static string BuildSegmentId(ReaderChunk source, int inputIndex, int segmentIndex, int start, int end) {
        var identity = new StringBuilder();
        AppendSegmentIdentity(identity, GetSourceIdentity(source));
        AppendSegmentIdentity(identity, source.Id);
        AppendSegmentIdentity(identity, inputIndex.ToString(CultureInfo.InvariantCulture));
        AppendSegmentIdentity(identity, segmentIndex.ToString(CultureInfo.InvariantCulture));
        AppendSegmentIdentity(identity, start.ToString(CultureInfo.InvariantCulture));
        AppendSegmentIdentity(identity, end.ToString(CultureInfo.InvariantCulture));
        return "rag:" + ComputeSha256Hex(identity.ToString());
    }

    private static void AppendSegmentIdentity(StringBuilder builder, string? value) {
        if (value == null) {
            builder.Append("-1:");
            return;
        }
        builder.Append(value.Length.ToString(CultureInfo.InvariantCulture));
        builder.Append(':');
        builder.Append(value);
    }

    private static string ComputeSegmentHash(string chunkId, string content) =>
        ComputeSha256Hex(chunkId + "|" + content);

    private static string ComputeSha256Hex(string value) {
        using SHA256 sha = SHA256.Create();
        byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
        var builder = new StringBuilder(hash.Length * 2);
        for (int index = 0; index < hash.Length; index++) builder.Append(hash[index].ToString("x2", CultureInfo.InvariantCulture));
        return builder.ToString();
    }
}
