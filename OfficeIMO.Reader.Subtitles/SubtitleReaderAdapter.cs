namespace OfficeIMO.Reader.Subtitles;

internal static class SubtitleReaderAdapter {
    internal static OfficeDocumentReadResult ReadDocument(
        string path,
        ReaderOptions readerOptions,
        ReaderSubtitleOptions subtitleOptions,
        CancellationToken cancellationToken) {
        ReaderAdapterInputSnapshot input = DocumentReaderEngine.ReadAdapterInput(path, readerOptions, cancellationToken);
        return BuildDocument(input, readerOptions, subtitleOptions, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(
        Stream stream,
        string? sourceName,
        ReaderOptions readerOptions,
        ReaderSubtitleOptions subtitleOptions,
        CancellationToken cancellationToken) {
        ReaderAdapterInputSnapshot input = DocumentReaderEngine.ReadAdapterInput(
            stream,
            string.IsNullOrWhiteSpace(sourceName) ? "subtitles.srt" : sourceName,
            readerOptions,
            cancellationToken);
        return BuildDocument(input, readerOptions, subtitleOptions, cancellationToken);
    }

    private static OfficeDocumentReadResult BuildDocument(
        ReaderAdapterInputSnapshot input,
        ReaderOptions readerOptions,
        ReaderSubtitleOptions subtitleOptions,
        CancellationToken cancellationToken) {
        string content = DecodeText(input.Bytes);
        SubtitleParseResult parsed = SubtitleParser.Parse(content, input.Source.Path, subtitleOptions, cancellationToken);
        var chunks = new List<ReaderChunk>(parsed.Cues.Count);
        for (int index = 0; index < parsed.Cues.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            SubtitleCue cue = parsed.Cues[index];
            string baseAnchor = "subtitle-cue-" + index.ToString("D6", CultureInfo.InvariantCulture);
            string markdownPrefix = subtitleOptions.IncludeTimestampsInMarkdown
                ? "**" + SubtitleParser.FormatTimestamp(cue.Start) + " → " + SubtitleParser.FormatTimestamp(cue.End) + "**" +
                    Environment.NewLine + Environment.NewLine
                : string.Empty;
            IReadOnlyList<SubtitleProjectionPart> cueParts = SplitCueProjection(
                cue.Text, readerOptions.MaxChars, markdownPrefix.Length);
            string[]? warnings = BuildCueWarnings(cue.Truncated, cueParts.Count > 1);
            for (int partIndex = 0; partIndex < cueParts.Count; partIndex++) {
                string anchor = partIndex == 0
                    ? baseAnchor
                    : baseAnchor + "-part-" + partIndex.ToString("D4", CultureInfo.InvariantCulture);
                var location = new ReaderLocation {
                    Path = input.Source.Path,
                    BlockIndex = chunks.Count,
                    SourceBlockIndex = index,
                    StartLine = cue.StartLine,
                    EndLine = cue.EndLine,
                    SourceBlockKind = "subtitle-cue",
                    BlockAnchor = anchor
                };
                SubtitleProjectionPart part = cueParts[partIndex];
                var chunk = new ReaderChunk {
                    Id = anchor,
                    Kind = ReaderInputKind.Text,
                    ContinuesPreviousChunk = partIndex > 0,
                    Location = location,
                    Text = part.Text,
                    Markdown = partIndex == 0 ? markdownPrefix + part.Markdown : part.Markdown,
                    Warnings = warnings
                };
                DocumentReaderEngine.ApplyAdapterSource(chunk, input, readerOptions.ComputeHashes);
                chunks.Add(chunk);
            }
        }

        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Text,
            input.Source,
            new[] { OfficeDocumentReaderBuilderSubtitleExtensions.HandlerId, "subtitle." + parsed.Format });
        result.Source.Title = Path.GetFileName(input.Source.Path ?? "subtitles.srt");
        result.Metadata = result.Metadata.Concat(BuildMetadata(parsed, input.Source.Path)).ToArray();
        if (parsed.Warnings.Count > 0) {
            result.Diagnostics = result.Diagnostics.Concat(parsed.Warnings.Select((warning, index) =>
                new OfficeDocumentDiagnostic {
                    Severity = OfficeDocumentDiagnosticSeverity.Warning,
                    Category = warning.Contains("MaxCues", StringComparison.Ordinal)
                        ? OfficeDocumentDiagnosticCategory.Limit
                        : OfficeDocumentDiagnosticCategory.Parsing,
                    Code = warning.Contains("MaxCues", StringComparison.Ordinal)
                        ? "subtitle-cue-limit"
                        : "subtitle-invalid-block",
                    Message = warning,
                    Source = OfficeDocumentReaderBuilderSubtitleExtensions.HandlerId,
                    IsRecoverable = true
                })).ToArray();
        }
        return result;
    }

    private static IReadOnlyList<SubtitleProjectionPart> SplitCueProjection(
        string value,
        int maxChars,
        int firstPartPrefixLength) {
        int limit = Math.Max(1, maxChars);
        int firstMarkdownLimit = Math.Max(1, limit - firstPartPrefixLength);
        var result = new List<SubtitleProjectionPart>();
        int offset = 0;
        int markdownLimit = firstMarkdownLimit;
        while (offset < value.Length) {
            int length = FindAlignedCuePartLength(value, offset, limit, markdownLimit);
            string text = value.Substring(offset, length);
            result.Add(new SubtitleProjectionPart(text, EscapeCueMarkdown(text)));
            offset += length;
            markdownLimit = limit;
        }
        return result;
    }

    private static int FindAlignedCuePartLength(
        string value,
        int offset,
        int textLimit,
        int markdownLimit) {
        int length = 0;
        int markdownLength = 0;
        while (offset + length < value.Length) {
            int scalarLength = char.IsHighSurrogate(value[offset + length]) &&
                offset + length + 1 < value.Length && char.IsLowSurrogate(value[offset + length + 1])
                ? 2
                : 1;
            int escapedLength = scalarLength == 1
                ? EscapedCueCharacterLength(value[offset + length])
                : scalarLength;
            if (length > 0 &&
                (length + scalarLength > textLimit || markdownLength + escapedLength > markdownLimit)) {
                break;
            }
            length += scalarLength;
            markdownLength += escapedLength;
            if (length >= textLimit || markdownLength >= markdownLimit) break;
        }
        return Math.Max(1, length);
    }

    private static int EscapedCueCharacterLength(char value) {
        switch (value) {
            case '&': return 5;
            case '<':
            case '>': return 4;
            default: return 1;
        }
    }

    private static string EscapeCueMarkdown(string value) => value
        .Replace("&", "&amp;")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;");

    private static string[]? BuildCueWarnings(bool cueTruncated, bool wasSplit) {
        if (!cueTruncated && !wasSplit) return null;

        var warnings = new List<string>(2);
        if (cueTruncated) warnings.Add("Subtitle cue text was truncated at MaxCueCharacters.");
        if (wasSplit) warnings.Add("Subtitle cue projection was split due to MaxChars.");
        return warnings.ToArray();
    }

    private static IEnumerable<OfficeDocumentMetadataEntry> BuildMetadata(SubtitleParseResult result, string? path) {
        yield return new OfficeDocumentMetadataEntry {
            Id = "subtitle-format",
            Category = "subtitle.document",
            Name = "Format",
            Value = result.Format,
            ValueType = "string"
        };
        yield return new OfficeDocumentMetadataEntry {
            Id = "subtitle-cue-count",
            Category = "subtitle.document",
            Name = "CueCount",
            Value = result.Cues.Count.ToString(CultureInfo.InvariantCulture),
            ValueType = "count"
        };
        for (int index = 0; index < result.Cues.Count; index++) {
            SubtitleCue cue = result.Cues[index];
            string anchor = "subtitle-cue-" + index.ToString("D6", CultureInfo.InvariantCulture);
            yield return new OfficeDocumentMetadataEntry {
                Id = anchor + "-timing",
                Category = "subtitle.cue",
                Name = "Timing",
                Value = SubtitleParser.FormatTimestamp(cue.Start) + " --> " + SubtitleParser.FormatTimestamp(cue.End),
                ValueType = "string",
                SourceObjectId = cue.Identifier,
                Location = new ReaderLocation {
                    Path = path,
                    SourceBlockIndex = index,
                    StartLine = cue.StartLine,
                    EndLine = cue.EndLine,
                    SourceBlockKind = "subtitle-cue",
                    BlockAnchor = anchor
                },
                Attributes = new Dictionary<string, string> {
                    ["startMilliseconds"] = (cue.Start.Ticks / TimeSpan.TicksPerMillisecond).ToString(CultureInfo.InvariantCulture),
                    ["endMilliseconds"] = (cue.End.Ticks / TimeSpan.TicksPerMillisecond).ToString(CultureInfo.InvariantCulture)
                }
            };
        }
    }

    private static string DecodeText(byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }

    private readonly struct SubtitleProjectionPart {
        internal SubtitleProjectionPart(string text, string markdown) {
            Text = text;
            Markdown = markdown;
        }

        internal string Text { get; }
        internal string Markdown { get; }
    }
}
