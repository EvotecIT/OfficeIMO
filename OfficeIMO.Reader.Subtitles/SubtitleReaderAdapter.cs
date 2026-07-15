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
        SubtitleParseResult parsed = SubtitleParser.Parse(content, subtitleOptions, cancellationToken);
        var chunks = new List<ReaderChunk>(parsed.Cues.Count);
        for (int index = 0; index < parsed.Cues.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            SubtitleCue cue = parsed.Cues[index];
            string anchor = "subtitle-cue-" + index.ToString("D6", CultureInfo.InvariantCulture);
            var location = new ReaderLocation {
                Path = input.Source.Path,
                BlockIndex = index,
                SourceBlockIndex = index,
                StartLine = cue.StartLine,
                EndLine = cue.EndLine,
                SourceBlockKind = "subtitle-cue",
                BlockAnchor = anchor
            };
            string markdown = subtitleOptions.IncludeTimestampsInMarkdown
                ? "**" + SubtitleParser.FormatTimestamp(cue.Start) + " → " + SubtitleParser.FormatTimestamp(cue.End) + "**" +
                    Environment.NewLine + Environment.NewLine + cue.Text
                : cue.Text;
            var chunk = new ReaderChunk {
                Id = anchor,
                Kind = ReaderInputKind.Text,
                Location = location,
                Text = cue.Text,
                Markdown = markdown,
                Warnings = cue.Truncated ? new[] { "Subtitle cue text was truncated at MaxCueCharacters." } : null
            };
            DocumentReaderEngine.ApplyAdapterSource(chunk, input, readerOptions.ComputeHashes);
            chunks.Add(chunk);
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
}
