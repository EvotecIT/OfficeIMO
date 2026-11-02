using System;
using System.Collections.Concurrent;
using System.Linq;
using OfficeIMO.Reader;

namespace PSWriteOffice.Services.Reader;

/// <summary>Aggregates the existing resilient outcomes without rereading or retaining full document envelopes.</summary>
internal sealed class OfficeReadIngestSummary {
    private readonly bool _noChunks;
    private readonly ConcurrentBag<ReaderIngestFileResult> _files = new();
    private readonly ConcurrentBag<ReaderChunk> _chunks = new();
    internal OfficeReadIngestSummary(bool noChunks) { _noChunks = noChunks; }
    internal void Add(ReaderDocumentReadOutcome outcome) {
        if (outcome.Succeeded) AddDocument(outcome.Document!);
        else _files.Add(new ReaderIngestFileResult { Path = outcome.Path, Parsed = false, Warnings = new[] { outcome.Error!.Message } });
    }
    internal void AddDocument(OfficeDocumentReadResult document) {
        var source = document.Source;
        _files.Add(new ReaderIngestFileResult {
            Path = source?.Path ?? string.Empty, Parsed = true,
            SourceId = source?.SourceId, SourceHash = source?.SourceHash,
            SourceLastWriteUtc = source?.LastWriteUtc, SourceLengthBytes = source?.LengthBytes,
            ChunksProduced = document.Chunks.Count,
            Warnings = document.Diagnostics.Select(item => item.Message).Distinct().ToArray()
        });
        if (!_noChunks) foreach (var chunk in document.Chunks) _chunks.Add(chunk);
    }
    internal ReaderIngestResult Complete() {
        var files = _files.OrderBy(file => file.Path, StringComparer.Ordinal).ToArray();
        return new ReaderIngestResult {
            Files = files, Chunks = _chunks.OrderBy(chunk => chunk.Location?.Path, StringComparer.Ordinal).ThenBy(chunk => chunk.Location?.BlockIndex).ToArray(),
            FilesScanned = files.Length, FilesParsed = files.Count(file => file.Parsed), FilesSkipped = files.Count(file => !file.Parsed),
            BytesRead = files.Where(file => file.Parsed).Sum(file => file.SourceLengthBytes ?? 0), ChunksProduced = files.Sum(file => file.ChunksProduced),
            Warnings = files.SelectMany(file => file.Warnings ?? Array.Empty<string>()).Distinct().ToArray()
        };
    }
}
