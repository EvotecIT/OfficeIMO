using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static IEnumerable<ReaderChunk> ChunkPlainTextByParagraphs(
        string path,
        ReaderOptions opt,
        ReaderInputKind kind,
        CancellationToken ct,
        bool treatAsMarkdown) {
        var fileName = Path.GetFileName(path);
        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        int chunkIndex = 0;
        int? firstLine = null;
        var warnings = new List<string>(capacity: 2);
        int lineNo = 0;

        foreach (var line in File.ReadLines(path)) {
            ct.ThrowIfCancellationRequested();
            lineNo++;

            if (firstLine == null) firstLine = lineNo;

            // Prefer splitting at empty lines when close to the cap.
            if (current.Length > 0 && current.Length >= (opt.MaxChars - 256) && string.IsNullOrWhiteSpace(line)) {
                yield return BuildTextChunk(path, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = null;
                continue;
            }

            if (WouldExceed(opt, current, line)) {
                yield return BuildTextChunk(path, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = lineNo;
            }

            AppendLineCapped(opt, current, line, warnings);
        }

        if (current.Length > 0) {
            yield return BuildTextChunk(path, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
        }
    }

    private static IEnumerable<ReaderChunk> ChunkPlainTextFromText(
        string text,
        string? sourceName,
        string fileName,
        ReaderOptions opt,
        ReaderInputKind kind,
        CancellationToken ct,
        bool treatAsMarkdown) {
        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        int chunkIndex = 0;
        int? firstLine = null;
        var warnings = new List<string>(capacity: 2);
        int lineNo = 0;

        using var sr = new StringReader(text ?? string.Empty);
        string? line;
        while ((line = sr.ReadLine()) != null) {
            ct.ThrowIfCancellationRequested();
            lineNo++;

            if (firstLine == null) firstLine = lineNo;

            if (current.Length > 0 && current.Length >= (opt.MaxChars - 256) && string.IsNullOrWhiteSpace(line)) {
                yield return BuildTextChunk(sourceName ?? fileName, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = null;
                continue;
            }

            if (WouldExceed(opt, current, line)) {
                yield return BuildTextChunk(sourceName ?? fileName, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = lineNo;
            }

            AppendLineCapped(opt, current, line, warnings);
        }

        if (current.Length > 0) {
            yield return BuildTextChunk(sourceName ?? fileName, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
        }
    }

    private static IEnumerable<ReaderChunk> ChunkMarkdownFromText(string text, string? sourceName, string fileName, ReaderOptions opt, CancellationToken ct) {
        if (!opt.MarkdownChunkByHeadings) {
            foreach (var c in ChunkPlainTextFromText(text, sourceName, fileName, opt, ReaderInputKind.Markdown, ct, treatAsMarkdown: true))
                yield return c;
            yield break;
        }

        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        int chunkIndex = 0;
        int? firstLine = null;
        int? lastLine = null;
        int? firstSourceStartLine = null;
        int? lastSourceEndLine = null;
        int? firstSourceBlockIndex = null;
        string? firstHeadingPath = null;
        string? firstHierarchyHeadingPath = null;
        string? firstHeadingSlug = null;
        string? firstSourceBlockKind = null;
        string? firstBlockAnchor = null;
        var warnings = new List<string>(capacity: 2);
        bool oversizeBlockWarningAdded = false;
        List<ReaderTable>? tables = null;
        List<ReaderVisual>? visuals = null;

        foreach (var block in ParseMarkdownBlocksForChunking(text, opt, ct)) {
            ct.ThrowIfCancellationRequested();

            if (block.StartsHeading && current.Length > 0) {
                yield return BuildMarkdownChunk(sourceName ?? fileName, fileName, chunkIndex, firstSourceStartLine, lastSourceEndLine, firstLine, lastLine, firstSourceBlockIndex, firstHeadingPath, firstHierarchyHeadingPath, firstHeadingSlug, firstSourceBlockKind, firstBlockAnchor, current.ToString().TrimEnd(), warnings, tables, visuals);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                oversizeBlockWarningAdded = false;
                tables = null;
                visuals = null;
                firstLine = null;
                lastLine = null;
                firstSourceStartLine = null;
                lastSourceEndLine = null;
                firstSourceBlockIndex = null;
                firstHeadingPath = null;
                firstHierarchyHeadingPath = null;
                firstHeadingSlug = null;
                firstSourceBlockKind = null;
                firstBlockAnchor = null;
            }

            if (WouldExceedMarkdownBlock(opt, current, block.Markdown)) {
                yield return BuildMarkdownChunk(sourceName ?? fileName, fileName, chunkIndex, firstSourceStartLine, lastSourceEndLine, firstLine, lastLine, firstSourceBlockIndex, firstHeadingPath, firstHierarchyHeadingPath, firstHeadingSlug, firstSourceBlockKind, firstBlockAnchor, current.ToString().TrimEnd(), warnings, tables, visuals);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                oversizeBlockWarningAdded = false;
                tables = null;
                visuals = null;
                firstLine = null;
                lastLine = null;
                firstSourceStartLine = null;
                lastSourceEndLine = null;
                firstSourceBlockIndex = null;
                firstHeadingPath = null;
                firstHierarchyHeadingPath = null;
                firstHeadingSlug = null;
                firstSourceBlockKind = null;
                firstBlockAnchor = null;
            }

            if (firstLine == null) firstLine = block.StartLine;
            if (firstSourceStartLine == null) firstSourceStartLine = block.SourceStartLine;
            if (firstSourceBlockIndex == null) firstSourceBlockIndex = block.BlockIndex;
            if (firstHeadingPath == null) firstHeadingPath = block.HeadingPath;
            if (firstHierarchyHeadingPath == null) firstHierarchyHeadingPath = block.HierarchyHeadingPath;
            if (firstHeadingSlug == null) firstHeadingSlug = block.HeadingSlug;
            if (firstSourceBlockKind == null) firstSourceBlockKind = block.BlockKind;
            if (firstBlockAnchor == null) firstBlockAnchor = block.BlockAnchor;
            lastLine = block.EndLine;
            lastSourceEndLine = block.SourceEndLine;

            AppendMarkdownBlock(current, block.Markdown);
            if (block.Markdown.Length > opt.MaxChars && !oversizeBlockWarningAdded) {
                warnings.Add("A single markdown block exceeded MaxChars and was preserved as one chunk.");
                oversizeBlockWarningAdded = true;
            }

            if (block.Tables.Count > 0) {
                tables ??= new List<ReaderTable>(capacity: block.Tables.Count);
                tables.AddRange(block.Tables);
            }

            if (block.Visuals.Count > 0) {
                visuals ??= new List<ReaderVisual>(capacity: block.Visuals.Count);
                visuals.AddRange(block.Visuals);
            }
        }

        if (current.Length > 0) {
            yield return BuildMarkdownChunk(sourceName ?? fileName, fileName, chunkIndex, firstSourceStartLine, lastSourceEndLine, firstLine, lastLine, firstSourceBlockIndex, firstHeadingPath, firstHierarchyHeadingPath, firstHeadingSlug, firstSourceBlockKind, firstBlockAnchor, current.ToString().TrimEnd(), warnings, tables, visuals);
        }
    }

    private static List<ReaderChunk> ChunkPdfText(
        string path,
        string fileName,
        int pageNumber,
        string text,
        string? markdown,
        ReaderOptions opt,
        int startChunkIndex,
        CancellationToken ct,
        out int nextChunkIndex) {
        var list = new List<ReaderChunk>();
        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        var outIndex = startChunkIndex;
        int? firstLine = null;
        var warnings = new List<string>(capacity: 2);
        int lineNo = 0;

        using var sr = new StringReader(text ?? string.Empty);
        string? line;
        while ((line = sr.ReadLine()) != null) {
            ct.ThrowIfCancellationRequested();
            lineNo++;

            if (firstLine == null) firstLine = lineNo;

            if (current.Length > 0 && current.Length >= (opt.MaxChars - 256) && string.IsNullOrWhiteSpace(line)) {
                list.Add(BuildPdfChunk(path, fileName, pageNumber, outIndex, firstLine, current.ToString().TrimEnd(), warnings));
                outIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = null;
                continue;
            }

            if (WouldExceed(opt, current, line)) {
                list.Add(BuildPdfChunk(path, fileName, pageNumber, outIndex, firstLine, current.ToString().TrimEnd(), warnings));
                outIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = lineNo;
            }

            AppendLineCapped(opt, current, line, warnings);
        }

        if (current.Length > 0) {
            list.Add(BuildPdfChunk(path, fileName, pageNumber, outIndex, firstLine, current.ToString().TrimEnd(), warnings));
            outIndex++;
        }

        if (!string.IsNullOrWhiteSpace(markdown) && list.Count == 1 && markdown!.Length <= opt.MaxChars) {
            list[0].Markdown = markdown.Trim();
        }

        nextChunkIndex = outIndex;
        return list;
    }

    private static string BuildPdfPageText(PdfLogicalPage page) {
        if (page.TextBlocks.Count == 0) {
            return string.Empty;
        }

        var builder = new StringBuilder();
        for (int i = 0; i < page.TextBlocks.Count; i++) {
            if (i > 0) {
                builder.AppendLine();
            }

            builder.Append(page.TextBlocks[i].Text);
        }

        return builder.ToString();
    }

    private static ReaderChunk BuildMarkdownChunk(
        string path,
        string fileName,
        int chunkIndex,
        int? sourceStartLine,
        int? sourceEndLine,
        int? firstLine,
        int? lastLine,
        int? firstSourceBlockIndex,
        string? headingPath,
        string? hierarchyHeadingPath,
        string? headingSlug,
        string? sourceBlockKind,
        string? blockAnchor,
        string markdown,
        List<string> warnings,
        List<ReaderTable>? tables,
        List<ReaderVisual>? visuals) {
        var id = BuildStableId("md", fileName, chunkIndex, firstSourceBlockIndex ?? firstLine);
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Markdown,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = chunkIndex,
                SourceBlockIndex = firstSourceBlockIndex,
                StartLine = sourceStartLine,
                EndLine = sourceEndLine,
                HeadingPath = headingPath,
                HierarchyHeadingPath = hierarchyHeadingPath,
                HierarchyHeadingDisplayPath = headingPath,
                HeadingSlug = headingSlug,
                SourceBlockKind = sourceBlockKind,
                BlockAnchor = blockAnchor,
                NormalizedStartLine = firstLine,
                NormalizedEndLine = lastLine
            },
            Text = markdown,
            Markdown = markdown,
            Tables = tables != null && tables.Count > 0 ? tables.ToArray() : null,
            Visuals = visuals != null && visuals.Count > 0 ? visuals.ToArray() : null,
            Warnings = warnings.Count > 0 ? warnings.ToArray() : null
        };
    }

    private static ReaderChunk BuildTextChunk(
        string path,
        string fileName,
        ReaderInputKind kind,
        int chunkIndex,
        int? firstLine,
        string text,
        List<string> warnings,
        bool treatAsMarkdown) {
        var id = BuildStableId(GetTextChunkIdNamespace(kind), fileName, chunkIndex, firstLine);
        return new ReaderChunk {
            Id = id,
            Kind = kind,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = chunkIndex,
                StartLine = firstLine
            },
            Text = text,
            Markdown = treatAsMarkdown ? text : null,
            Warnings = warnings.Count > 0 ? warnings.ToArray() : null
        };
    }

    private static string GetTextChunkIdNamespace(ReaderInputKind kind) {
        switch (kind) {
            case ReaderInputKind.Text:
                return "text";
            case ReaderInputKind.Calendar:
                return "calendar";
            case ReaderInputKind.VCard:
                return "vcard";
            default:
                return "unknown";
        }
    }

    private static ReaderChunk BuildPdfChunk(
        string path,
        string fileName,
        int pageNumber,
        int chunkIndex,
        int? firstLine,
        string text,
        List<string> warnings) {
        var id = BuildStableId("pdf", fileName, chunkIndex, firstLine);
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Pdf,
            Location = new ReaderLocation {
                Path = path,
                Page = pageNumber,
                BlockIndex = chunkIndex,
                SourceBlockIndex = pageNumber - 1,
                StartLine = firstLine
            },
            Text = text,
            Markdown = null,
            Warnings = warnings.Count > 0 ? warnings.ToArray() : null
        };
    }

    private static ReaderChunk BuildPdfEmptyChunk(
        string path,
        string fileName,
        int pageNumber,
        int chunkIndex) {
        var id = BuildStableId("pdf", fileName, chunkIndex, null);
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Pdf,
            Location = new ReaderLocation {
                Path = path,
                Page = pageNumber,
                BlockIndex = chunkIndex,
                SourceBlockIndex = pageNumber - 1
            },
            Text = string.Empty,
            Markdown = null,
            Warnings = new[] { "No extractable text found on this PDF page." }
        };
    }

    private static ReaderChunk BuildFolderWarningChunk(string path, int warningIndex, string warning) {
        var fileName = Path.GetFileName(path);
        if (string.IsNullOrWhiteSpace(fileName)) fileName = "folder";

        return new ReaderChunk {
            Id = BuildStableId("warn", fileName, warningIndex, null),
            Kind = ReaderInputKind.Unknown,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = warningIndex
            },
            Text = string.Empty,
            Markdown = null,
            Warnings = new[] { warning }
        };
    }

    private static void NotifyProgress(
        Action<ReaderProgress>? onProgress,
        ReaderProgressEventKind kind,
        FolderIngestState state,
        SourceInfo? source,
        string? message,
        int? fileChunkCount) {
        if (onProgress == null) return;

        onProgress(new ReaderProgress {
            Kind = kind,
            Path = source?.Path,
            SourceId = source?.SourceId,
            SourceHash = source?.SourceHash,
            FilesScanned = state.FilesScanned,
            FilesParsed = state.FilesParsed,
            FilesSkipped = state.FilesSkipped,
            BytesRead = state.BytesRead,
            ChunksProduced = state.ChunksProduced,
            Message = message,
            CurrentFileBytes = source?.LengthBytes,
            CurrentFileChunks = fileChunkCount,
            CurrentFileLastWriteUtc = source?.LastWriteUtc
        });
    }

    private static ReaderSourceDocument ReadSingleDocument(
        string path,
        ReaderOptions? options,
        CancellationToken cancellationToken) {
        var opt = NormalizeOptions(options);
        var source = BuildSourceInfoFromPath(path, computeHash: false,
            cancellationToken);

        List<ReaderChunk>? chunks = null;
        string? warning = null;
        try {
            EnforceFileSize(path, ResolveInitialMaxInputBytes(path, opt));
            chunks = ReadPathCore(path, opt, source, cancellationToken).ToList();
            if (opt.ComputeHashes) {
                source.SourceHash = TryComputeFileSha256(path,
                    cancellationToken);
                if (!string.IsNullOrWhiteSpace(source.SourceHash)) {
                    for (int i = 0; i < chunks.Count; i++) {
                        chunks[i].SourceHash ??= source.SourceHash;
                    }
                }
            }
        } catch (OperationCanceledException) {
            throw;
        } catch (NotSupportedException ex) {
            warning = $"Skipped (unsupported): {path} ({ex.Message})";
        } catch (IOException ex) {
            warning = $"Skipped (I/O): {path} ({ex.Message})";
        } catch (Exception ex) {
            warning = $"Skipped (error): {path} ({ex.Message})";
        }

        return chunks == null
            ? BuildSourceDocument(source, parsed: false, chunks: null, sourceWarnings: warning == null ? null : new[] { warning })
            : BuildSourceDocument(source, parsed: true, chunks: chunks, sourceWarnings: null);
    }

    private static ReaderSourceDocument ShapeSourceDocument(
        ReaderSourceDocument source,
        bool includeDocumentChunks,
        ref int remainingChunkBudget,
        ref bool truncated,
        ref int returnedChunkCount,
        ref int returnedTokenEstimate,
        List<string> aggregateWarnings) {
        if (source == null) throw new ArgumentNullException(nameof(source));

        if (source.Warnings != null) {
            for (var i = 0; i < source.Warnings.Count; i++) {
                AddWarning(aggregateWarnings, source.Warnings[i]);
            }
        }

        IReadOnlyList<ReaderChunk> returnedChunks = Array.Empty<ReaderChunk>();
        if (includeDocumentChunks && source.Chunks.Count > 0) {
            if (remainingChunkBudget >= source.Chunks.Count) {
                returnedChunks = source.Chunks;
                remainingChunkBudget -= source.Chunks.Count;
            } else {
                var take = Math.Max(remainingChunkBudget, 0);
                returnedChunks = take == 0 ? Array.Empty<ReaderChunk>() : source.Chunks.Take(take).ToArray();
                remainingChunkBudget = 0;
                truncated = true;
            }
        }

        for (var i = 0; i < returnedChunks.Count; i++) {
            returnedChunkCount++;
            returnedTokenEstimate += returnedChunks[i].TokenEstimate ?? EstimateTokenCount(returnedChunks[i].Text);
        }

        List<string>? shapedWarnings = null;
        if (source.Warnings != null && source.Warnings.Count > 0) {
            shapedWarnings = source.Warnings.ToList();
        }

        if (includeDocumentChunks && returnedChunks.Count < source.Chunks.Count) {
            shapedWarnings ??= new List<string>();
            AddWarning(shapedWarnings, "Document chunk payload was truncated due to MaxReturnedChunks.");
        }

        return new ReaderSourceDocument {
            Path = source.Path,
            SourceId = source.SourceId,
            SourceHash = source.SourceHash,
            SourceLastWriteUtc = source.SourceLastWriteUtc,
            SourceLengthBytes = source.SourceLengthBytes,
            Parsed = source.Parsed,
            ChunksProduced = source.ChunksProduced,
            TokenEstimateTotal = source.TokenEstimateTotal,
            Warnings = shapedWarnings is null || shapedWarnings.Count == 0 ? null : shapedWarnings.ToArray(),
            Chunks = returnedChunks
        };
    }

    private static ReaderSourceDocument BuildSourceDocument(
        SourceInfo source,
        bool parsed,
        IReadOnlyList<ReaderChunk>? chunks,
        IReadOnlyList<string>? sourceWarnings) {
        var chunkList = chunks ?? Array.Empty<ReaderChunk>();
        var warnings = new List<string>();
        if (sourceWarnings != null) {
            for (int i = 0; i < sourceWarnings.Count; i++) {
                AddWarning(warnings, sourceWarnings[i]);
            }
        }

        int tokenEstimateTotal = 0;
        for (int i = 0; i < chunkList.Count; i++) {
            var chunk = chunkList[i];
            tokenEstimateTotal += chunk.TokenEstimate ?? EstimateTokenCount(chunk.Text);

            if (chunk.Warnings == null) continue;
            for (int j = 0; j < chunk.Warnings.Count; j++) {
                AddWarning(warnings, chunk.Warnings[j]);
            }
        }

        return new ReaderSourceDocument {
            Path = source.Path,
            SourceId = source.SourceId,
            SourceHash = source.SourceHash,
            SourceLastWriteUtc = source.LastWriteUtc,
            SourceLengthBytes = source.LengthBytes,
            Parsed = parsed,
            ChunksProduced = chunkList.Count,
            TokenEstimateTotal = tokenEstimateTotal,
            Warnings = warnings.Count > 0 ? warnings : null,
            Chunks = chunkList
        };
    }

    private static void AddWarning(List<string> warnings, string? warning) {
        if (warnings == null || string.IsNullOrWhiteSpace(warning)) {
            return;
        }

        if (warnings.Any(existing => string.Equals(existing, warning, StringComparison.OrdinalIgnoreCase))) {
            return;
        }

        warnings.Add(warning!);
    }

    private static ReaderChunk EnrichChunk(ReaderChunk chunk, SourceInfo source, bool computeHashes) {
        if (chunk == null) throw new ArgumentNullException(nameof(chunk));
        if (source == null) throw new ArgumentNullException(nameof(source));

        chunk.SourceId ??= source.SourceId;
        chunk.SourceHash ??= source.SourceHash;
        chunk.SourceLastWriteUtc ??= source.LastWriteUtc;
        chunk.SourceLengthBytes ??= source.LengthBytes;
        if (!chunk.TokenEstimate.HasValue) {
            chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown ?? chunk.Text);
        }
        if (computeHashes && string.IsNullOrWhiteSpace(chunk.ChunkHash)) {
            chunk.ChunkHash = ComputeChunkHash(chunk);
        }
        return chunk;
    }

    private static int EstimateTokenCount(string? text) {
        var safeText = text ?? string.Empty;
        if (safeText.Length == 0) return 0;
        // Heuristic: roughly 4 characters per token for mixed English/code.
        return Math.Max(1, (safeText.Length + 3) / 4);
    }

    private static string ComputeChunkHash(ReaderChunk chunk) {
        var data = string.Join("|",
            chunk.Kind.ToString(),
            chunk.SourceId ?? string.Empty,
            chunk.Location.Path ?? string.Empty,
            chunk.Location.HeadingPath ?? string.Empty,
            chunk.Location.HeadingSlug ?? string.Empty,
            chunk.Location.SourceBlockKind ?? string.Empty,
            chunk.Location.BlockAnchor ?? string.Empty,
            chunk.Location.Sheet ?? string.Empty,
            chunk.Location.A1Range ?? string.Empty,
            chunk.Location.Page?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.Slide?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.StartLine?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.NormalizedStartLine?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.NormalizedEndLine?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Text ?? string.Empty,
            chunk.Markdown ?? string.Empty);
        return ComputeSha256Hex(data);
    }

    private static SourceInfo BuildSourceInfoFromPath(string path,
        bool computeHash, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string normalizedPath = NormalizePathForId(path);
        string sourceId = BuildSourceId(normalizedPath);

        DateTime? lastWriteUtc = null;
        long? lengthBytes = null;
        try {
            var fi = new FileInfo(path);
            if (fi.Exists) {
                lastWriteUtc = fi.LastWriteTimeUtc;
                lengthBytes = fi.Length;
            }
        } catch {
            // Best-effort metadata; leave null on failure.
        }

        string? sourceHash = null;
        if (computeHash) {
            sourceHash = TryComputeFileSha256(path, cancellationToken);
        }

        return new SourceInfo {
            Path = path,
            SourceId = sourceId,
            SourceHash = sourceHash,
            LastWriteUtc = lastWriteUtc,
            LengthBytes = lengthBytes
        };
    }

    private static SourceInfo BuildSourceInfoFromStream(Stream stream,
        string? sourceName, bool computeHash,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string logicalName = "memory";
        if (!string.IsNullOrWhiteSpace(sourceName)) {
            logicalName = sourceName!.Trim();
        }
        string sourceId = BuildSourceId(logicalName);

        long? lengthBytes = null;
        try {
            if (stream.CanSeek) {
                lengthBytes = stream.Length;
            }
        } catch {
            // Best-effort metadata.
        }

        string? sourceHash = null;
        if (computeHash) {
            sourceHash = TryComputeStreamSha256(stream, cancellationToken);
        }

        return new SourceInfo {
            Path = logicalName,
            SourceId = sourceId,
            SourceHash = sourceHash,
            LastWriteUtc = null,
            LengthBytes = lengthBytes
        };
    }

    private static string? TryComputeFileSha256(string path,
        CancellationToken cancellationToken = default) {
        try {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return ComputeSha256Hex(fs, cancellationToken);
        } catch (OperationCanceledException) {
            throw;
        } catch {
            return null;
        }
    }

    private static string? TryComputeStreamSha256(Stream stream,
        CancellationToken cancellationToken = default) {
        if (stream == null || !stream.CanSeek) return null;
        long position;
        try {
            position = stream.Position;
        } catch {
            return null;
        }

        try {
            cancellationToken.ThrowIfCancellationRequested();
            stream.Position = 0;
            var hash = ComputeSha256Hex(stream, cancellationToken);
            stream.Position = position;
            return hash;
        } catch (OperationCanceledException) {
            try {
                stream.Position = position;
            } catch {
                // ignore
            }
            throw;
        } catch {
            try {
                stream.Position = position;
            } catch {
                // ignore
            }
            return null;
        }
    }

    private static string ComputeSha256Hex(string value) {
        using var sha = SHA256.Create();
        var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
        var hash = sha.ComputeHash(bytes);
        return ConvertToHexLower(hash);
    }

    private static string ComputeSha256Hex(Stream stream,
        CancellationToken cancellationToken = default) {
        using var sha = SHA256.Create();
        var buffer = new byte[81920];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            sha.TransformBlock(buffer, 0, read, buffer, 0);
        }
        cancellationToken.ThrowIfCancellationRequested();
        sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return ConvertToHexLower(sha.Hash!);
    }

    private static string ConvertToHexLower(byte[] bytes) {
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            sb.Append(bytes[i].ToString("x2", CultureInfo.InvariantCulture));
        }
        return sb.ToString();
    }

    private static string BuildSourceId(string sourceKey) {
        var normalized = sourceKey ?? string.Empty;
        if (IsWindows()) {
            normalized = normalized.ToLowerInvariant();
        }
        return "src:" + ComputeSha256Hex(normalized);
    }

    private static bool IsWindows() {
        return Path.DirectorySeparatorChar == '\\';
    }

    private static string NormalizePathForId(string path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;
        string full;
        try {
            full = Path.GetFullPath(path);
        } catch {
            full = path;
        }
        return full.Replace('\\', '/');
    }

    private static void UpdateHeadingStack(List<MarkdownHeadingState> stack, int level, string text, string slug) {
        if (level < 1) return;
        if (string.IsNullOrWhiteSpace(text)) text = $"Heading {level}";

        for (int i = stack.Count - 1; i >= 0; i--) {
            if (stack[i].Level >= level) stack.RemoveAt(i);
        }
        stack.Add(new MarkdownHeadingState(level, CollapseWhitespace(text), slug));
    }

    private static string? BuildHeadingPath(List<MarkdownHeadingState> stack) {
        string[] values = stack
            .Select(static heading => heading.Text.Trim())
            .Where(static heading => heading.Length > 0)
            .ToArray();
        return values.Length == 0 ? null : string.Join(" > ", values);
    }

    private static string? BuildHierarchyHeadingPath(List<MarkdownHeadingState> stack) {
        return ReaderHeadingPath.Combine(stack.Select(static heading => heading.Text));
    }

    private static string? BuildHeadingSlug(List<MarkdownHeadingState> stack) {
        if (stack.Count == 0) return null;
        var slug = stack[stack.Count - 1].Slug?.Trim();
        return string.IsNullOrEmpty(slug) ? null : slug;
    }

    private static bool WouldExceed(ReaderOptions opt, StringBuilder current, string nextLine) {
        // +1 for newline to keep final chunk shape similar to file.
        int nextLen = nextLine?.Length ?? 0;
        int extra = (current.Length == 0 ? 0 : 1) + nextLen;
        return current.Length > 0 && (current.Length + extra) > opt.MaxChars;
    }

    private static bool WouldExceedMarkdownBlock(ReaderOptions opt, StringBuilder current, string nextBlockMarkdown) {
        int nextLen = nextBlockMarkdown?.Length ?? 0;
        int extra = current.Length == 0 ? 0 : 2 + nextLen;
        return current.Length > 0 && (current.Length + extra) > opt.MaxChars;
    }

    private static void AppendLineCapped(ReaderOptions opt, StringBuilder sb, string line, List<string> warnings) {
        if (sb.Length > 0) sb.AppendLine();

        var s = line ?? string.Empty;
        // Hard-cap pathological single lines so callers don't accidentally ingest megabytes in one chunk.
        if (s.Length > opt.MaxChars) {
            s = s.Substring(0, opt.MaxChars) + " <!-- truncated -->";
            warnings.Add("A single line exceeded MaxChars and was truncated.");
        }
        sb.Append(s);
    }

    private static void AppendMarkdownBlock(StringBuilder sb, string markdown) {
        if (sb.Length > 0) {
            sb.AppendLine();
            sb.AppendLine();
        }
        sb.Append(NormalizeMarkdownLineEndings(markdown).TrimEnd());
    }

}
