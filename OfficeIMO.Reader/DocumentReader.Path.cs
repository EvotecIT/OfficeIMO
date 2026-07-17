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
    /// <summary>
    /// Reads a supported document file and emits normalized extraction chunks.
    /// </summary>
    /// <param name="path">Source file path.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderChunk> Read(string path, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (Directory.Exists(path)) {
            // Keep Read(file) semantics intact; require explicit folder method for directories.
            throw new IOException($"'{path}' is a directory. Use {nameof(ReadFolder)}(...) to ingest directories.");
        }
        if (!File.Exists(path)) throw new FileNotFoundException($"File '{path}' doesn't exist.", path);

        var opt = NormalizeOptions(options);
        EnforceFileSize(path, ResolveInitialMaxInputBytes(path, opt));
        var source = BuildSourceInfoFromPath(path, opt.ComputeHashes);
        foreach (var chunk in ReadPathCore(path, opt, source, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadPathCore(
        string path,
        ReaderOptions opt,
        SourceInfo source,
        CancellationToken cancellationToken) {
        IEnumerable<ReaderChunk> raw;
        bool hasCustomPathHandler = TryResolvePathHandler(
            path,
            opt,
            out ReaderHandlerDescriptor customPathHandler,
            out ReaderDetectionResult detection);
        if (hasCustomPathHandler) {
            if (customPathHandler.ReadPath != null || customPathHandler.ReadDocumentPath != null) {
                raw = customPathHandler.ReadPath != null
                    ? customPathHandler.ReadPath(path, opt, cancellationToken)
                    : GetDocumentResultChunks(customPathHandler.ReadDocumentPath!(path, opt, cancellationToken), customPathHandler.Id);
            } else if (customPathHandler.ReadDocumentPathAsync != null) {
                throw CreateAsyncOnlyHandlerException(customPathHandler.Id, "path");
            } else {
                raw = ReadBuiltInPath(path, opt, cancellationToken, detection.Kind);
            }
        } else {
            raw = ReadBuiltInPath(path, opt, cancellationToken, detection.Kind);
        }

        foreach (var chunk in raw) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return EnrichChunk(chunk, source, opt.ComputeHashes);
        }
    }

    private static IEnumerable<ReaderChunk> ReadBuiltInPath(
        string path,
        ReaderOptions opt,
        CancellationToken cancellationToken,
        ReaderInputKind detectedKind) {
        ReaderInputKind kind = NormalizeBuiltInDispatchKind(detectedKind);
        if (kind == ReaderInputKind.Unknown && IsEmailArtifact(path, opt, cancellationToken)) {
            kind = ReaderInputKind.Email;
        }
        return kind switch {
            ReaderInputKind.Word => ReadWord(path, opt, cancellationToken),
            ReaderInputKind.Excel => ReadExcel(path, opt, cancellationToken),
            ReaderInputKind.PowerPoint => ReadPowerPoint(path, opt, cancellationToken),
            ReaderInputKind.Markdown => ReadMarkdown(path, opt, cancellationToken),
            ReaderInputKind.Pdf => ReadPdf(path, opt, cancellationToken),
            ReaderInputKind.Email => ReadEmail(path, opt, cancellationToken),
            ReaderInputKind.Calendar => ReadCalendar(path, opt, cancellationToken),
            ReaderInputKind.VCard => ReadVCard(path, opt, cancellationToken),
            ReaderInputKind.Text => ReadText(path, opt, cancellationToken),
            _ => ReadUnknown(path, opt, cancellationToken)
        };
    }

    /// <summary>
    /// Enumerates a folder and ingests all supported files (best-effort), emitting warning chunks for skipped files.
    /// </summary>
    /// <param name="folderPath">Folder path.</param>
    /// <param name="folderOptions">Folder enumeration options.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderChunk> ReadFolder(string folderPath, ReaderFolderOptions? folderOptions = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        foreach (var chunk in ReadFolder(folderPath, folderOptions, options, onProgress: null, cancellationToken))
            yield return chunk;
    }

    /// <summary>
    /// Enumerates a folder and ingests all supported files (best-effort), emitting warning chunks for skipped files.
    /// </summary>
    /// <param name="folderPath">Folder path.</param>
    /// <param name="folderOptions">Folder enumeration options.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="onProgress">Optional progress callback for file-level lifecycle and aggregate counts.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderChunk> ReadFolder(
        string folderPath,
        ReaderFolderOptions? folderOptions,
        ReaderOptions? options,
        Action<ReaderProgress>? onProgress,
        CancellationToken cancellationToken = default) {
        if (folderPath == null) throw new ArgumentNullException(nameof(folderPath));
        if (!Directory.Exists(folderPath)) throw new DirectoryNotFoundException($"Folder '{folderPath}' doesn't exist.");

        var fo = NormalizeFolderOptions(folderOptions);
        var opt = NormalizeOptions(options);
        var fileReadOptions = CloneOptions(opt, computeHashes: false);
        var allowedExt = NormalizeExtensions(fo.Extensions);
        long total = 0;
        int warningIndex = 0;
        var state = new FolderIngestState();

        foreach (var file in EnumerateFilesSafeDeterministic(folderPath, fo, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();

            var ext = TryGetExtension(file);
            if (string.IsNullOrEmpty(ext)) continue;
            if (!allowedExt.Contains(ext!) && !IsDefaultWinmailDat(file, fo.Extensions)) continue;
            if (state.FilesScanned >= fo.MaxFiles) break;

            state.FilesScanned++;
            var source = BuildSourceInfoFromPath(file, computeHash: false);
            NotifyProgress(onProgress, ReaderProgressEventKind.FileStarted, state, source, null, fileChunkCount: null);

            string? statWarning = null;
            var length = source.LengthBytes;
            if (!length.HasValue) {
                statWarning = "Skipped file because metadata could not be read.";
            }
            if (statWarning != null) {
                state.FilesSkipped++;
                state.ChunksProduced++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, statWarning, fileChunkCount: 1);
                yield return EnrichChunk(BuildFolderWarningChunk(file, warningIndex++, statWarning), source, opt.ComputeHashes);
                continue;
            }

            var lengthValue = length.GetValueOrDefault();
            if (fo.MaxTotalBytes.HasValue) {
                if ((total + lengthValue) > fo.MaxTotalBytes.Value) {
                    state.FilesSkipped++;
                    var limitWarning = $"Stopped folder ingestion after reaching MaxTotalBytes ({fo.MaxTotalBytes.Value.ToString(CultureInfo.InvariantCulture)}).";
                    state.ChunksProduced++;
                    NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, limitWarning, fileChunkCount: 1);
                    yield return EnrichChunk(
                        BuildFolderWarningChunk(
                        file,
                        warningIndex,
                        limitWarning),
                        source,
                        opt.ComputeHashes);
                    NotifyProgress(onProgress, ReaderProgressEventKind.Completed, state, source: null, message: null, fileChunkCount: null);
                    yield break;
                }
            }
            total += lengthValue;

            if (opt.MaxInputBytes.HasValue && lengthValue > opt.MaxInputBytes.Value) {
                // Skip too-large files rather than failing the whole folder.
                state.FilesSkipped++;
                var warning = $"Skipped file because it exceeds MaxInputBytes ({lengthValue.ToString(CultureInfo.InvariantCulture)} > {opt.MaxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).";
                state.ChunksProduced++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, warning, fileChunkCount: 1);
                yield return EnrichChunk(
                    BuildFolderWarningChunk(
                    file,
                    warningIndex++,
                    warning),
                    source,
                    opt.ComputeHashes);
                continue;
            }

            if (opt.ComputeHashes && string.IsNullOrWhiteSpace(source.SourceHash)) {
                source.SourceHash = TryComputeFileSha256(file);
            }

            List<ReaderChunk>? fileChunks = null;
            string? readWarning = null;
            try {
                fileChunks = Read(file, fileReadOptions, cancellationToken)
                    .Select(c => EnrichChunk(c, source, opt.ComputeHashes))
                    .ToList();
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                // Keep folder ingestion best-effort; skip files that fail parsing.
                readWarning = $"Skipped file due read error: {ex.GetType().Name}.";
            }
            if (readWarning != null) {
                state.FilesSkipped++;
                state.ChunksProduced++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, readWarning, fileChunkCount: 1);
                yield return EnrichChunk(BuildFolderWarningChunk(file, warningIndex++, readWarning), source, opt.ComputeHashes);
                continue;
            }

            if (fileChunks == null) {
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, "File parsing produced no chunks.", fileChunkCount: 0);
                yield return EnrichChunk(BuildFolderWarningChunk(file, warningIndex++, "File parsing produced no chunks."), source, opt.ComputeHashes);
                continue;
            }

            foreach (var chunk in fileChunks) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return chunk;
            }

            state.FilesParsed++;
            state.BytesRead += lengthValue;
            state.ChunksProduced += fileChunks.Count;
            NotifyProgress(onProgress, ReaderProgressEventKind.FileCompleted, state, source, null, fileChunks.Count);
        }

        NotifyProgress(onProgress, ReaderProgressEventKind.Completed, state, source: null, message: null, fileChunkCount: null);
    }

    /// <summary>
    /// Enumerates a folder and emits one source-level payload per file, ready for direct DB upserts.
    /// </summary>
    /// <param name="folderPath">Folder path.</param>
    /// <param name="folderOptions">Folder enumeration options.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="onProgress">Optional progress callback for file-level lifecycle and aggregate counts.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderSourceDocument> ReadFolderDocuments(
        string folderPath,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        if (folderPath == null) throw new ArgumentNullException(nameof(folderPath));
        if (!Directory.Exists(folderPath)) throw new DirectoryNotFoundException($"Folder '{folderPath}' doesn't exist.");

        var fo = NormalizeFolderOptions(folderOptions);
        var opt = NormalizeOptions(options);
        var fileReadOptions = CloneOptions(opt, computeHashes: false);
        var allowedExt = NormalizeExtensions(fo.Extensions);
        long total = 0;
        var state = new FolderIngestState();

        foreach (var file in EnumerateFilesSafeDeterministic(folderPath, fo, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();

            var ext = TryGetExtension(file);
            if (string.IsNullOrEmpty(ext)) continue;
            if (!allowedExt.Contains(ext!) && !IsDefaultWinmailDat(file, fo.Extensions)) continue;
            if (state.FilesScanned >= fo.MaxFiles) break;

            state.FilesScanned++;
            var source = BuildSourceInfoFromPath(file, computeHash: false);
            NotifyProgress(onProgress, ReaderProgressEventKind.FileStarted, state, source, null, fileChunkCount: null);

            var length = source.LengthBytes;
            if (!length.HasValue) {
                var warning = "Skipped file because metadata could not be read.";
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, warning, fileChunkCount: 0);
                yield return BuildSourceDocument(source, parsed: false, chunks: null, sourceWarnings: new[] { warning });
                continue;
            }

            var lengthValue = length.GetValueOrDefault();
            if (fo.MaxTotalBytes.HasValue && (total + lengthValue) > fo.MaxTotalBytes.Value) {
                var limitWarning = $"Stopped folder ingestion after reaching MaxTotalBytes ({fo.MaxTotalBytes.Value.ToString(CultureInfo.InvariantCulture)}).";
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, limitWarning, fileChunkCount: 0);
                yield return BuildSourceDocument(source, parsed: false, chunks: null, sourceWarnings: new[] { limitWarning });
                NotifyProgress(onProgress, ReaderProgressEventKind.Completed, state, source: null, message: null, fileChunkCount: null);
                yield break;
            }
            total += lengthValue;

            if (opt.MaxInputBytes.HasValue && lengthValue > opt.MaxInputBytes.Value) {
                var warning = $"Skipped file because it exceeds MaxInputBytes ({lengthValue.ToString(CultureInfo.InvariantCulture)} > {opt.MaxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).";
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, warning, fileChunkCount: 0);
                yield return BuildSourceDocument(source, parsed: false, chunks: null, sourceWarnings: new[] { warning });
                continue;
            }

            if (opt.ComputeHashes && string.IsNullOrWhiteSpace(source.SourceHash)) {
                source.SourceHash = TryComputeFileSha256(file);
            }

            List<ReaderChunk>? fileChunks = null;
            string? readWarning = null;
            try {
                fileChunks = Read(file, fileReadOptions, cancellationToken)
                    .Select(c => EnrichChunk(c, source, opt.ComputeHashes))
                    .ToList();
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                // Keep folder ingestion best-effort; skip files that fail parsing.
                readWarning = $"Skipped file due read error: {ex.GetType().Name}.";
            }

            if (readWarning != null) {
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, readWarning, fileChunkCount: 0);
                yield return BuildSourceDocument(source, parsed: false, chunks: null, sourceWarnings: new[] { readWarning });
                continue;
            }

            if (fileChunks == null) {
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, "File parsing produced no chunks.", fileChunkCount: 0);
                yield return BuildSourceDocument(source, parsed: false, chunks: null, sourceWarnings: new[] { "File parsing produced no chunks." });
                continue;
            }

            state.FilesParsed++;
            state.BytesRead += lengthValue;
            state.ChunksProduced += fileChunks.Count;
            NotifyProgress(onProgress, ReaderProgressEventKind.FileCompleted, state, source, null, fileChunks.Count);

            yield return BuildSourceDocument(source, parsed: true, chunks: fileChunks, sourceWarnings: null);
        }

        NotifyProgress(onProgress, ReaderProgressEventKind.Completed, state, source: null, message: null, fileChunkCount: null);
    }

    /// <summary>
    /// Reads a folder and returns ingestion-ready summary/counts with optional chunk materialization.
    /// </summary>
    /// <param name="folderPath">Folder path.</param>
    /// <param name="folderOptions">Folder enumeration options.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="includeChunks">When true, materializes chunks in the result object.</param>
    /// <param name="onProgress">Optional progress callback.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static ReaderIngestResult ReadFolderDetailed(
        string folderPath,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        bool includeChunks = true,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        var chunks = includeChunks ? new List<ReaderChunk>() : null;
        var files = new Dictionary<string, ReaderIngestFileResult>(IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
        var warnings = new List<string>();
        ReaderProgress? completed = null;
        string? currentFilePath = null;

        static void AddWarningToFile(ReaderIngestFileResult file, string? warning) {
            if (string.IsNullOrWhiteSpace(warning)) {
                return;
            }

            var list = file.Warnings?.ToList() ?? new List<string>();
            AddWarning(list, warning);
            file.Warnings = list.Count > 0 ? list : null;
        }

        void HandleProgress(ReaderProgress progress) {
            onProgress?.Invoke(progress);
            if (progress.Kind == ReaderProgressEventKind.Completed) {
                completed = progress;
                return;
            }
            var progressPath = progress.Path;
            if (string.IsNullOrWhiteSpace(progressPath)) return;
            var filePath = progressPath!;

            if (!files.TryGetValue(filePath, out var file)) {
                file = new ReaderIngestFileResult {
                    Path = filePath
                };
                files[filePath] = file;
            }

            file.SourceId = progress.SourceId ?? file.SourceId;
            file.SourceHash = progress.SourceHash ?? file.SourceHash;
            file.SourceLengthBytes = progress.CurrentFileBytes ?? file.SourceLengthBytes;
            file.SourceLastWriteUtc = progress.CurrentFileLastWriteUtc ?? file.SourceLastWriteUtc;

            if (progress.Kind == ReaderProgressEventKind.FileStarted) {
                currentFilePath = filePath;
            } else if (progress.Kind == ReaderProgressEventKind.FileCompleted) {
                file.Parsed = true;
                file.ChunksProduced = progress.CurrentFileChunks ?? file.ChunksProduced;
                currentFilePath = null;
            } else if (progress.Kind == ReaderProgressEventKind.FileSkipped) {
                file.Parsed = false;
                file.ChunksProduced = progress.CurrentFileChunks ?? Math.Max(file.ChunksProduced, 1);
                AddWarningToFile(file, progress.Message);
                AddWarning(warnings, progress.Message);
                currentFilePath = null;
            }
        }

        foreach (var chunk in ReadFolder(folderPath, folderOptions, options, HandleProgress, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (includeChunks) {
                chunks!.Add(chunk);
            }
            if (chunk.Warnings != null && chunk.Warnings.Count > 0) {
                if (!string.IsNullOrWhiteSpace(currentFilePath) && files.TryGetValue(currentFilePath!, out var file)) {
                    for (int i = 0; i < chunk.Warnings.Count; i++) {
                        AddWarningToFile(file, chunk.Warnings[i]);
                    }
                }

                for (int i = 0; i < chunk.Warnings.Count; i++) {
                    AddWarning(warnings, chunk.Warnings[i]);
                }
            }
        }

        var snapshot = completed ?? new ReaderProgress {
            Kind = ReaderProgressEventKind.Completed,
            FilesScanned = files.Count,
            FilesParsed = files.Values.Count(f => f.Parsed),
            FilesSkipped = files.Values.Count(f => !f.Parsed),
            BytesRead = files.Values.Where(f => f.Parsed).Sum(f => f.SourceLengthBytes ?? 0),
            ChunksProduced = includeChunks ? chunks!.Count : files.Values.Sum(f => f.ChunksProduced)
        };

        return new ReaderIngestResult {
            Files = files.Values
                .OrderBy(static f => f.Path, StringComparer.Ordinal)
                .ToArray(),
            Chunks = includeChunks ? chunks! : Array.Empty<ReaderChunk>(),
            FilesScanned = snapshot.FilesScanned,
            FilesParsed = snapshot.FilesParsed,
            FilesSkipped = snapshot.FilesSkipped,
            BytesRead = snapshot.BytesRead,
            ChunksProduced = snapshot.ChunksProduced,
            Warnings = warnings.Count > 0 ? warnings.ToArray() : null
        };
    }

    /// <summary>
    /// Reads a supported file or folder path and returns source-level document payloads with optional chunk shaping.
    /// </summary>
    /// <param name="path">Source file or folder path.</param>
    /// <param name="folderOptions">Folder enumeration options when <paramref name="path"/> is a directory.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="includeDocumentChunks">When true, includes chunk arrays in returned source documents.</param>
    /// <param name="maxReturnedChunks">Optional cap across all returned document chunks.</param>
    /// <param name="onProgress">Optional progress callback for folder reads.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static ReaderPathDocumentResult ReadPathDocumentsDetailed(
        string path,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        bool includeDocumentChunks = true,
        int? maxReturnedChunks = null,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (maxReturnedChunks.HasValue && maxReturnedChunks.Value < 0) {
            throw new ArgumentOutOfRangeException(nameof(maxReturnedChunks), "Chunk cap must be non-negative.");
        }

        IEnumerable<ReaderSourceDocument> documents;
        if (Directory.Exists(path)) {
            documents = ReadFolderDocuments(path, folderOptions, options, onProgress, cancellationToken);
        } else if (File.Exists(path)) {
            documents = new[] { ReadSingleDocument(path, options, cancellationToken) };
        } else {
            throw new FileNotFoundException($"Path '{path}' doesn't exist.", path);
        }

        var files = new List<string>();
        var warnings = new List<string>();
        var remainingChunkBudget = includeDocumentChunks
            ? maxReturnedChunks ?? int.MaxValue
            : 0;
        var truncated = false;
        var returnedChunkCount = 0;
        var returnedTokenEstimate = 0;
        var filesScanned = 0;
        var filesParsed = 0;
        var filesSkipped = 0;
        long bytesRead = 0;
        var chunksProduced = 0;
        var shaped = new List<ReaderSourceDocument>();

        foreach (var source in documents) {
            cancellationToken.ThrowIfCancellationRequested();
            filesScanned++;
            if (!string.IsNullOrWhiteSpace(source.Path)) {
                files.Add(source.Path);
            }
            if (source.Parsed) {
                filesParsed++;
                bytesRead += source.SourceLengthBytes ?? 0;
            } else {
                filesSkipped++;
            }
            chunksProduced += source.ChunksProduced;

            var shapedSource = ShapeSourceDocument(
                source,
                includeDocumentChunks,
                ref remainingChunkBudget,
                ref truncated,
                ref returnedChunkCount,
                ref returnedTokenEstimate,
                warnings);
            shaped.Add(shapedSource);
        }

        if (truncated && includeDocumentChunks && maxReturnedChunks.HasValue) {
            AddWarning(
                warnings,
                $"Stopped after reaching MaxReturnedChunks ({maxReturnedChunks.Value.ToString(CultureInfo.InvariantCulture)}).");
        }

        return new ReaderPathDocumentResult {
            Files = files.ToArray(),
            Documents = shaped.ToArray(),
            FilesScanned = filesScanned,
            FilesParsed = filesParsed,
            FilesSkipped = filesSkipped,
            BytesRead = bytesRead,
            ChunksProduced = chunksProduced,
            ChunksReturned = returnedChunkCount,
            TokenEstimateReturned = returnedTokenEstimate,
            Truncated = truncated,
            Warnings = warnings.Count > 0 ? warnings.ToArray() : null
        };
    }

}
