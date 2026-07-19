using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    public static IEnumerable<ReaderChunk> Read(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ValidateFilePath(path);
        ReaderOptions effective = NormalizeOptions(options);
        EnforceFileSize(path, ResolveInitialMaxInputBytes(path, effective));
        if (!TryResolvePathHandler(
                path,
                effective,
                cancellationToken,
                out ReaderHandlerDescriptor handler,
                out ReaderDetectionResult detection)) {
            throw CreateUnsupportedInputException(path, detection);
        }

        SourceInfo source = BuildSourceInfoFromPath(path, ShouldComputeSourceHash(handler, effective), cancellationToken);
        IEnumerable<ReaderChunk> chunks;
        if (handler.ReadPath != null) {
            chunks = handler.ReadPath(path, effective, cancellationToken)
                ?? throw new InvalidOperationException($"Reader handler '{handler.Id}' returned null chunks.");
        } else if (handler.ReadDocumentPath != null) {
            OfficeDocumentReadResult result = ValidateDocumentResult(
                handler.ReadDocumentPath(path, effective, cancellationToken), handler.Id);
            chunks = result.Chunks ?? Array.Empty<ReaderChunk>();
        } else if (handler.ReadDocumentPathAsync != null) {
            throw CreateAsyncOnlyHandlerException(handler.Id, "path");
        } else if (handler.SupportsStreamInput) {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return Read(stream, path, effective, cancellationToken).ToArray();
        } else {
            throw new NotSupportedException($"Reader handler '{handler.Id}' does not support path input.");
        }

        return chunks.Select(chunk => EnrichChunk(chunk, source, effective.ComputeHashes)).ToArray();
    }

    public static IEnumerable<ReaderChunk> ReadFolder(
        string folderPath,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) =>
        ReadFolder(folderPath, folderOptions, options, onProgress: null, cancellationToken);

    public static IEnumerable<ReaderChunk> ReadFolder(
        string folderPath,
        ReaderFolderOptions? folderOptions,
        ReaderOptions? options,
        Action<ReaderProgress>? onProgress,
        CancellationToken cancellationToken = default) {
        foreach (ReaderSourceDocument document in ReadFolderDocumentsCore(
                     folderPath, folderOptions, options, includeSkippedWarningChunks: true,
                     onProgress, cancellationToken)) {
            foreach (ReaderChunk chunk in document.Chunks) yield return chunk;
        }
    }

    public static IEnumerable<ReaderSourceDocument> ReadFolderDocuments(
        string folderPath,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) =>
        ReadFolderDocumentsCore(folderPath, folderOptions, options,
            includeSkippedWarningChunks: false, onProgress, cancellationToken);

    public static ReaderIngestResult ReadFolderDetailed(
        string folderPath,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        bool includeChunks = true,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        ReaderSourceDocument[] documents = ReadFolderDocumentsCore(
            folderPath, folderOptions, options, includeSkippedWarningChunks: true,
            onProgress, cancellationToken).ToArray();
        var files = documents.Select(static document => new ReaderIngestFileResult {
            Path = document.Path,
            SourceId = document.SourceId,
            SourceHash = document.SourceHash,
            SourceLastWriteUtc = document.SourceLastWriteUtc,
            SourceLengthBytes = document.SourceLengthBytes,
            Parsed = document.Parsed,
            ChunksProduced = document.ChunksProduced,
            Warnings = document.Warnings
        }).ToArray();
        return new ReaderIngestResult {
            Files = files,
            Chunks = includeChunks ? documents.SelectMany(static document => document.Chunks).ToArray() : Array.Empty<ReaderChunk>(),
            FilesScanned = files.Length,
            FilesParsed = files.Count(static file => file.Parsed),
            FilesSkipped = files.Count(static file => !file.Parsed),
            BytesRead = files.Where(static file => file.Parsed).Sum(static file => file.SourceLengthBytes ?? 0),
            ChunksProduced = files.Sum(static file => file.ChunksProduced),
            Warnings = MergeWarnings(documents)
        };
    }

    public static ReaderPathDocumentResult ReadPathDocumentsDetailed(
        string path,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        bool includeDocumentChunks = true,
        int? maxReturnedChunks = null,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        ReaderSourceDocument[] documents = Directory.Exists(path)
            ? ReadFolderDocumentsCore(path, folderOptions, options,
                includeSkippedWarningChunks: true, onProgress, cancellationToken).ToArray()
            : new[] { ReadSingleDocument(path, options, cancellationToken) };

        int remaining = Math.Max(0, maxReturnedChunks ?? int.MaxValue);
        bool truncated = false;
        const string truncationWarning = "Returned chunks were truncated because MaxReturnedChunks was reached.";
        var shaped = new List<ReaderSourceDocument>(documents.Length);
        foreach (ReaderSourceDocument document in documents) {
            IReadOnlyList<ReaderChunk> chunks = Array.Empty<ReaderChunk>();
            bool documentTruncated = false;
            if (includeDocumentChunks && remaining > 0) {
                int take = Math.Min(remaining, document.Chunks.Count);
                chunks = document.Chunks.Take(take).ToArray();
                remaining -= take;
                documentTruncated = take < document.Chunks.Count;
                truncated |= documentTruncated;
            } else if (includeDocumentChunks && document.Chunks.Count > 0) {
                documentTruncated = true;
                truncated = true;
            }
            shaped.Add(CloneSourceDocument(document, chunks,
                documentTruncated ? truncationWarning : null));
        }

        ReaderChunk[] returned = shaped.SelectMany(static document => document.Chunks).ToArray();
        return new ReaderPathDocumentResult {
            Files = documents.Select(static document => document.Path).ToArray(),
            Documents = shaped,
            FilesScanned = documents.Length,
            FilesParsed = documents.Count(static document => document.Parsed),
            FilesSkipped = documents.Count(static document => !document.Parsed),
            BytesRead = documents.Where(static document => document.Parsed).Sum(static document => document.SourceLengthBytes ?? 0),
            ChunksProduced = documents.Sum(static document => document.ChunksProduced),
            ChunksReturned = returned.Length,
            TokenEstimateReturned = returned.Sum(static chunk => chunk.TokenEstimate ?? EstimateTokenCount(chunk.Text)),
            Truncated = truncated,
            Warnings = MergeWarnings(shaped)
        };
    }

    private static IEnumerable<ReaderSourceDocument> ReadFolderDocumentsCore(
        string folderPath,
        ReaderFolderOptions? folderOptions,
        ReaderOptions? options,
        bool includeSkippedWarningChunks,
        Action<ReaderProgress>? onProgress,
        CancellationToken cancellationToken) {
        if (folderPath == null) throw new ArgumentNullException(nameof(folderPath));
        if (!Directory.Exists(folderPath)) throw new DirectoryNotFoundException($"Directory '{folderPath}' does not exist.");
        ReaderFolderOptions effectiveFolder = NormalizeFolderOptions(folderOptions);
        HashSet<string> allowedExtensions = NormalizeExtensions(effectiveFolder.Extensions);
        var state = new FolderIngestState();
        foreach (string file in EnumerateFilesSafeDeterministic(folderPath, effectiveFolder, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (state.FilesScanned >= effectiveFolder.MaxFiles) break;
            string extension = NormalizeExtension(TryGetExtension(file));
            if (!allowedExtensions.Contains(extension) &&
                !string.Equals(Path.GetFileName(file), "winmail.dat", StringComparison.OrdinalIgnoreCase)) continue;
            state.FilesScanned++;
            SourceInfo source = BuildSourceInfoFromPath(file, computeHash: false, cancellationToken);
            NotifyProgress(onProgress, ReaderProgressEventKind.FileStarted, state, source, null, null);
            ReaderSourceDocument document = ReadSingleDocument(file, options, cancellationToken);
            if (document.Parsed) {
                long nextBytes = state.BytesRead + (document.SourceLengthBytes ?? 0);
                if (effectiveFolder.MaxTotalBytes.HasValue && nextBytes > effectiveFolder.MaxTotalBytes.Value) {
                    document = BuildSourceDocument(source, false, null, new[] { "Skipped because MaxTotalBytes would be exceeded." });
                }
            }
            if (!document.Parsed && includeSkippedWarningChunks) {
                string warning = document.Warnings?.FirstOrDefault() ?? "Skipped file because it could not be read.";
                ReaderChunk warningChunk = EnrichChunk(BuildFolderWarningChunk(file, 0, warning), source, computeHashes: false);
                document = BuildSourceDocument(source, false, new[] { warningChunk }, document.Warnings);
            }
            if (document.Parsed) {
                state.FilesParsed++;
                state.BytesRead += document.SourceLengthBytes ?? 0;
                state.ChunksProduced += document.ChunksProduced;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileCompleted, state, source, null, document.ChunksProduced);
            } else {
                state.FilesSkipped++;
                state.ChunksProduced += document.ChunksProduced;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source,
                    document.Warnings?.FirstOrDefault(), document.ChunksProduced);
            }
            yield return document;
        }
        NotifyProgress(onProgress, ReaderProgressEventKind.Completed, state, null, null, null);
    }

    private static ReaderSourceDocument ReadSingleDocument(string path, ReaderOptions? options, CancellationToken cancellationToken) {
        SourceInfo source = BuildSourceInfoFromPath(path, computeHash: false, cancellationToken);
        try {
            ReaderChunk[] chunks = Read(path, options, cancellationToken).ToArray();
            ReaderChunk? first = chunks.FirstOrDefault();
            source.SourceHash = first?.SourceHash;
            return BuildSourceDocument(source, true, chunks, null);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception exception) {
            string warning = $"Skipped file due read error: {exception.GetType().Name}. {exception.Message}";
            return BuildSourceDocument(source, false, null, new[] { warning });
        }
    }

    private static IEnumerable<string> EnumerateFilesSafeDeterministic(string folderPath, ReaderFolderOptions options, CancellationToken cancellationToken) {
        var directories = new Queue<string>();
        directories.Enqueue(folderPath);
        while (directories.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            string directory = directories.Dequeue();
            string[] entries;
            try { entries = Directory.GetFileSystemEntries(directory); } catch { continue; }
            if (options.DeterministicOrder) Array.Sort(entries, StringComparer.Ordinal);
            foreach (string entry in entries) {
                FileAttributes attributes;
                try { attributes = File.GetAttributes(entry); } catch { continue; }
                if ((attributes & FileAttributes.Directory) != 0) {
                    if (options.Recurse && (!options.SkipReparsePoints || (attributes & FileAttributes.ReparsePoint) == 0)) directories.Enqueue(entry);
                } else {
                    yield return entry;
                }
            }
        }
    }

    private static ReaderSourceDocument CloneSourceDocument(
        ReaderSourceDocument source, IReadOnlyList<ReaderChunk> chunks, string? warning = null) => new ReaderSourceDocument {
        Path = source.Path,
        SourceId = source.SourceId,
        SourceHash = source.SourceHash,
        SourceLastWriteUtc = source.SourceLastWriteUtc,
        SourceLengthBytes = source.SourceLengthBytes,
        Parsed = source.Parsed,
        ChunksProduced = source.ChunksProduced,
        TokenEstimateTotal = source.TokenEstimateTotal,
        Warnings = string.IsNullOrWhiteSpace(warning)
            ? source.Warnings
            : (source.Warnings ?? Array.Empty<string>()).Concat(new[] { warning! }).Distinct(StringComparer.Ordinal).ToArray(),
        Chunks = chunks
    };

    private static IReadOnlyList<string>? MergeWarnings(IEnumerable<ReaderSourceDocument> documents) {
        string[] warnings = documents.SelectMany(static document => document.Warnings ?? Array.Empty<string>())
            .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        return warnings.Length == 0 ? null : warnings;
    }

    private static void NotifyProgress(Action<ReaderProgress>? callback, ReaderProgressEventKind kind, FolderIngestState state, SourceInfo? source, string? message, int? fileChunks) {
        callback?.Invoke(new ReaderProgress {
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
            CurrentFileChunks = fileChunks,
            CurrentFileLastWriteUtc = source?.LastWriteUtc
        });
    }
}
