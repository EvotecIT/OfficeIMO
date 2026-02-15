using OfficeIMO.Excel;
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
using System.Text;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Unified, read-only document extraction facade intended for AI ingestion.
/// </summary>
/// <remarks>
/// This facade is intentionally dependency-free and deterministic.
/// It normalizes extraction into <see cref="ReaderChunk"/> instances with stable IDs and location metadata.
/// The API is thread-safe as it does not use shared mutable state.
/// </remarks>
public static class DocumentReader {
    private static readonly string[] DefaultFolderExtensions = {
        ".docx", ".docm",
        ".xlsx", ".xlsm",
        ".pptx", ".pptm",
        ".md", ".markdown",
        ".pdf",
        ".txt", ".log", ".csv", ".tsv", ".json", ".xml", ".yml", ".yaml"
    };

    private static string? TryGetExtension(string path) {
        if (path == null) return null;
        try {
            return Path.GetExtension(path);
        } catch (ArgumentException) {
            return null;
        } catch (NotSupportedException) {
            return null;
        }
    }

    /// <summary>
    /// Detects the input kind based on file extension.
    /// </summary>
    /// <param name="path">Source file path.</param>
    public static ReaderInputKind DetectKind(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var extLower = (TryGetExtension(path) ?? string.Empty).ToLowerInvariant();
        if (extLower.Length == 0) return ReaderInputKind.Unknown;
        return extLower switch {
            ".docx" or ".docm" => ReaderInputKind.Word,
            ".xlsx" or ".xlsm" => ReaderInputKind.Excel,
            ".pptx" or ".pptm" => ReaderInputKind.PowerPoint,
            ".md" or ".markdown" => ReaderInputKind.Markdown,
            ".pdf" => ReaderInputKind.Pdf,
            ".txt" or ".log" or ".csv" or ".tsv" or ".json" or ".xml" or ".yml" or ".yaml" => ReaderInputKind.Text,
            ".doc" or ".xls" or ".ppt" => ReaderInputKind.Unknown, // Legacy binary formats are not supported.
            _ => ReaderInputKind.Unknown
        };
    }

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
        EnforceFileSize(path, opt.MaxInputBytes);

        var kind = DetectKind(path);
        return kind switch {
            ReaderInputKind.Word => ReadWord(path, opt, cancellationToken),
            ReaderInputKind.Excel => ReadExcel(path, opt, cancellationToken),
            ReaderInputKind.PowerPoint => ReadPowerPoint(path, opt, cancellationToken),
            ReaderInputKind.Markdown => ReadMarkdown(path, opt, cancellationToken),
            ReaderInputKind.Pdf => ReadPdf(path, opt, cancellationToken),
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
        if (folderPath == null) throw new ArgumentNullException(nameof(folderPath));
        if (!Directory.Exists(folderPath)) throw new DirectoryNotFoundException($"Folder '{folderPath}' doesn't exist.");

        var fo = NormalizeFolderOptions(folderOptions);
        var opt = NormalizeOptions(options);
        var allowedExt = NormalizeExtensions(fo.Extensions);
        long total = 0;
        int count = 0;
        int warningIndex = 0;

        foreach (var file in EnumerateFilesSafeDeterministic(folderPath, fo, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();

            if (count >= fo.MaxFiles) yield break;
            var ext = TryGetExtension(file);
            if (string.IsNullOrEmpty(ext)) continue;
            if (!allowedExt.Contains(ext!)) continue;

            long length = 0;
            string? statWarning = null;
            try {
                length = new FileInfo(file).Length;
            } catch {
                statWarning = "Skipped file because metadata could not be read.";
            }
            if (statWarning != null) {
                yield return BuildFolderWarningChunk(file, warningIndex++, statWarning);
                continue;
            }

            if (fo.MaxTotalBytes.HasValue) {
                if ((total + length) > fo.MaxTotalBytes.Value) {
                    yield return BuildFolderWarningChunk(
                        file,
                        warningIndex++,
                        $"Stopped folder ingestion after reaching MaxTotalBytes ({fo.MaxTotalBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
                    yield break;
                }
            }
            total += length;

            if (opt.MaxInputBytes.HasValue && length > opt.MaxInputBytes.Value) {
                // Skip too-large files rather than failing the whole folder.
                yield return BuildFolderWarningChunk(
                    file,
                    warningIndex++,
                    $"Skipped file because it exceeds MaxInputBytes ({length.ToString(CultureInfo.InvariantCulture)} > {opt.MaxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
                continue;
            }

            count++;
            List<ReaderChunk>? fileChunks = null;
            string? readWarning = null;
            try {
                fileChunks = Read(file, opt, cancellationToken).ToList();
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                // Keep folder ingestion best-effort; skip files that fail parsing.
                readWarning = $"Skipped file due read error: {ex.GetType().Name}.";
            }
            if (readWarning != null) {
                yield return BuildFolderWarningChunk(file, warningIndex++, readWarning);
                continue;
            }

            foreach (var chunk in fileChunks!) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return chunk;
            }
        }
    }

    /// <summary>
    /// Reads a supported document from a stream and emits normalized extraction chunks.
    /// </summary>
    /// <param name="stream">Source stream. This method does not close the stream.</param>
    /// <param name="sourceName">
    /// Optional source name used for kind detection (via extension) and citations/IDs.
    /// For example: "Policy.docx" or "Workbook.xlsx".
    /// </param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderChunk> Read(Stream stream, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var opt = NormalizeOptions(options);
        EnforceStreamSize(stream, opt.MaxInputBytes);

        var kind = string.IsNullOrWhiteSpace(sourceName) ? ReaderInputKind.Unknown : DetectKind(sourceName!);
        return kind switch {
            ReaderInputKind.Word => ReadWord(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.Excel => ReadExcel(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.PowerPoint => ReadPowerPoint(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.Markdown => ReadMarkdown(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.Pdf => ReadPdf(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.Text => ReadText(stream, sourceName, opt, cancellationToken),
            _ => ReadUnknown(stream, sourceName, opt, cancellationToken)
        };
    }

    /// <summary>
    /// Reads a supported document from bytes and emits normalized extraction chunks.
    /// </summary>
    /// <param name="bytes">Source bytes.</param>
    /// <param name="sourceName">
    /// Optional source name used for kind detection (via extension) and citations/IDs.
    /// For example: "Policy.docx" or "Workbook.xlsx".
    /// </param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderChunk> Read(byte[] bytes, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var ms = new MemoryStream(bytes, writable: false);
        foreach (var c in Read(ms, sourceName, options, cancellationToken))
            yield return c;
    }

    private static IEnumerable<ReaderChunk> ReadWord(string path, ReaderOptions opt, CancellationToken ct) {
        using var doc = WordDocument.Load(path, readOnly: true, autoSave: false, openSettings: CreateOpenSettings(opt));
        var chunks = doc.ExtractMarkdownChunks(
            markdownOptions: new WordToMarkdownOptions(),
            chunking: new WordMarkdownChunkingOptions { MaxChars = opt.MaxChars, IncludeFootnotes = opt.IncludeWordFootnotes },
            sourcePath: path,
            cancellationToken: ct);

        int outIndex = 0;
        foreach (var c in chunks) {
            ct.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = c.Id,
                Kind = ReaderInputKind.Word,
                Location = new ReaderLocation {
                    Path = c.Location.Path,
                    BlockIndex = outIndex,
                    SourceBlockIndex = c.Location.BlockIndex,
                    HeadingPath = c.Location.HeadingPath
                },
                Text = c.Text,
                Markdown = c.Markdown,
                Warnings = c.Warnings
            };
            outIndex++;
        }
    }

    private static IEnumerable<ReaderChunk> ReadWord(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // Copy input so we can open read-only without affecting caller's stream.
        using var ms = CopyToMemory(stream, ct);
        using var doc = WordDocument.Load(ms, readOnly: true, autoSave: false, openSettings: CreateOpenSettings(opt));

        var chunks = doc.ExtractMarkdownChunks(
            markdownOptions: new WordToMarkdownOptions(),
            chunking: new WordMarkdownChunkingOptions { MaxChars = opt.MaxChars, IncludeFootnotes = opt.IncludeWordFootnotes },
            sourcePath: sourceName,
            cancellationToken: ct);

        int outIndex = 0;
        foreach (var c in chunks) {
            ct.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = c.Id,
                Kind = ReaderInputKind.Word,
                Location = new ReaderLocation {
                    Path = sourceName,
                    BlockIndex = outIndex,
                    SourceBlockIndex = c.Location.BlockIndex,
                    HeadingPath = c.Location.HeadingPath
                },
                Text = c.Text,
                Markdown = c.Markdown,
                Warnings = c.Warnings
            };
            outIndex++;
        }
    }

    private static IEnumerable<ReaderChunk> ReadExcel(string path, ReaderOptions opt, CancellationToken ct) {
        // Use OpenSettings for basic OpenXML hardening (best-effort) and open from stream to avoid file handle collisions.
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var openSettings = CreateOpenSettings(opt);
        using var openXml = openSettings == null
            ? SpreadsheetDocument.Open(fs, false)
            : SpreadsheetDocument.Open(fs, false, openSettings);
        using var reader = ExcelDocumentReader.Wrap(openXml);
        var sheets = ResolveSheetNames(reader, opt.ExcelSheetName);

        int outIndex = 0;
        foreach (var sheet in sheets) {
            ct.ThrowIfCancellationRequested();

            var chunks = reader.ExtractChunks(
                sheetName: sheet,
                a1Range: opt.ExcelA1Range,
                extract: new ExcelExtractionExtensions.ExcelExtractOptions {
                    HeadersInFirstRow = opt.ExcelHeadersInFirstRow,
                    ChunkRows = opt.ExcelChunkRows,
                    EmitMarkdownTable = true
                },
                chunking: new ExcelExtractChunkingOptions { MaxChars = opt.MaxChars, MaxTableRows = opt.MaxTableRows },
                sourcePath: path,
                cancellationToken: ct);

            foreach (var c in chunks) {
                ct.ThrowIfCancellationRequested();

                IReadOnlyList<ReaderTable>? tables = null;
                if (c.Tables != null && c.Tables.Count > 0) {
                    tables = c.Tables.Select(MapTable).ToArray();
                }

                yield return new ReaderChunk {
                    Id = c.Id,
                    Kind = ReaderInputKind.Excel,
                    Location = new ReaderLocation {
                        Path = c.Location.Path,
                        Sheet = c.Location.Sheet,
                        A1Range = c.Location.A1Range,
                        BlockIndex = outIndex,
                        SourceBlockIndex = c.Location.BlockIndex
                    },
                    Text = c.Text,
                    Markdown = c.Markdown,
                    Tables = tables,
                    Warnings = c.Warnings
                };
                outIndex++;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadExcel(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // Avoid exposing OpenXml types in the public API surface; internally we can wrap.
        using var ms = CopyToMemory(stream, ct);
        var openSettings = CreateOpenSettings(opt);
        using var openXml = openSettings == null
            ? SpreadsheetDocument.Open(ms, false)
            : SpreadsheetDocument.Open(ms, false, openSettings);
        using var reader = ExcelDocumentReader.Wrap(openXml);

        var sheets = ResolveSheetNames(reader, opt.ExcelSheetName);

        int outIndex = 0;
        foreach (var sheet in sheets) {
            ct.ThrowIfCancellationRequested();

            var chunks = reader.ExtractChunks(
                sheetName: sheet,
                a1Range: opt.ExcelA1Range,
                extract: new ExcelExtractionExtensions.ExcelExtractOptions {
                    HeadersInFirstRow = opt.ExcelHeadersInFirstRow,
                    ChunkRows = opt.ExcelChunkRows,
                    EmitMarkdownTable = true
                },
                chunking: new ExcelExtractChunkingOptions { MaxChars = opt.MaxChars, MaxTableRows = opt.MaxTableRows },
                sourcePath: sourceName,
                cancellationToken: ct);

            foreach (var c in chunks) {
                ct.ThrowIfCancellationRequested();

                IReadOnlyList<ReaderTable>? tables = null;
                if (c.Tables != null && c.Tables.Count > 0) {
                    tables = c.Tables.Select(MapTable).ToArray();
                }

                yield return new ReaderChunk {
                    Id = c.Id,
                    Kind = ReaderInputKind.Excel,
                    Location = new ReaderLocation {
                        Path = sourceName,
                        Sheet = c.Location.Sheet,
                        A1Range = c.Location.A1Range,
                        BlockIndex = outIndex,
                        SourceBlockIndex = c.Location.BlockIndex
                    },
                    Text = c.Text,
                    Markdown = c.Markdown,
                    Tables = tables,
                    Warnings = c.Warnings
                };
                outIndex++;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadPowerPoint(string path, ReaderOptions opt, CancellationToken ct) {
        using var presentation = PowerPointPresentation.OpenRead(path);
        var chunks = presentation.ExtractMarkdownChunks(
            extract: new PowerPointExtractionExtensions.PowerPointExtractOptions { IncludeNotes = opt.IncludePowerPointNotes },
            chunking: new PowerPointExtractChunkingOptions { MaxChars = opt.MaxChars },
            sourcePath: path,
            cancellationToken: ct);

        int outIndex = 0;
        foreach (var c in chunks) {
            ct.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = c.Id,
                Kind = ReaderInputKind.PowerPoint,
                Location = new ReaderLocation {
                    Path = c.Location.Path,
                    Slide = c.Location.Slide,
                    BlockIndex = outIndex,
                    SourceBlockIndex = c.Location.BlockIndex
                },
                Text = c.Text,
                Markdown = c.Markdown,
                Warnings = c.Warnings
            };
            outIndex++;
        }
    }

    private static IEnumerable<ReaderChunk> ReadPowerPoint(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // PowerPointPresentation.Open(stream, readOnly:true) already copies to an internal stream for safety.
        using var presentation = PowerPointPresentation.Open(stream, readOnly: true, autoSave: false);
        var chunks = presentation.ExtractMarkdownChunks(
            extract: new PowerPointExtractionExtensions.PowerPointExtractOptions { IncludeNotes = opt.IncludePowerPointNotes },
            chunking: new PowerPointExtractChunkingOptions { MaxChars = opt.MaxChars },
            sourcePath: sourceName,
            cancellationToken: ct);

        int outIndex = 0;
        foreach (var c in chunks) {
            ct.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = c.Id,
                Kind = ReaderInputKind.PowerPoint,
                Location = new ReaderLocation {
                    Path = sourceName,
                    Slide = c.Location.Slide,
                    BlockIndex = outIndex,
                    SourceBlockIndex = c.Location.BlockIndex
                },
                Text = c.Text,
                Markdown = c.Markdown,
                Warnings = c.Warnings
            };
            outIndex++;
        }
    }

    private static IEnumerable<ReaderChunk> ReadPdf(string path, ReaderOptions opt, CancellationToken ct) {
        var fileName = Path.GetFileName(path);
        var doc = PdfReadDocument.Load(path);
        int outIndex = 0;

        for (int pageIndex = 0; pageIndex < doc.Pages.Count; pageIndex++) {
            ct.ThrowIfCancellationRequested();

            var pageNumber = pageIndex + 1;
            var pageText = doc.Pages[pageIndex].ExtractText();
            if (string.IsNullOrWhiteSpace(pageText)) {
                yield return BuildPdfEmptyChunk(path, fileName, pageNumber, outIndex);
                outIndex++;
                continue;
            }

            var pageChunks = ChunkPdfText(path, fileName, pageNumber, pageText, opt, outIndex, ct, out var nextIndex);
            outIndex = nextIndex;
            foreach (var chunk in pageChunks) {
                yield return chunk;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadPdf(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        using var ms = CopyToMemory(stream, ct);
        var fileName = string.IsNullOrWhiteSpace(sourceName) ? "memory.pdf" : Path.GetFileName(sourceName!.Trim());
        var doc = PdfReadDocument.Load(ms.ToArray());
        int outIndex = 0;

        for (int pageIndex = 0; pageIndex < doc.Pages.Count; pageIndex++) {
            ct.ThrowIfCancellationRequested();

            var pageNumber = pageIndex + 1;
            var pageText = doc.Pages[pageIndex].ExtractText();
            if (string.IsNullOrWhiteSpace(pageText)) {
                yield return BuildPdfEmptyChunk(sourceName ?? fileName, fileName, pageNumber, outIndex);
                outIndex++;
                continue;
            }

            var pageChunks = ChunkPdfText(sourceName ?? fileName, fileName, pageNumber, pageText, opt, outIndex, ct, out var nextIndex);
            outIndex = nextIndex;
            foreach (var chunk in pageChunks) {
                yield return chunk;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadMarkdown(string path, ReaderOptions opt, CancellationToken ct) {
        // Keep it simple: chunk by headings (ATX, best-effort), with size cap.
        if (!opt.MarkdownChunkByHeadings) {
            foreach (var c in ChunkPlainTextByParagraphs(path, opt, ReaderInputKind.Markdown, ct, treatAsMarkdown: true))
                yield return c;
            yield break;
        }

        var fileName = Path.GetFileName(path);
        var headingStack = new List<(int Level, string Text)>();

        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        int chunkIndex = 0;
        int? firstLine = null;
        string? firstHeadingPath = null;
        var warnings = new List<string>(capacity: 2);

        int lineNo = 0;
        foreach (var line in File.ReadLines(path)) {
            ct.ThrowIfCancellationRequested();
            lineNo++;

            if (TryParseAtxHeading(line, out var level, out var headingText)) {
                // Flush current section before starting a new heading section.
                if (current.Length > 0) {
                    yield return BuildMarkdownChunk(path, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
                    chunkIndex++;
                    current.Clear();
                    warnings.Clear();
                    firstLine = null;
                    firstHeadingPath = null;
                }

                UpdateHeadingStack(headingStack, level, headingText);
                var headingPath = BuildHeadingPath(headingStack);
                firstHeadingPath = headingPath;
                firstLine = lineNo;

                // Keep the heading line as part of the new chunk content.
                AppendLineCapped(opt, current, line, warnings);
                continue;
            }

            if (firstLine == null) firstLine = lineNo;
            if (firstHeadingPath == null) firstHeadingPath = BuildHeadingPath(headingStack);

            // If adding this line would exceed MaxChars, flush a chunk boundary.
            if (WouldExceed(opt, current, line)) {
                yield return BuildMarkdownChunk(path, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = lineNo;
                firstHeadingPath = BuildHeadingPath(headingStack);
            }

            AppendLineCapped(opt, current, line, warnings);
        }

        if (current.Length > 0) {
            yield return BuildMarkdownChunk(path, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
        }
    }

    private static IEnumerable<ReaderChunk> ReadMarkdown(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        var fileName = string.IsNullOrWhiteSpace(sourceName) ? "memory.md" : Path.GetFileName(sourceName!.Trim());
        var text = ReadAllText(stream, ct);
        foreach (var c in ChunkMarkdownFromText(text, sourceName, fileName, opt, ct))
            yield return c;
    }

    private static IEnumerable<ReaderChunk> ReadText(string path, ReaderOptions opt, CancellationToken ct) {
        foreach (var c in ChunkPlainTextByParagraphs(path, opt, ReaderInputKind.Text, ct, treatAsMarkdown: false))
            yield return c;
    }

    private static IEnumerable<ReaderChunk> ReadText(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        var fileName = string.IsNullOrWhiteSpace(sourceName) ? "memory.txt" : Path.GetFileName(sourceName!.Trim());
        var text = ReadAllText(stream, ct);
        foreach (var c in ChunkPlainTextFromText(text, sourceName, fileName, opt, ReaderInputKind.Text, ct, treatAsMarkdown: false))
            yield return c;
    }

    private static IEnumerable<ReaderChunk> ReadUnknown(string path, ReaderOptions opt, CancellationToken ct) {
        var extLower = (TryGetExtension(path) ?? string.Empty).ToLowerInvariant();
        if (extLower is ".doc" or ".xls" or ".ppt") {
            throw new NotSupportedException($"Legacy binary format '{extLower}' is not supported. Convert to OpenXML (.docx/.xlsx/.pptx) first.");
        }

        // Try plain text; if it fails (binary), the caller can decide how to handle it.
        foreach (var c in ChunkPlainTextByParagraphs(path, opt, ReaderInputKind.Unknown, ct, treatAsMarkdown: false))
            yield return c;
    }

    private static IEnumerable<ReaderChunk> ReadUnknown(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // When we can't detect kind, treat as plain text.
        var fileName = string.IsNullOrWhiteSpace(sourceName) ? "memory" : Path.GetFileName(sourceName!.Trim());
        var text = ReadAllText(stream, ct);
        foreach (var c in ChunkPlainTextFromText(text, sourceName, fileName, opt, ReaderInputKind.Unknown, ct, treatAsMarkdown: false))
            yield return c;
    }

    private static IReadOnlyList<string> ResolveSheetNames(ExcelDocumentReader reader, string? singleSheet) {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        if (!string.IsNullOrWhiteSpace(singleSheet)) return new[] { singleSheet!.Trim() };
        return reader.GetSheetNames();
    }

    private static ReaderTable MapTable(ExcelExtractTable t) {
        return new ReaderTable {
            Title = t.Title,
            Columns = t.Columns,
            Rows = t.Rows,
            TotalRowCount = t.TotalRowCount,
            Truncated = t.Truncated
        };
    }

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

        var headingStack = new List<(int Level, string Text)>();
        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        int chunkIndex = 0;
        int? firstLine = null;
        string? firstHeadingPath = null;
        var warnings = new List<string>(capacity: 2);

        int lineNo = 0;
        using var sr = new StringReader(text ?? string.Empty);
        string? line;
        while ((line = sr.ReadLine()) != null) {
            ct.ThrowIfCancellationRequested();
            lineNo++;

            if (TryParseAtxHeading(line, out var level, out var headingText)) {
                if (current.Length > 0) {
                    yield return BuildMarkdownChunk(sourceName ?? fileName, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
                    chunkIndex++;
                    current.Clear();
                    warnings.Clear();
                    firstLine = null;
                    firstHeadingPath = null;
                }

                UpdateHeadingStack(headingStack, level, headingText);
                var headingPath = BuildHeadingPath(headingStack);
                firstHeadingPath = headingPath;
                firstLine = lineNo;

                AppendLineCapped(opt, current, line, warnings);
                continue;
            }

            if (firstLine == null) firstLine = lineNo;
            if (firstHeadingPath == null) firstHeadingPath = BuildHeadingPath(headingStack);

            if (WouldExceed(opt, current, line)) {
                yield return BuildMarkdownChunk(sourceName ?? fileName, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = lineNo;
                firstHeadingPath = BuildHeadingPath(headingStack);
            }

            AppendLineCapped(opt, current, line, warnings);
        }

        if (current.Length > 0) {
            yield return BuildMarkdownChunk(sourceName ?? fileName, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
        }
    }

    private static List<ReaderChunk> ChunkPdfText(
        string path,
        string fileName,
        int pageNumber,
        string text,
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
        nextChunkIndex = outIndex;
        return list;
    }

    private static ReaderChunk BuildMarkdownChunk(
        string path,
        string fileName,
        int chunkIndex,
        int? firstLine,
        string? headingPath,
        string markdown,
        List<string> warnings) {
        var id = BuildStableId("md", fileName, chunkIndex, firstLine);
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Markdown,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = chunkIndex,
                StartLine = firstLine,
                HeadingPath = headingPath
            },
            Text = markdown,
            Markdown = markdown,
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
        var id = BuildStableId(kind == ReaderInputKind.Text ? "text" : "unknown", fileName, chunkIndex, firstLine);
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

    private static bool TryParseAtxHeading(string line, out int level, out string text) {
        level = 0;
        text = string.Empty;
        if (line == null) return false;

        int i = 0;
        while (i < line.Length && line[i] == '#') i++;
        if (i < 1 || i > 6) return false;
        if (i >= line.Length) return false;
        if (line[i] != ' ' && line[i] != '\t') return false;

        level = i;
        text = line.Substring(i).Trim();
        if (text.Length == 0) text = $"Heading {level}";
        return true;
    }

    private static void UpdateHeadingStack(List<(int Level, string Text)> stack, int level, string text) {
        if (level < 1) return;
        if (string.IsNullOrWhiteSpace(text)) text = $"Heading {level}";

        for (int i = stack.Count - 1; i >= 0; i--) {
            if (stack[i].Level >= level) stack.RemoveAt(i);
        }
        stack.Add((level, CollapseWhitespace(text)));
    }

    private static string? BuildHeadingPath(List<(int Level, string Text)> stack) {
        if (stack.Count == 0) return null;
        var sb = new StringBuilder();
        for (int i = 0; i < stack.Count; i++) {
            if (i > 0) sb.Append(" > ");
            sb.Append(stack[i].Text);
        }
        var s = sb.ToString().Trim();
        return s.Length == 0 ? null : s;
    }

    private static bool WouldExceed(ReaderOptions opt, StringBuilder current, string nextLine) {
        // +1 for newline to keep final chunk shape similar to file.
        int nextLen = nextLine?.Length ?? 0;
        int extra = (current.Length == 0 ? 0 : 1) + nextLen;
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

    private static string CollapseWhitespace(string text) {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var sb = new StringBuilder(text.Length);
        bool prevWs = false;
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            bool ws = char.IsWhiteSpace(c);
            if (ws) {
                if (!prevWs) sb.Append(' ');
                prevWs = true;
            } else {
                sb.Append(c);
                prevWs = false;
            }
        }
        return sb.ToString().Trim();
    }

    private static string BuildStableId(string kind, string fileName, int chunkIndex, int? blockIndex) {
        // Keep IDs short, stable and ASCII-only; do not leak full paths.
        var l = blockIndex.HasValue ? blockIndex.Value.ToString(CultureInfo.InvariantCulture) : "na";
        return $"{kind}:{fileName}:c{chunkIndex}:l{l}";
    }

    private static MemoryStream CopyToMemory(Stream stream, CancellationToken ct) {
        ct.ThrowIfCancellationRequested();
        var ms = new MemoryStream();
        var buffer = new byte[64 * 1024];
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
            ct.ThrowIfCancellationRequested();
            ms.Write(buffer, 0, read);
        }
        ms.Position = 0;
        return ms;
    }

    private static ReaderOptions NormalizeOptions(ReaderOptions? options) {
        // Avoid mutating a caller-provided options instance.
        var o = options;
        var clone = new ReaderOptions {
            MaxInputBytes = o?.MaxInputBytes,
            OpenXmlMaxCharactersInPart = o?.OpenXmlMaxCharactersInPart,
            MaxChars = o?.MaxChars ?? 8_000,
            MaxTableRows = o?.MaxTableRows ?? 200,
            IncludeWordFootnotes = o?.IncludeWordFootnotes ?? true,
            IncludePowerPointNotes = o?.IncludePowerPointNotes ?? true,
            ExcelHeadersInFirstRow = o?.ExcelHeadersInFirstRow ?? true,
            ExcelChunkRows = o?.ExcelChunkRows ?? 200,
            ExcelSheetName = o?.ExcelSheetName,
            ExcelA1Range = o?.ExcelA1Range,
            MarkdownChunkByHeadings = o?.MarkdownChunkByHeadings ?? true
        };

        if (clone.MaxChars < 256) clone.MaxChars = 256;
        if (clone.MaxTableRows < 1) clone.MaxTableRows = 1;
        if (clone.ExcelChunkRows < 1) clone.ExcelChunkRows = 1;
        if (clone.OpenXmlMaxCharactersInPart.HasValue && clone.OpenXmlMaxCharactersInPart.Value < 1) clone.OpenXmlMaxCharactersInPart = null;

        return clone;
    }

    private static ReaderFolderOptions NormalizeFolderOptions(ReaderFolderOptions? folderOptions) {
        var o = folderOptions;
        var clone = new ReaderFolderOptions {
            Recurse = o?.Recurse ?? true,
            MaxFiles = o?.MaxFiles ?? 500,
            MaxTotalBytes = o?.MaxTotalBytes,
            Extensions = (o?.Extensions == null || o.Extensions.Count == 0) ? null : o.Extensions.ToArray(),
            SkipReparsePoints = o?.SkipReparsePoints ?? true,
            DeterministicOrder = o?.DeterministicOrder ?? true
        };

        if (clone.MaxFiles < 1) clone.MaxFiles = 1;
        if (clone.MaxTotalBytes.HasValue && clone.MaxTotalBytes.Value < 1) clone.MaxTotalBytes = 1;
        return clone;
    }

    private static HashSet<string> NormalizeExtensions(IReadOnlyList<string>? configuredExtensions) {
        var allowedExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var source = (configuredExtensions == null || configuredExtensions.Count == 0)
            ? DefaultFolderExtensions
            : configuredExtensions;

        foreach (var e in source) {
            if (string.IsNullOrWhiteSpace(e)) continue;
            var normalized = e.StartsWith(".", StringComparison.Ordinal) ? e.Trim() : "." + e.Trim();
            if (normalized.Length > 1) allowedExt.Add(normalized);
        }

        return allowedExt;
    }

    private static IEnumerable<string> EnumerateFilesSafeDeterministic(string folderPath, ReaderFolderOptions options, CancellationToken cancellationToken) {
        var dirs = new Queue<string>();
        dirs.Enqueue(folderPath);

        while (dirs.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var dir = dirs.Dequeue();

            IEnumerable<string> entries;
            try {
                entries = Directory.EnumerateFileSystemEntries(dir);
            } catch {
                // Best-effort traversal: unreadable directories are ignored.
                continue;
            }

            var ordered = options.DeterministicOrder
                ? entries.OrderBy(static x => x, StringComparer.Ordinal).ToArray()
                : entries.ToArray();

            foreach (var entry in ordered) {
                cancellationToken.ThrowIfCancellationRequested();

                FileAttributes attrs;
                try {
                    attrs = File.GetAttributes(entry);
                } catch {
                    continue;
                }

                var isDirectory = (attrs & FileAttributes.Directory) == FileAttributes.Directory;
                if (isDirectory) {
                    if (!options.Recurse) continue;

                    if (options.SkipReparsePoints && (attrs & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint) {
                        continue;
                    }

                    dirs.Enqueue(entry);
                    continue;
                }

                yield return entry;
            }
        }
    }

    private static OpenSettings? CreateOpenSettings(ReaderOptions opt) {
        if (opt == null) return null;
        if (!opt.OpenXmlMaxCharactersInPart.HasValue) return null;
        return new OpenSettings {
            MaxCharactersInPart = opt.OpenXmlMaxCharactersInPart.Value
        };
    }

    private static void EnforceFileSize(string path, long? maxBytes) {
        if (!maxBytes.HasValue) return;
        try {
            var fi = new FileInfo(path);
            if (fi.Length > maxBytes.Value) {
                throw new IOException($"Input exceeds MaxInputBytes ({fi.Length.ToString(CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }
        } catch (IOException) {
            throw;
        } catch {
            // If we can't stat, don't block reads.
        }
    }

    private static void EnforceStreamSize(Stream stream, long? maxBytes) {
        if (!maxBytes.HasValue) return;
        if (!stream.CanSeek) return;
        try {
            if (stream.Length > maxBytes.Value) {
                throw new IOException($"Input exceeds MaxInputBytes ({stream.Length.ToString(CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }
        } catch (NotSupportedException) {
            // ignore
        }
    }

    private static string ReadAllText(Stream stream, CancellationToken ct) {
        ct.ThrowIfCancellationRequested();
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: true);
        var sb = new StringBuilder();
        var buffer = new char[16 * 1024];
        const int HardCapChars = 50_000_000; // Defensive: avoid runaway memory usage on huge "text" streams.
        int read;
        while ((read = reader.Read(buffer, 0, buffer.Length)) > 0) {
            ct.ThrowIfCancellationRequested();
            sb.Append(buffer, 0, read);
            if (sb.Length >= HardCapChars) break;
        }
        return sb.ToString();
    }
}
