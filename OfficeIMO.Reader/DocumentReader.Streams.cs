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

public static partial class DocumentReader {
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
        string? logicalSourceName = null;
        if (sourceName != null) {
            var trimmedSourceName = sourceName.Trim();
            if (trimmedSourceName.Length > 0) {
                logicalSourceName = trimmedSourceName;
            }
        }
        var source = BuildSourceInfoFromStream(stream, logicalSourceName, opt.ComputeHashes);

        IEnumerable<ReaderChunk> raw;
        if (TryResolveCustomHandlerBySourceName(logicalSourceName, out var customStreamHandler) && customStreamHandler.ReadStream != null) {
            raw = customStreamHandler.ReadStream(stream, logicalSourceName, opt, cancellationToken);
        } else {
            var kind = string.IsNullOrWhiteSpace(logicalSourceName) ? ReaderInputKind.Unknown : DetectKind(logicalSourceName!);
            raw = kind switch {
                ReaderInputKind.Word => ReadWord(stream, logicalSourceName, opt, cancellationToken),
                ReaderInputKind.Excel => ReadExcel(stream, logicalSourceName, opt, cancellationToken),
                ReaderInputKind.PowerPoint => ReadPowerPoint(stream, logicalSourceName, opt, cancellationToken),
                ReaderInputKind.Markdown => ReadMarkdown(stream, logicalSourceName, opt, cancellationToken),
                ReaderInputKind.Pdf => ReadPdf(stream, logicalSourceName, opt, cancellationToken),
                ReaderInputKind.Text => ReadText(stream, logicalSourceName, opt, cancellationToken),
                _ => ReadUnknown(stream, logicalSourceName, opt, cancellationToken)
            };
        }

        foreach (var chunk in raw) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return EnrichChunk(chunk, source, opt.ComputeHashes);
        }
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
        int tableIndex = 0;
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
                    tables = MapTables(c.Tables, c.Location, ref tableIndex);
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
        int tableIndex = 0;
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
                    tables = MapTables(c.Tables, c.Location, ref tableIndex);
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
        var doc = PdfLogicalDocument.Load(path);
        int outIndex = 0;

        for (int pageIndex = 0; pageIndex < doc.Pages.Count; pageIndex++) {
            ct.ThrowIfCancellationRequested();

            var page = doc.Pages[pageIndex];
            var pageNumber = page.PageNumber;
            var pageText = BuildPdfPageText(page);
            if (string.IsNullOrWhiteSpace(pageText)) {
                yield return BuildPdfEmptyChunk(path, fileName, pageNumber, outIndex);
                outIndex++;
                continue;
            }

            string pageMarkdown = page.ToMarkdown();
            var pageChunks = ChunkPdfText(path, fileName, pageNumber, pageText, pageMarkdown, opt, outIndex, ct, out var nextIndex);
            outIndex = nextIndex;
            foreach (var chunk in pageChunks) {
                yield return chunk;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadPdf(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        using var ms = CopyToMemory(stream, ct);
        var fileName = string.IsNullOrWhiteSpace(sourceName) ? "memory.pdf" : Path.GetFileName(sourceName!.Trim());
        var doc = PdfLogicalDocument.Load(ms.ToArray());
        int outIndex = 0;

        for (int pageIndex = 0; pageIndex < doc.Pages.Count; pageIndex++) {
            ct.ThrowIfCancellationRequested();

            var page = doc.Pages[pageIndex];
            var pageNumber = page.PageNumber;
            var pageText = BuildPdfPageText(page);
            if (string.IsNullOrWhiteSpace(pageText)) {
                yield return BuildPdfEmptyChunk(sourceName ?? fileName, fileName, pageNumber, outIndex);
                outIndex++;
                continue;
            }

            string pageMarkdown = page.ToMarkdown();
            var pageChunks = ChunkPdfText(sourceName ?? fileName, fileName, pageNumber, pageText, pageMarkdown, opt, outIndex, ct, out var nextIndex);
            outIndex = nextIndex;
            foreach (var chunk in pageChunks) {
                yield return chunk;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadMarkdown(string path, ReaderOptions opt, CancellationToken ct) {
        if (!opt.MarkdownChunkByHeadings) {
            foreach (var c in ChunkPlainTextByParagraphs(path, opt, ReaderInputKind.Markdown, ct, treatAsMarkdown: true))
                yield return c;
            yield break;
        }

        var fileName = Path.GetFileName(path);
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var text = ReadAllText(stream, ct, hardCapChars: null);
        foreach (var chunk in ChunkMarkdownFromText(text, path, fileName, opt, ct)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadMarkdown(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        var fileName = string.IsNullOrWhiteSpace(sourceName) ? "memory.md" : Path.GetFileName(sourceName!.Trim());
        var text = ReadAllText(stream, ct, hardCapChars: null);
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

    private static ReaderTable MapTable(ExcelExtractTable t, ExcelExtractLocation location, int tableIndex) {
        return new ReaderTable {
            Title = t.Title,
            Location = new ReaderLocation {
                Path = location.Path,
                Sheet = location.Sheet,
                A1Range = location.A1Range,
                SourceBlockIndex = location.BlockIndex,
                SourceBlockKind = "table",
                TableIndex = tableIndex
            },
            Columns = t.Columns,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(t.Columns, t.Rows),
            Rows = t.Rows,
            TotalRowCount = t.TotalRowCount,
            Truncated = t.Truncated
        };
    }

    private static IReadOnlyList<ReaderTable> MapTables(IReadOnlyList<ExcelExtractTable> tables, ExcelExtractLocation location, ref int nextTableIndex) {
        var mapped = new ReaderTable[tables.Count];
        for (int i = 0; i < tables.Count; i++) {
            mapped[i] = MapTable(tables[i], location, nextTableIndex);
            nextTableIndex++;
        }

        return mapped;
    }

    private static ReaderTable MapTable(TableBlock t, ReaderOptions opt) {
        var totalRowCount = t.Rows.Count;
        var rows = t.Rows;
        bool truncatedByOptions = false;

        if (opt.MaxTableRows > 0 && rows.Count > opt.MaxTableRows) {
            rows = rows.Take(opt.MaxTableRows).ToList();
            truncatedByOptions = true;
        }

        int columnCount = Math.Max(t.Headers.Count, rows.Count == 0 ? 0 : rows.Max(static row => row?.Count ?? 0));
        var columns = t.Headers.Count > 0
            ? EnsureMarkdownTableColumns(t.Headers, columnCount)
            : BuildMarkdownTableFallbackColumns(columnCount);

        var normalizedRows = rows
            .Select(row => NormalizeMarkdownTableRow(row, columnCount))
            .ToArray();

        return new ReaderTable {
            Columns = columns,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, normalizedRows),
            Rows = normalizedRows,
            TotalRowCount = totalRowCount,
            Truncated = truncatedByOptions || t.SkippedRowCount > 0 || t.SkippedColumnCount > 0
        };
    }

}
