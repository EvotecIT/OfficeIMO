using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
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

        string? logicalSourceName = null;
        var opt = NormalizeOptions(options);
        if (sourceName != null) {
            var trimmedSourceName = sourceName.Trim();
            if (trimmedSourceName.Length > 0) {
                logicalSourceName = trimmedSourceName;
            }
        }

        Stream readStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            ResolveStreamMaxInputBytes(logicalSourceName, opt,
                stream.CanSeek),
            cancellationToken,
            out bool ownsReadStream);
        try {
            var source = BuildSourceInfoFromStream(readStream,
                logicalSourceName, opt.ComputeHashes, cancellationToken);

            IEnumerable<ReaderChunk> raw;
            bool hasCustomStreamHandler = TryResolveStreamHandler(
                readStream,
                logicalSourceName,
                opt,
                cancellationToken,
                out ReaderHandlerDescriptor customStreamHandler,
                out ReaderDetectionResult detection);
            if (hasCustomStreamHandler) {
                if (customStreamHandler.ReadStream != null || customStreamHandler.ReadDocumentStream != null) {
                    raw = customStreamHandler.ReadStream != null
                        ? customStreamHandler.ReadStream(readStream, logicalSourceName, opt, cancellationToken)
                        : GetDocumentResultChunks(customStreamHandler.ReadDocumentStream!(readStream, logicalSourceName, opt, cancellationToken), customStreamHandler.Id);
                } else if (customStreamHandler.ReadDocumentStreamAsync != null) {
                    throw CreateAsyncOnlyHandlerException(customStreamHandler.Id, "stream");
                } else {
                    raw = ReadBuiltInStream(readStream, logicalSourceName, opt, cancellationToken, detection.Kind);
                }
            } else {
                raw = ReadBuiltInStream(readStream, logicalSourceName, opt, cancellationToken, detection.Kind);
            }

            foreach (var chunk in raw) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return EnrichChunk(chunk, source, opt.ComputeHashes);
            }
        } finally {
            if (ownsReadStream) {
                readStream.Dispose();
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadBuiltInStream(
        Stream stream,
        string? sourceName,
        ReaderOptions opt,
        CancellationToken cancellationToken,
        ReaderInputKind detectedKind) {
        ReaderInputKind kind = NormalizeBuiltInDispatchKind(detectedKind);
        if (kind == ReaderInputKind.Unknown && IsEmailArtifact(stream, opt, cancellationToken)) {
            kind = ReaderInputKind.Email;
        }
        return kind switch {
            ReaderInputKind.Word => ReadWord(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.Excel => ReadExcel(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.PowerPoint => ReadPowerPoint(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.Markdown => ReadMarkdown(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.Pdf => ReadPdf(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.Email => ReadEmail(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.Calendar => ReadCalendar(stream, sourceName, opt, cancellationToken),
            ReaderInputKind.VCard => ReadVCard(stream, sourceName, opt, cancellationToken),
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
        using var doc = LoadWordForReader(path, opt);
        IReadOnlyList<string>? legacyWarnings = BuildLegacyWordWarnings(doc);
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
                Warnings = CombineWarnings(c.Warnings, legacyWarnings)
            };
            outIndex++;
        }
    }

    private static IEnumerable<ReaderChunk> ReadWord(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // Copy input so we can open read-only without affecting caller's stream.
        using var ms = CopyToMemory(stream, ct);
        using var doc = LoadWordForReader(ms, opt);
        IReadOnlyList<string>? legacyWarnings = BuildLegacyWordWarnings(doc);

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
                Warnings = CombineWarnings(c.Warnings, legacyWarnings)
            };
            outIndex++;
        }
    }

    private static WordDocument LoadWordForReader(string path,
        ReaderOptions opt) {
        WordLoadOptions loadOptions = CreateWordLoadOptions(opt);
        try {
            return WordDocument.Load(path, loadOptions);
        } catch (Exception exception) when (ShouldRetryEncryptedWordOpen(
                     exception, opt)) {
            return WordDocument.LoadEncrypted(path, opt.OpenPassword!,
                loadOptions);
        }
    }

    private static WordDocument LoadWordForReader(Stream stream,
        ReaderOptions opt) {
        WordLoadOptions loadOptions = CreateWordLoadOptions(opt);
        stream.Position = 0;
        try {
            return WordDocument.Load(stream, loadOptions);
        } catch (Exception exception) when (ShouldRetryEncryptedWordOpen(
                     exception, opt)) {
            stream.Position = 0;
            return WordDocument.LoadEncrypted(stream, opt.OpenPassword!,
                loadOptions);
        }
    }

    private static WordLoadOptions CreateWordLoadOptions(
        ReaderOptions opt) => new WordLoadOptions {
        AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly,
        OpenSettings = CreateOpenSettings(opt)
    };

    private static bool ShouldRetryEncryptedWordOpen(Exception exception,
        ReaderOptions opt) =>
        !string.IsNullOrEmpty(opt.OpenPassword)
        && (exception is InvalidDataException
            || exception is OpenXmlPackageException
            || exception is IOException);

    private static IEnumerable<ReaderChunk> ReadExcel(string path, ReaderOptions opt, CancellationToken ct) {
        if (IsLegacyExcelExtension(path)) {
            using var legacyDocument = LoadLegacyExcelForReader(path, opt);
            using var legacyReader = legacyDocument.CreateReader();
            IReadOnlyList<string>? legacyWarnings = BuildLegacyExcelWarnings(legacyDocument);
            foreach (var chunk in ReadExcelChunks(legacyReader, path, opt, ct, legacyWarnings)) {
                yield return chunk;
            }
            yield break;
        }

        if (!string.IsNullOrEmpty(opt.OpenPassword)) {
            using var encryptedDocument = LoadOpenXmlExcelForReader(path, opt);
            using var encryptedReader = encryptedDocument.CreateReader();
            foreach (var chunk in ReadExcelChunks(encryptedReader, path, opt, ct, legacyWarnings: null)) {
                yield return chunk;
            }
            yield break;
        }

        // Use OpenSettings for basic OpenXML hardening (best-effort) and open from stream to avoid file handle collisions.
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        OpenSettings? openSettings = CreateOpenSettings(opt);
        using var openXml = openSettings == null
            ? SpreadsheetDocument.Open(fs, false)
            : SpreadsheetDocument.Open(fs, false, openSettings);
        using var reader = ExcelDocumentReader.Wrap(openXml);
        foreach (var chunk in ReadExcelChunks(reader, path, opt, ct, legacyWarnings: null)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadExcel(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // Avoid exposing OpenXml types in the public API surface; internally we can wrap.
        using var ms = CopyToMemory(stream, ct);
        OpenSettings? openSettings = CreateOpenSettings(opt);

        if (IsLegacyExcelExtension(sourceName)) {
            using var document = LoadLegacyExcelForReader(ms, opt);
            using var reader = document.CreateReader();
            IReadOnlyList<string>? legacyWarnings = BuildLegacyExcelWarnings(document);
            foreach (var chunk in ReadExcelChunks(reader, sourceName, opt, ct, legacyWarnings)) {
                yield return chunk;
            }
            yield break;
        }

        if (!string.IsNullOrEmpty(opt.OpenPassword)) {
            using var document = LoadOpenXmlExcelForReader(ms, opt);
            using var reader = document.CreateReader();
            foreach (var chunk in ReadExcelChunks(reader, sourceName, opt, ct, legacyWarnings: null)) {
                yield return chunk;
            }
            yield break;
        }

        using var openXml = openSettings == null
            ? SpreadsheetDocument.Open(ms, false)
            : SpreadsheetDocument.Open(ms, false, openSettings);
        using var wrappedReader = ExcelDocumentReader.Wrap(openXml);
        foreach (var chunk in ReadExcelChunks(wrappedReader, sourceName, opt, ct, legacyWarnings: null)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadExcelChunks(
        ExcelDocumentReader reader,
        string? sourceName,
        ReaderOptions opt,
        CancellationToken ct,
        IReadOnlyList<string>? legacyWarnings) {
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
                        Path = c.Location.Path,
                        Sheet = c.Location.Sheet,
                        A1Range = c.Location.A1Range,
                        BlockIndex = outIndex,
                        SourceBlockIndex = c.Location.BlockIndex
                    },
                    Text = c.Text,
                    Markdown = c.Markdown,
                    Tables = tables,
                    Warnings = CombineWarnings(c.Warnings, legacyWarnings)
                };
                outIndex++;
            }
        }
    }

    private static ExcelDocument LoadLegacyExcelForReader(string path, ReaderOptions opt) {
        OpenSettings? openSettings = CreateOpenSettings(opt);
        if (!string.IsNullOrEmpty(opt.OpenPassword)) {
            return ExcelDocument.LoadLegacyXls(path, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                Password = opt.OpenPassword
            });
        }

        return ExcelDocument.Load(path, new ExcelLoadOptions {
            AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly,
            OpenSettings = openSettings
        });
    }

    private static ExcelDocument LoadLegacyExcelForReader(Stream stream, ReaderOptions opt) {
        OpenSettings? openSettings = CreateOpenSettings(opt);
        if (!string.IsNullOrEmpty(opt.OpenPassword)) {
            stream.Position = 0;
            return ExcelDocument.LoadLegacyXls(stream, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                Password = opt.OpenPassword
            });
        }

        stream.Position = 0;
        return ExcelDocument.Load(stream, new ExcelLoadOptions {
            AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly,
            OpenSettings = openSettings
        });
    }

    private static ExcelDocument LoadOpenXmlExcelForReader(string path, ReaderOptions opt) {
        OpenSettings? openSettings = CreateOpenSettings(opt);
        try {
            return ExcelDocument.Load(path, new ExcelLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly,
                OpenSettings = openSettings
            });
        } catch (Exception ex) when (ShouldRetryEncryptedExcelOpen(ex, opt)) {
            try {
                return ExcelDocument.LoadEncrypted(path, opt.OpenPassword!, new ExcelLoadOptions {
                    AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly,
                    OpenSettings = openSettings
                });
            } catch {
                ExceptionDispatchInfo.Capture(ex).Throw();
                throw;
            }
        }
    }

    private static ExcelDocument LoadOpenXmlExcelForReader(Stream stream, ReaderOptions opt) {
        OpenSettings? openSettings = CreateOpenSettings(opt);
        stream.Position = 0;
        try {
            return ExcelDocument.Load(stream, new ExcelLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly,
                OpenSettings = openSettings
            });
        } catch (Exception ex) when (ShouldRetryEncryptedExcelOpen(ex, opt)) {
            stream.Position = 0;
            try {
                return ExcelDocument.LoadEncrypted(stream, opt.OpenPassword!, new ExcelLoadOptions {
                    AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly,
                    OpenSettings = openSettings
                });
            } catch {
                ExceptionDispatchInfo.Capture(ex).Throw();
                throw;
            }
        }
    }

    private static bool ShouldRetryEncryptedExcelOpen(Exception exception, ReaderOptions opt) {
        return !string.IsNullOrEmpty(opt.OpenPassword)
            && (exception is InvalidDataException
                || exception is OpenXmlPackageException
                || exception is IOException);
    }

    private static IReadOnlyList<string>? BuildLegacyWordWarnings(WordDocument document) {
        if (document.SourceFormat != WordFileFormat.Doc) {
            return null;
        }

        var warnings = new List<string>();
        AddBoundedWarnings(
            warnings,
            document.LegacyDocImportDiagnostics.Select(static diagnostic => "Legacy DOC import diagnostic: " + diagnostic),
            maxItems: 8,
            overflowMessage: "Additional legacy DOC import diagnostics were omitted.");
        AddBoundedWarnings(
            warnings,
            document.LegacyDocUnsupportedFeatures.Select(static feature => $"Legacy DOC unsupported feature: {feature.Code} ({feature.Kind}) - {feature.Description}"),
            maxItems: 8,
            overflowMessage: "Additional legacy DOC unsupported features were omitted.");
        AddBoundedWarnings(
            warnings,
            document.LegacyDocPreservedFeatures.Select(static feature => $"Legacy DOC preserved feature: {feature.Code} ({feature.Kind}) - {feature.Description}"),
            maxItems: 8,
            overflowMessage: "Additional legacy DOC preserved features were omitted.");
        AddBoundedWarnings(
            warnings,
            document.LegacyDocCompoundFeatures.Select(static feature => $"Legacy DOC compound feature: {feature.Code} ({feature.Kind}) - {feature.Description}"),
            maxItems: 8,
            overflowMessage: "Additional legacy DOC compound features were omitted.");
        return warnings.Count == 0 ? null : warnings;
    }

    internal static IReadOnlyList<string>? BuildLegacyExcelWarnings(ExcelDocument document) {
        if (document.SourceFormat != ExcelFileFormat.Xls) {
            return null;
        }

        var warnings = new List<string>();
        AddBoundedWarnings(
            warnings,
            document.LegacyXlsImportDiagnostics.Select(static diagnostic => "Legacy XLS import diagnostic: " + diagnostic),
            maxItems: 8,
            overflowMessage: "Additional legacy XLS import diagnostics were omitted.");
        AddBoundedWarnings(
            warnings,
            document.LegacyXlsUnsupportedFeatures.Select(static feature => $"Legacy XLS unsupported feature: {feature.Code} ({feature.Kind}) - {feature.Description}"),
            maxItems: 8,
            overflowMessage: "Additional legacy XLS unsupported features were omitted.");
        AddBoundedWarnings(
            warnings,
            document.LegacyXlsPreservedFeatures.Select(static feature => $"Legacy XLS preserved feature: {feature.Code} ({feature.Kind}) - {feature.Description}"),
            maxItems: 8,
            overflowMessage: "Additional legacy XLS preserved features were omitted.");
        AddBoundedWarnings(
            warnings,
            document.LegacyXlsUnsupportedSheets.Select(static sheet => $"Legacy XLS unsupported sheet: {sheet.Name} ({sheet.Kind}, {sheet.VisibilityName})"),
            maxItems: 8,
            overflowMessage: "Additional legacy XLS unsupported sheets were omitted.");
        AddBoundedWarnings(
            warnings,
            document.LegacyXlsChartSheets.Select(static sheet => $"Legacy XLS chart sheet: {sheet.Name} ({sheet.VisibilityName})"),
            maxItems: 8,
            overflowMessage: "Additional legacy XLS chart sheets were omitted.");
        AddBoundedWarnings(
            warnings,
            document.LegacyXlsCompoundFeatures.Select(static feature => $"Legacy XLS compound feature: {feature.Kind} - entries: {feature.Entries.Count}"),
            maxItems: 8,
            overflowMessage: "Additional legacy XLS compound features were omitted.");
        return warnings.Count == 0 ? null : warnings;
    }

    private static IReadOnlyList<string>? CombineWarnings(IReadOnlyList<string>? primary, IReadOnlyList<string>? secondary) {
        if (secondary == null || secondary.Count == 0) {
            return primary;
        }

        var warnings = new List<string>();
        AddWarnings(warnings, primary);
        AddWarnings(warnings, secondary);
        return warnings.Count == 0 ? null : warnings;
    }

    private static void AddWarnings(List<string> target, IEnumerable<string>? warnings) {
        if (warnings == null) {
            return;
        }

        foreach (string warning in warnings) {
            if (!string.IsNullOrWhiteSpace(warning)) {
                target.Add(warning);
            }
        }
    }

    private static void AddBoundedWarnings(List<string> target, IEnumerable<string> warnings, int maxItems, string overflowMessage) {
        int count = 0;
        foreach (string warning in warnings) {
            if (count < maxItems && !string.IsNullOrWhiteSpace(warning)) {
                target.Add(warning);
            }

            count++;
        }

        if (count > maxItems) {
            target.Add($"{overflowMessage} ({count - maxItems} more)");
        }
    }

    private static IEnumerable<ReaderChunk> ReadPowerPoint(string path, ReaderOptions opt, CancellationToken ct) {
        using var presentation = LoadPowerPointForReader(path, opt, ct);
        IReadOnlyList<string>? legacyWarnings =
            BuildLegacyPowerPointWarnings(presentation);
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
                Warnings = CombineWarnings(c.Warnings, legacyWarnings)
            };
            outIndex++;
        }
    }

    private static IEnumerable<ReaderChunk> ReadPowerPoint(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // Read-only stream opening already copies to an internal stream for safety.
        using var presentation = LoadPowerPointForReader(stream, opt, ct);
        IReadOnlyList<string>? legacyWarnings =
            BuildLegacyPowerPointWarnings(presentation);
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
                Warnings = CombineWarnings(c.Warnings, legacyWarnings)
            };
            outIndex++;
        }
    }

    private static PowerPointPresentation LoadPowerPointForReader(
        string path, ReaderOptions options,
        CancellationToken cancellationToken = default) {
        PowerPointLoadOptions loadOptions =
            CreatePowerPointReaderLoadOptions(options);
        try {
            return PowerPointPresentation.Load(path, loadOptions,
                cancellationToken);
        } catch (Exception exception) when (
            ShouldRetryEncryptedPowerPointOpen(exception, options)) {
            try {
                return PowerPointPresentation.LoadEncrypted(path,
                    options.OpenPassword!, loadOptions,
                    cancellationToken);
            } catch (CryptographicException) {
                throw;
            } catch {
                ExceptionDispatchInfo.Capture(exception).Throw();
                throw;
            }
        }
    }

    private static PowerPointPresentation LoadPowerPointForReader(
        Stream stream, ReaderOptions options,
        CancellationToken cancellationToken = default) {
        if (stream.CanSeek) stream.Position = 0;
        PowerPointLoadOptions loadOptions =
            CreatePowerPointReaderLoadOptions(options);
        try {
            return PowerPointPresentation.Load(stream, loadOptions,
                cancellationToken);
        } catch (Exception exception) when (stream.CanSeek
            && ShouldRetryEncryptedPowerPointOpen(exception, options)) {
            stream.Position = 0;
            try {
                return PowerPointPresentation.LoadEncrypted(stream,
                    options.OpenPassword!, loadOptions,
                    cancellationToken);
            } catch (CryptographicException) {
                throw;
            } catch {
                ExceptionDispatchInfo.Capture(exception).Throw();
                throw;
            }
        }
    }

    private static bool ShouldRetryEncryptedPowerPointOpen(
        Exception exception, ReaderOptions options) =>
        !string.IsNullOrEmpty(options.OpenPassword)
        && (exception is InvalidDataException
            || exception is OpenXmlPackageException
            || exception is IOException);

    private static PowerPointLoadOptions CreatePowerPointReaderLoadOptions(
        ReaderOptions options) {
        long requestedMaxInputBytes = options.MaxInputBytes
            ?? LegacyPptImportOptions.DefaultMaxInputBytes;
        var loadOptions = new PowerPointLoadOptions {
            AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly,
            OpenSettings = CreateOpenSettings(options),
            LegacyPptImportOptions = new LegacyPptImportOptions {
                MaxInputBytes = requestedMaxInputBytes > int.MaxValue
                    ? int.MaxValue
                    : checked((int)requestedMaxInputBytes),
                Password = options.OpenPassword,
                ReportUnsupportedContent = true
            }
        };
        return loadOptions;
    }

    internal static IReadOnlyList<string>? BuildLegacyPowerPointWarnings(
        PowerPointPresentation presentation) {
        if (presentation.SourceFormat is not PowerPointFileFormat.Ppt
            and not PowerPointFileFormat.Pot
            and not PowerPointFileFormat.Pps) {
            return null;
        }

        var warnings = new List<string>();
        AddBoundedWarnings(warnings,
            presentation.LegacyPptImportDiagnostics.Select(
                static diagnostic =>
                    "Legacy PPT import diagnostic: " + diagnostic),
            maxItems: 16,
            overflowMessage:
                "Additional legacy PPT import diagnostics were omitted.");
        return warnings.Count == 0 ? null : warnings;
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
