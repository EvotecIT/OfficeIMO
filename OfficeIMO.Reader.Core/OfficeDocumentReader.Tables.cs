using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Extracts Markdown tables using the configured reader parsing contract.</summary>
    public IReadOnlyList<ReaderTable> ExtractMarkdownTables(
        string markdown,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractMarkdownTables(markdown, options, cancellationToken);
    }

    /// <summary>Extracts tables from already-read chunks.</summary>
    public IReadOnlyList<ReaderTable> ExtractTables(
        IEnumerable<ReaderChunk> chunks,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractTables(chunks, cancellationToken);
    }

    /// <summary>Exports tables as deterministic CSV, Markdown, and JSON payloads.</summary>
    public IReadOnlyList<ReaderTableExportBundle> ExportTables(
        IEnumerable<ReaderTable> tables,
        bool indentedJson = false,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExportTables(tables, indentedJson, cancellationToken);
    }

    /// <summary>Reads a file and returns discovered tables in source order.</summary>
    public IReadOnlyList<ReaderTable> ReadTables(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractTables(Read(path, options, cancellationToken), cancellationToken);
    }

    /// <summary>Reads a caller-owned stream and returns discovered tables in source order.</summary>
    public IReadOnlyList<ReaderTable> ReadTables(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractTables(Read(stream, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>Reads bytes and returns discovered tables in source order.</summary>
    public IReadOnlyList<ReaderTable> ReadTables(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        return DocumentReaderEngine.ExtractTables(Read(bytes, sourceName, options, cancellationToken), cancellationToken);
    }

    /// <summary>Reads a file and exports its tables as deterministic CSV, Markdown, and JSON payloads.</summary>
    public IReadOnlyList<ReaderTableExportBundle> ReadTableExports(
        string path,
        ReaderOptions? options = null,
        bool indentedJson = false,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExportTables(ReadTables(path, options, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>Reads a caller-owned stream and exports its tables as deterministic CSV, Markdown, and JSON payloads.</summary>
    public IReadOnlyList<ReaderTableExportBundle> ReadTableExports(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        bool indentedJson = false,
        CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExportTables(
            ReadTables(stream, sourceName, options, cancellationToken),
            indentedJson,
            cancellationToken);
    }

    /// <summary>Reads bytes and exports their tables as deterministic CSV, Markdown, and JSON payloads.</summary>
    public IReadOnlyList<ReaderTableExportBundle> ReadTableExports(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        bool indentedJson = false,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        return DocumentReaderEngine.ExportTables(
            ReadTables(bytes, sourceName, options, cancellationToken),
            indentedJson,
            cancellationToken);
    }
}
