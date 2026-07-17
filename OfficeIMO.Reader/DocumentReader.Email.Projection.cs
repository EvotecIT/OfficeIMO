using OfficeIMO.Email;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>Projects one parsed email item while preserving indexes across an item-at-a-time store enumeration.</summary>
    internal static IReadOnlyList<ReaderChunk> ProjectEmailDocumentToChunks(
        EmailDocument document,
        string logicalPath,
        IReadOnlyList<EmailDiagnostic> diagnostics,
        string sourceName,
        ReaderOptions options,
        EmailDocumentProjectionCursor cursor,
        CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (logicalPath == null) throw new ArgumentNullException(nameof(logicalPath));
        if (diagnostics == null) throw new ArgumentNullException(nameof(diagnostics));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (cursor == null) throw new ArgumentNullException(nameof(cursor));

        var extraction = new EmailExtraction(document.Format, sourceName);
        extraction.Documents.Add(document);
        extraction.MailboxEntries.Add(null);
        extraction.LogicalPaths.Add(logicalPath);
        extraction.Diagnostics.AddRange(diagnostics);
        var context = new EmailChunkContext(
            extraction, options, cancellationToken,
            cursor.NextBlockIndex, cursor.DiagnosticsAttached);
        BuildEmailDocumentChunks(document, context, cursor.NextMessageIndex, depth: 0,
            parentHeading: null, logicalPath, mailboxEntry: null);
        cursor.NextMessageIndex++;
        cursor.NextBlockIndex = context.NextBlockIndex;
        cursor.DiagnosticsAttached = context.DiagnosticsAttached;
        return extraction.Chunks;
    }

    /// <summary>
    /// Projects already parsed email documents through Reader's shared email chunking pipeline.
    /// Parser adapters remain responsible for source-format parsing and logical item paths.
    /// </summary>
    internal static IReadOnlyList<ReaderChunk> ProjectEmailDocumentsToChunks(
        IReadOnlyList<EmailDocument> documents,
        IReadOnlyList<string?> logicalPaths,
        IReadOnlyList<EmailDiagnostic> diagnostics,
        EmailFileFormat format,
        string sourceName,
        ReaderOptions options,
        CancellationToken cancellationToken) {
        EmailExtraction extraction = CreateProjectedEmailExtraction(
            documents, logicalPaths, diagnostics, format, sourceName, options, cancellationToken);
        return extraction.Chunks;
    }

    /// <summary>Projects already parsed email documents into a rich Reader result for a file source.</summary>
    internal static OfficeDocumentReadResult ProjectEmailDocumentsToPathResult(
        IReadOnlyList<EmailDocument> documents,
        IReadOnlyList<string?> logicalPaths,
        IReadOnlyList<EmailDiagnostic> diagnostics,
        EmailFileFormat format,
        string sourceName,
        string path,
        ReaderOptions options,
        CancellationToken cancellationToken,
        bool? computeSourceHash = null) {
        EmailExtraction extraction = CreateProjectedEmailExtraction(
            documents, logicalPaths, diagnostics, format, sourceName, options, cancellationToken);
        SourceInfo source = BuildSourceInfoFromPath(path,
            computeSourceHash ?? options.ComputeHashes, cancellationToken);
        EnrichEmailChunks(extraction.Chunks, source, options.ComputeHashes);
        return BuildEmailDocumentResult(extraction, sourceName, BuildPathDocumentSource(path, extraction.Chunks));
    }

    /// <summary>Projects already parsed email documents into a rich Reader result for a stream source.</summary>
    internal static OfficeDocumentReadResult ProjectEmailDocumentsToStreamResult(
        IReadOnlyList<EmailDocument> documents,
        IReadOnlyList<string?> logicalPaths,
        IReadOnlyList<EmailDiagnostic> diagnostics,
        EmailFileFormat format,
        string sourceName,
        Stream stream,
        ReaderOptions options,
        CancellationToken cancellationToken,
        bool? computeSourceHash = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        EmailExtraction extraction = CreateProjectedEmailExtraction(
            documents, logicalPaths, diagnostics, format, sourceName, options, cancellationToken);
        SourceInfo source = BuildSourceInfoFromStream(stream, sourceName,
            computeSourceHash ?? options.ComputeHashes, cancellationToken);
        EnrichEmailChunks(extraction.Chunks, source, options.ComputeHashes);
        return BuildEmailDocumentResult(extraction, sourceName,
            BuildStreamDocumentSource(stream, sourceName, extraction.Chunks));
    }

    private static EmailExtraction CreateProjectedEmailExtraction(
        IReadOnlyList<EmailDocument> documents,
        IReadOnlyList<string?> logicalPaths,
        IReadOnlyList<EmailDiagnostic> diagnostics,
        EmailFileFormat format,
        string sourceName,
        ReaderOptions options,
        CancellationToken cancellationToken) {
        if (documents == null) throw new ArgumentNullException(nameof(documents));
        if (logicalPaths == null) throw new ArgumentNullException(nameof(logicalPaths));
        if (diagnostics == null) throw new ArgumentNullException(nameof(diagnostics));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (logicalPaths.Count != documents.Count) {
            throw new ArgumentException("A logical path is required for every projected email document.", nameof(logicalPaths));
        }

        var extraction = new EmailExtraction(format, sourceName);
        for (int index = 0; index < documents.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            extraction.Documents.Add(documents[index] ??
                throw new ArgumentException("Projected email documents cannot contain null entries.", nameof(documents)));
            extraction.MailboxEntries.Add(null);
            extraction.LogicalPaths.Add(logicalPaths[index]);
        }
        extraction.Diagnostics.AddRange(diagnostics);
        BuildEmailChunks(extraction, options, cancellationToken);
        return extraction;
    }
}

internal sealed class EmailDocumentProjectionCursor {
    internal int NextMessageIndex { get; set; }
    internal int NextBlockIndex { get; set; }
    internal bool DiagnosticsAttached { get; set; }
}
