using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO;
using OfficeIMO.Opml;
using OfficeIMO.Reader;

namespace OfficeIMO.Reader.Opml;

internal static partial class OpmlReaderAdapter {
    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions? readerOptions = null, ReaderOpmlOptions? opmlOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, reader.MaxInputBytes);
        ReaderOpmlOptions adapter = ReaderOpmlOptionsCloner.Clone(opmlOptions);
        ApplyReaderLimit(adapter.ReadOptions, reader.MaxInputBytes);
        cancellationToken.ThrowIfCancellationRequested();
        return BuildDocumentResult(OpmlDocument.Load(path, adapter.ReadOptions, cancellationToken), path, reader, adapter, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderOpmlOptions? opmlOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderOpmlOptions adapter = ReaderOpmlOptionsCloner.Clone(opmlOptions);
        ApplyReaderLimit(adapter.ReadOptions, reader.MaxInputBytes);
        string name = string.IsNullOrWhiteSpace(sourceName) ? "document.opml" : sourceName!;
        return BuildDocumentResult(OpmlDocument.Load(stream, adapter.ReadOptions, cancellationToken), name, reader, adapter, cancellationToken);
    }

    private static OfficeDocumentReadResult BuildDocumentResult(OpmlDocument document, string sourceName, ReaderOptions reader, ReaderOpmlOptions options, CancellationToken cancellationToken) {
        OpmlProjection projection = CreateProjection(document, sourceName, reader, options, includeChunkWarnings: false, cancellationToken);
        OfficeDocumentModel model = projection.Model;
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            projection.Chunks,
            ReaderInputKind.Opml,
            new OfficeDocumentSource { Path = sourceName, Title = model.Source.Title, Author = model.Source.Author },
            new[] { "officeimo.reader.opml.rich-v5", "officeimo.opml.lossless" });
        result.Metadata = model.Metadata.Select(MapMetadata).ToArray();
        result.Links = model.Links.Select(link => MapLink(link, sourceName)).ToArray();
        result.Diagnostics = projection.Diagnostics.Select(diagnostic => MapDiagnostic(diagnostic, sourceName)).ToArray();
        return result;
    }

    private static OfficeDocumentMetadataEntry MapMetadata(OfficeDocumentModelMetadataEntry metadata) => new OfficeDocumentMetadataEntry {
        Id = metadata.Id,
        Category = metadata.Category,
        Name = metadata.Name,
        Value = metadata.Value,
        ValueType = metadata.ValueType,
        SourceObjectId = metadata.SourceObjectId,
        Location = metadata.Location == null ? null : MapLocation(metadata.Location, metadata.Location.Path),
        Attributes = metadata.Attributes
    };

    private static OfficeDocumentLink MapLink(OfficeDocumentModelLink link, string sourceName) => new OfficeDocumentLink {
        Id = link.Id,
        Kind = link.Kind,
        Uri = link.Uri,
        DestinationName = link.DestinationName,
        DestinationPageNumber = link.DestinationPageNumber,
        DestinationMode = link.DestinationMode,
        NamedAction = link.NamedAction,
        RemoteFile = link.RemoteFile,
        RemoteDestinationName = link.RemoteDestinationName,
        RemoteDestinationPageNumber = link.RemoteDestinationPageNumber,
        Text = link.Text,
        Location = MapLocation(link.Location, sourceName)
    };

    private static OfficeDocumentDiagnostic MapDiagnostic(OpmlDiagnostic diagnostic, string sourceName) => new OfficeDocumentDiagnostic {
        Severity = diagnostic.Severity == OpmlDiagnosticSeverity.Error ? OfficeDocumentDiagnosticSeverity.Error
            : diagnostic.Severity == OpmlDiagnosticSeverity.Warning ? OfficeDocumentDiagnosticSeverity.Warning
            : OfficeDocumentDiagnosticSeverity.Information,
        Category = MapDiagnosticCategory(diagnostic.Code),
        Code = diagnostic.Code,
        Message = diagnostic.Message,
        Source = OfficeDocumentReaderBuilderOpmlExtensions.HandlerId,
        IsRecoverable = diagnostic.Severity != OpmlDiagnosticSeverity.Error,
        Location = new ReaderLocation { Path = sourceName, HeadingPath = diagnostic.Path }
    };

    private static OfficeDocumentDiagnosticCategory MapDiagnosticCategory(string code) {
        if (code.StartsWith("OPML", StringComparison.Ordinal) &&
            int.TryParse(code.Substring(4), out int number)) {
            if (number >= 1 && number <= 9) return OfficeDocumentDiagnosticCategory.Parsing;
            if (number >= 100) return OfficeDocumentDiagnosticCategory.Adapter;
        }
        return OfficeDocumentDiagnosticCategory.Content;
    }

    private static ReaderLocation MapLocation(OfficeDocumentModelLocation location, string? sourceName) => new ReaderLocation {
        Path = location.Path ?? sourceName,
        BlockIndex = location.BlockIndex,
        SourceBlockIndex = location.SourceBlockIndex,
        StartLine = location.StartLine,
        EndLine = location.EndLine,
        NormalizedStartLine = location.NormalizedStartLine,
        NormalizedEndLine = location.NormalizedEndLine,
        HeadingPath = location.HeadingPath,
        HeadingSlug = location.HeadingSlug,
        SourceBlockKind = location.SourceBlockKind,
        BlockAnchor = location.BlockAnchor,
        Sheet = location.Sheet,
        A1Range = location.A1Range,
        Slide = location.Slide,
        Page = location.Page,
        TableIndex = location.TableIndex
    };
}
