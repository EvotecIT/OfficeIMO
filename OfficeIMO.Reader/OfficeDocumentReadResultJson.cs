using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Reader;

/// <summary>
/// JSON serialization helpers for the shared OfficeIMO document read result envelope.
/// </summary>
public static class OfficeDocumentReadResultJson {
    /// <summary>
    /// Serializes a document read result into the stable OfficeIMO transport shape.
    /// </summary>
    /// <param name="result">Read result to serialize.</param>
    /// <param name="indented">When true, writes indented JSON for diagnostics and fixtures.</param>
    public static string Serialize(OfficeDocumentReadResult result, bool indented = false) {
        if (result == null) throw new ArgumentNullException(nameof(result));

        return JsonSerializer.Serialize(ProjectResult(result), CreateOptions(indented));
    }

    /// <summary>
    /// Serializes a document read result into the stable OfficeIMO transport shape.
    /// </summary>
    /// <param name="result">Read result to serialize.</param>
    /// <param name="indented">When true, writes indented JSON for diagnostics and fixtures.</param>
    public static string ToJson(this OfficeDocumentReadResult result, bool indented = false) {
        return Serialize(result, indented);
    }

    private static JsonSerializerOptions CreateOptions(bool indented) {
        return new JsonSerializerOptions {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            WriteIndented = indented
        };
    }

    private static object ProjectResult(OfficeDocumentReadResult result) {
        return new {
            schemaId = string.IsNullOrWhiteSpace(result.SchemaId) ? OfficeDocumentReadResultSchema.Id : result.SchemaId,
            schemaVersion = result.SchemaVersion == 0 ? OfficeDocumentReadResultSchema.Version : result.SchemaVersion,
            kind = result.Kind.ToString(),
            source = ProjectSource(result.Source ?? new OfficeDocumentSource()),
            capabilitiesUsed = OrEmpty(result.CapabilitiesUsed),
            markdown = result.Markdown,
            html = result.Html,
            json = result.Json,
            chunks = ProjectCollection(result.Chunks, ProjectChunk),
            metadata = ProjectCollection(result.Metadata, ProjectMetadataEntry),
            pages = ProjectCollection(result.Pages, ProjectPage),
            blocks = ProjectCollection(result.Blocks, ProjectBlock),
            tables = ProjectCollection(result.Tables, ProjectTable),
            assets = ProjectCollection(result.Assets, ProjectAsset),
            links = ProjectCollection(result.Links, ProjectLink),
            forms = ProjectCollection(result.Forms, ProjectForm),
            ocrCandidates = ProjectCollection(result.OcrCandidates, ProjectOcrCandidate),
            visuals = ProjectCollection(result.Visuals, ProjectVisual),
            diagnostics = ProjectCollection(result.Diagnostics, ProjectDiagnostic)
        };
    }

    private static object ProjectMetadataEntry(OfficeDocumentMetadataEntry entry) {
        return new {
            id = entry.Id,
            category = entry.Category,
            name = entry.Name,
            value = entry.Value,
            valueType = entry.ValueType,
            sourceObjectId = entry.SourceObjectId,
            location = ProjectLocation(entry.Location),
            attributes = ProjectAttributes(entry.Attributes)
        };
    }

    private static object ProjectSource(OfficeDocumentSource source) {
        return new {
            path = source.Path,
            sourceId = source.SourceId,
            sourceHash = source.SourceHash,
            lastWriteUtc = source.LastWriteUtc,
            lengthBytes = source.LengthBytes,
            title = source.Title,
            author = source.Author,
            subject = source.Subject,
            keywords = source.Keywords
        };
    }

    private static object ProjectChunk(ReaderChunk chunk) {
        return new {
            id = chunk.Id,
            kind = chunk.Kind.ToString(),
            location = ProjectLocation(chunk.Location),
            sourceId = chunk.SourceId,
            sourceHash = chunk.SourceHash,
            chunkHash = chunk.ChunkHash,
            sourceLastWriteUtc = chunk.SourceLastWriteUtc,
            sourceLengthBytes = chunk.SourceLengthBytes,
            tokenEstimate = chunk.TokenEstimate,
            text = chunk.Text,
            markdown = chunk.Markdown,
            tables = ProjectNullableCollection(chunk.Tables, ProjectTable),
            visuals = ProjectNullableCollection(chunk.Visuals, ProjectVisual),
            warnings = chunk.Warnings
        };
    }

    private static object ProjectPage(OfficeDocumentPage page) {
        return new {
            number = page.Number,
            name = page.Name,
            width = page.Width,
            height = page.Height,
            rotationDegrees = page.RotationDegrees,
            location = ProjectLocation(page.Location),
            blocks = ProjectCollection(page.Blocks, ProjectBlock),
            tables = ProjectCollection(page.Tables, ProjectTable),
            assets = ProjectCollection(page.Assets, ProjectAsset),
            links = ProjectCollection(page.Links, ProjectLink),
            forms = ProjectCollection(page.Forms, ProjectForm),
            ocrCandidates = ProjectCollection(page.OcrCandidates, ProjectOcrCandidate)
        };
    }

    private static object ProjectBlock(OfficeDocumentBlock block) {
        return new {
            id = block.Id,
            kind = block.Kind,
            text = block.Text,
            level = block.Level,
            marker = block.Marker,
            location = ProjectLocation(block.Location),
            region = ProjectRegion(block.Region)
        };
    }

    private static object ProjectTable(ReaderTable table) {
        return new {
            title = table.Title,
            kind = table.Kind,
            callId = table.CallId,
            summary = table.Summary,
            payloadHash = table.PayloadHash,
            location = ProjectLocation(table.Location),
            columns = OrEmpty(table.Columns),
            columnProfiles = ProjectCollection(table.ColumnProfiles, ProjectColumnProfile),
            rows = ProjectRows(table.Rows),
            totalRowCount = table.TotalRowCount,
            truncated = table.Truncated
        };
    }

    private static object ProjectColumnProfile(ReaderTableColumnProfile profile) {
        return new {
            index = profile.Index,
            name = profile.Name,
            kind = profile.Kind.ToString(),
            nonEmptyCellCount = profile.NonEmptyCellCount,
            numericCellCount = profile.NumericCellCount,
            confidence = profile.Confidence,
            isNumeric = profile.IsNumeric
        };
    }

    private static object ProjectAsset(OfficeDocumentAsset asset) {
        return new {
            id = asset.Id,
            kind = asset.Kind,
            mediaType = asset.MediaType,
            extension = asset.Extension,
            fileName = asset.FileName,
            altText = asset.AltText,
            title = asset.Title,
            width = asset.Width,
            height = asset.Height,
            lengthBytes = asset.LengthBytes,
            payloadHash = asset.PayloadHash,
            sourceObjectId = asset.SourceObjectId,
            location = ProjectLocation(asset.Location)
        };
    }

    private static object ProjectLink(OfficeDocumentLink link) {
        return new {
            id = link.Id,
            kind = link.Kind,
            uri = link.Uri,
            destinationName = link.DestinationName,
            destinationPageNumber = link.DestinationPageNumber,
            destinationMode = link.DestinationMode,
            destinationTop = link.DestinationTop,
            destinationLeft = link.DestinationLeft,
            destinationBottom = link.DestinationBottom,
            destinationRight = link.DestinationRight,
            namedAction = link.NamedAction,
            remoteFile = link.RemoteFile,
            remoteDestinationName = link.RemoteDestinationName,
            remoteDestinationPageNumber = link.RemoteDestinationPageNumber,
            remoteDestinationMode = link.RemoteDestinationMode,
            remoteDestinationTop = link.RemoteDestinationTop,
            remoteDestinationLeft = link.RemoteDestinationLeft,
            remoteDestinationBottom = link.RemoteDestinationBottom,
            remoteDestinationRight = link.RemoteDestinationRight,
            text = link.Text,
            location = ProjectLocation(link.Location),
            region = ProjectRegion(link.Region)
        };
    }

    private static object ProjectForm(OfficeDocumentFormField form) {
        return new {
            id = form.Id,
            name = form.Name,
            kind = form.Kind,
            value = form.Value,
            isReadOnly = form.IsReadOnly,
            isRequired = form.IsRequired,
            location = ProjectLocation(form.Location),
            region = ProjectRegion(form.Region)
        };
    }

    private static object ProjectVisual(ReaderVisual visual) {
        return new {
            kind = visual.Kind,
            language = visual.Language,
            content = visual.Content,
            payloadHash = visual.PayloadHash,
            location = ProjectLocation(visual.Location)
        };
    }

    private static object ProjectOcrCandidate(OfficeDocumentOcrCandidate candidate) {
        return new {
            id = candidate.Id,
            kind = candidate.Kind,
            reason = candidate.Reason,
            confidence = candidate.Confidence,
            assetId = candidate.AssetId,
            imageCount = candidate.ImageCount,
            textBlockCount = candidate.TextBlockCount,
            location = ProjectLocation(candidate.Location),
            region = ProjectRegion(candidate.Region)
        };
    }

    private static object ProjectDiagnostic(OfficeDocumentDiagnostic diagnostic) {
        return new {
            severity = diagnostic.Severity.ToString(),
            code = diagnostic.Code,
            message = diagnostic.Message,
            location = ProjectLocation(diagnostic.Location)
        };
    }

    private static object? ProjectLocation(ReaderLocation? location) {
        if (location == null) return null;

        return new {
            path = location.Path,
            blockIndex = location.BlockIndex,
            sourceBlockIndex = location.SourceBlockIndex,
            startLine = location.StartLine,
            endLine = location.EndLine,
            normalizedStartLine = location.NormalizedStartLine,
            normalizedEndLine = location.NormalizedEndLine,
            headingPath = location.HeadingPath,
            headingSlug = location.HeadingSlug,
            sourceBlockKind = location.SourceBlockKind,
            blockAnchor = location.BlockAnchor,
            sheet = location.Sheet,
            a1Range = location.A1Range,
            slide = location.Slide,
            page = location.Page,
            tableIndex = location.TableIndex
        };
    }

    private static object? ProjectRegion(OfficeDocumentRegion? region) {
        if (region == null) return null;

        return new {
            x = region.X,
            y = region.Y,
            width = region.Width,
            height = region.Height
        };
    }

    private static IReadOnlyList<IReadOnlyList<string>> ProjectRows(IReadOnlyList<IReadOnlyList<string>>? rows) {
        return rows ?? Array.Empty<IReadOnlyList<string>>();
    }

    private static IReadOnlyDictionary<string, string> ProjectAttributes(IReadOnlyDictionary<string, string>? attributes) {
        var sorted = new SortedDictionary<string, string>(StringComparer.Ordinal);
        if (attributes == null || attributes.Count == 0) {
            return sorted;
        }

        foreach (KeyValuePair<string, string> attribute in attributes) {
            sorted[attribute.Key] = attribute.Value;
        }

        return sorted;
    }

    private static IReadOnlyList<object> ProjectCollection<T>(IReadOnlyList<T>? values, Func<T, object> projector) {
        if (values == null || values.Count == 0) {
            return Array.Empty<object>();
        }

        var projected = new object[values.Count];
        for (int i = 0; i < values.Count; i++) {
            projected[i] = projector(values[i]);
        }

        return projected;
    }

    private static IReadOnlyList<object>? ProjectNullableCollection<T>(IReadOnlyList<T>? values, Func<T, object> projector) {
        return values == null ? null : ProjectCollection(values, projector);
    }

    private static IReadOnlyList<T> OrEmpty<T>(IReadOnlyList<T>? values) {
        return values ?? Array.Empty<T>();
    }
}
