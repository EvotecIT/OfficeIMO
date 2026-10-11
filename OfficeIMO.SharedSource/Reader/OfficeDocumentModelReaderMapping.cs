using OfficeIMO;

namespace OfficeIMO.Reader {

    /// <summary>Projects shared document values while callers retain format-specific fallback policy.</summary>
    internal static class OfficeDocumentModelReaderMapping {
        internal static OfficeDocumentMetadataEntry Metadata(OfficeDocumentModelMetadataEntry metadata, string? missingPath = null) =>
            new OfficeDocumentMetadataEntry {
                Id = metadata.Id,
                Category = metadata.Category,
                Name = metadata.Name,
                Value = metadata.Value,
                ValueType = metadata.ValueType,
                SourceObjectId = metadata.SourceObjectId,
                Location = metadata.Location == null ? null : Location(metadata.Location, missingPath),
                Attributes = metadata.Attributes
            };

        internal static OfficeDocumentLink Link(OfficeDocumentModelLink link, string sourceName, string? defaultSourceBlockKind = null) =>
            new OfficeDocumentLink {
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
                Location = Location(link.Location, sourceName, defaultSourceBlockKind)
            };

        internal static ReaderLocation Location(OfficeDocumentModelLocation? location, string? sourceName, string? defaultSourceBlockKind = null) =>
            new ReaderLocation {
                Path = location?.Path ?? sourceName,
                BlockIndex = location?.BlockIndex,
                SourceBlockIndex = location?.SourceBlockIndex,
                StartLine = location?.StartLine,
                EndLine = location?.EndLine,
                NormalizedStartLine = location?.NormalizedStartLine,
                NormalizedEndLine = location?.NormalizedEndLine,
                HeadingPath = location?.HeadingPath,
                HeadingSlug = location?.HeadingSlug,
                SourceBlockKind = location?.SourceBlockKind ?? defaultSourceBlockKind,
                BlockAnchor = location?.BlockAnchor,
                Sheet = location?.Sheet,
                A1Range = location?.A1Range,
                Slide = location?.Slide,
                Page = location?.Page,
                TableIndex = location?.TableIndex
            };
    }
}
