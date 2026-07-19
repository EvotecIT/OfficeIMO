using System;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private sealed class FolderIngestState {
        public int FilesScanned { get; set; }
        public int FilesParsed { get; set; }
        public int FilesSkipped { get; set; }
        public long BytesRead { get; set; }
        public int ChunksProduced { get; set; }
    }

    private sealed class SourceInfo {
        public string Path { get; set; } = string.Empty;
        public string SourceId { get; set; } = string.Empty;
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }
}
