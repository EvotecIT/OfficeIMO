namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    private sealed class EpubPackage {
        public string OpfPath { get; set; } = string.Empty;
        public string? PackageVersion { get; set; }
        public string? UniqueIdentifierId { get; set; }
        public string? Title { get; set; }
        public string? Identifier { get; set; }
        public string? Language { get; set; }
        public string? Creator { get; set; }
        public EpubRenditionLayout? RenditionLayout { get; set; }
        public List<EpubMetadataEntry> Metadata { get; } = new List<EpubMetadataEntry>();
        public Dictionary<string, ManifestItem> Manifest { get; } = new Dictionary<string, ManifestItem>(StringComparer.Ordinal);
        public List<SpineItem> Spine { get; } = new List<SpineItem>();
        public List<EpubNavigationItem> Guide { get; } = new List<EpubNavigationItem>();
        public string? NavDocumentPath { get; set; }
        public string? NcxPath { get; set; }
    }

    private sealed class ManifestItem {
        public string Id { get; set; } = string.Empty;
        public string Href { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string MediaType { get; set; } = string.Empty;
        public string Properties { get; set; } = string.Empty;
        public bool IsRemote { get; set; }
        public string? RemoteUri { get; set; }
    }

    private sealed class SpineItem {
        public string IdRef { get; set; } = string.Empty;
        public int SpineIndex { get; set; }
        public bool IsLinear { get; set; }
        public string? Properties { get; set; }
        public EpubRenditionLayout? RenditionLayout { get; set; }
    }

    private sealed class ChapterCandidate {
        public ZipArchiveEntry Entry { get; set; } = null!;
        public string Path { get; set; } = string.Empty;
        public string? ManifestId { get; set; }
        public string? MediaType { get; set; }
        public int? SpineIndex { get; set; }
        public bool? IsLinear { get; set; }
        public EpubRenditionLayout? RenditionLayout { get; set; }
    }

    private sealed class EpubNavigationResult {
        public Dictionary<string, string> TitleMap { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
        public List<EpubNavigationItem> TableOfContents { get; } = new List<EpubNavigationItem>();
        public List<EpubNavigationItem> PageList { get; } = new List<EpubNavigationItem>();
        public List<EpubNavigationItem> Landmarks { get; } = new List<EpubNavigationItem>();
    }

    private sealed class NavigationLimitState {
        public int Count { get; set; }
        public bool CountLimitReported { get; set; }
        public bool DepthLimitReported { get; set; }
    }
}
