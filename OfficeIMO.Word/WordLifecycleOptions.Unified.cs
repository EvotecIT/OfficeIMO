using OfficeIMO.Word.LegacyDoc;

namespace OfficeIMO.Word;

public sealed partial class WordCreateOptions {
    /// <summary>Theme for a new document. Null selects Modern for blank documents and preserves an existing template.</summary>
    public OfficeTheme? Theme { get; set; }

    /// <summary>Optional Word template copied into the new document.</summary>
    public string? Template { get; set; }
}

public sealed partial class WordLoadOptions {
    /// <summary>Rejects reported loss when a native Word source is projected into the editable model.</summary>
    public bool FailOnLoss { get; set; }
    /// <summary>Password for an encrypted Office source.</summary>
    public string? Password { get; set; }
    /// <summary>Maximum binary DOC stream size accepted by the native decoder.</summary>
    public int MaxLegacyInputBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Maximum aggregate decoded legacy image bytes.</summary>
    public int MaxDecodedImageBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Maximum legacy text characters decoded from the piece table.</summary>
    public int MaxDecodedCharacters { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum legacy paragraph-property traversal work, in bytes.</summary>
    public int MaxParagraphPropertyWorkBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Maximum legacy Data records followed by one paragraph-property chain.</summary>
    public int MaxDataPropertyChainLength { get; set; } = 256;
    /// <summary>Reports recognized unsupported content during legacy projection.</summary>
    public bool ReportUnsupportedContent { get; set; } = true;

    internal LegacyDocImportOptions CreateLegacyDocOptions() => new LegacyDocImportOptions {
        MaxInputBytes = (int)Math.Min(MaxInputBytes, MaxLegacyInputBytes),
        MaxDecodedImageBytes = MaxDecodedImageBytes,
        MaxDecodedCharacters = MaxDecodedCharacters,
        MaxParagraphPropertyWorkBytes = MaxParagraphPropertyWorkBytes,
        MaxDataPropertyChainLength = MaxDataPropertyChainLength,
        ReportUnsupportedContent = ReportUnsupportedContent
    };
}
