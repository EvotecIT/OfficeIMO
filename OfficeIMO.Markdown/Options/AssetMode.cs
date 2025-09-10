namespace OfficeIMO.Markdown;

/// <summary>
/// Connectivity mode used to decide how to include external assets (CSS/JS) referenced via URLs.
/// </summary>
public enum AssetMode {
    /// <summary>Reference assets by URL (CDN or provided links).</summary>
    Online,
    /// <summary>Inline assets so the output works offline. If URLs are provided, the renderer will attempt to download their content and inline it.</summary>
    Offline
}

