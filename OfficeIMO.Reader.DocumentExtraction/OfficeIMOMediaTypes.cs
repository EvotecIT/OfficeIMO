namespace OfficeIMO.Reader.DocumentExtraction;

/// <summary>Maps declared media types to the file extension OfficeIMO uses to select a reader handler.</summary>
internal static class OfficeIMOMediaTypes {
    private static readonly Dictionary<string, string> Extensions = new(StringComparer.OrdinalIgnoreCase) {
        ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"] = ".docx",
        ["application/vnd.ms-word.document.macroenabled.12"] = ".docm",
        ["application/vnd.openxmlformats-officedocument.wordprocessingml.template"] = ".dotx",
        ["application/msword"] = ".doc",
        ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"] = ".xlsx",
        ["application/vnd.ms-excel.sheet.macroenabled.12"] = ".xlsm",
        ["application/vnd.ms-excel"] = ".xls",
        ["application/vnd.openxmlformats-officedocument.presentationml.presentation"] = ".pptx",
        ["application/vnd.ms-powerpoint.presentation.macroenabled.12"] = ".pptm",
        ["application/vnd.ms-powerpoint"] = ".ppt",
        ["application/vnd.ms-visio.drawing"] = ".vsdx",
        ["application/vnd.visio"] = ".vsd",
        ["application/vnd.oasis.opendocument.text"] = ".odt",
        ["application/vnd.oasis.opendocument.spreadsheet"] = ".ods",
        ["application/vnd.oasis.opendocument.presentation"] = ".odp",
        ["application/pdf"] = ".pdf",
        ["application/oxps"] = ".oxps",
        ["application/vnd.ms-xpsdocument"] = ".xps",
        ["application/epub+zip"] = ".epub",
        ["text/html"] = ".html",
        ["application/xhtml+xml"] = ".xhtml",
        ["multipart/related"] = ".mhtml",
        ["message/rfc822"] = ".eml",
        ["application/vnd.ms-outlook"] = ".msg",
        ["text/markdown"] = ".md",
        ["text/x-markdown"] = ".md",
        ["text/plain"] = ".txt",
        ["text/csv"] = ".csv",
        ["text/tab-separated-values"] = ".tsv",
        ["application/rtf"] = ".rtf",
        ["text/rtf"] = ".rtf",
        ["application/json"] = ".json",
        ["application/xml"] = ".xml",
        ["text/xml"] = ".xml",
        ["application/yaml"] = ".yaml",
        ["text/yaml"] = ".yaml",
        ["application/toml"] = ".toml",
        ["application/x-ipynb+json"] = ".ipynb",
        ["application/zip"] = ".zip",
        ["application/vnd.ms-htmlhelp"] = ".chm",
        ["image/vnd.djvu"] = ".djvu",
        ["application/x-mspublisher"] = ".pub",
        ["application/onenote"] = ".one",
        ["application/vnd.apple.pages"] = ".pages",
        ["application/vnd.apple.numbers"] = ".numbers",
        ["application/vnd.apple.keynote"] = ".key",
        ["image/png"] = ".png",
        ["image/jpeg"] = ".jpg",
        ["image/gif"] = ".gif",
        ["image/tiff"] = ".tif",
        ["image/bmp"] = ".bmp",
        ["image/webp"] = ".webp"
    };

    /// <summary>
    /// Returns the name the reader uses to select a handler. The media type wins when it implies an extension the
    /// identifier does not already end with; otherwise the identifier is used as is and content detection decides.
    /// </summary>
    internal static string ResolveSourceName(string? identifier, string? mediaType) {
        string name = string.IsNullOrWhiteSpace(identifier) ? "document" : identifier!.Trim();
        string? extension = GetExtension(mediaType);
        if (extension == null || name.EndsWith(extension, StringComparison.OrdinalIgnoreCase)) return name;
        return name + extension;
    }

    internal static string? GetExtension(string? mediaType) {
        if (string.IsNullOrWhiteSpace(mediaType)) return null;
        string value = mediaType!;
        int parameters = value.IndexOf(';');
        if (parameters >= 0) value = value.Substring(0, parameters);
        return Extensions.TryGetValue(value.Trim(), out string? extension) ? extension : null;
    }
}
