namespace OfficeIMO.Web.Converter.Models;

public enum ConversionInputKind {
    File,
    Text
}

public sealed record ConversionRoute(
    string Id,
    string Source,
    string Target,
    string Title,
    string Description,
    ConversionInputKind InputKind,
    string Accept,
    string EnginePath,
    string AccentClass);

public sealed record SelectedDocument(string Name, string Extension, string FormatLabel, long Size, byte[] Bytes);

public sealed record SampleDocument(string Label, string Path, string FileName, string Extension);

public sealed record ConversionDiagnostic(string Title, string Message, string ToneClass);

public sealed record ConversionResult(
    byte[] Bytes,
    string FileName,
    string ContentType,
    string? Text,
    string? HtmlPreview,
    IReadOnlyList<string> Warnings);
