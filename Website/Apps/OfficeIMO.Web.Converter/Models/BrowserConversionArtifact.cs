namespace OfficeIMO.Web.Converter.Models;

public sealed record BrowserConversionArtifact(
    byte[] Bytes,
    string FileName,
    string ContentType);
