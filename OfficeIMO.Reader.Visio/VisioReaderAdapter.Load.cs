using OfficeIMO.Visio;

namespace OfficeIMO.Reader.Visio;

internal static partial class VisioReaderAdapter {
    private static VisioDocument LoadForReader(string path, VisioLoadOptions? options, SourceMetadata source, CancellationToken cancellationToken) {
        using var stream = File.OpenRead(path);
        return LoadForReader(stream, path, options, source, cancellationToken);
    }
    private static VisioDocument LoadForReader(Stream stream, string sourceName, VisioLoadOptions? options, SourceMetadata source, CancellationToken cancellationToken) {
        string extension = Path.GetExtension(sourceName).ToLowerInvariant();
        if (extension != ".vdx" && extension != ".vsx" && extension != ".vtx") return VisioDocument.Load(stream, options);
        VisioPackageType type = extension == ".vsx" ? VisioPackageType.Stencil : extension == ".vtx" ? VisioPackageType.Template : VisioPackageType.Drawing;
        var imported = VisioDocument.LoadLegacyXml(stream, type, options, cancellationToken);
        source.ImportDiagnostics = imported.Report.FidelityDiagnostics;
        return imported.Value;
    }
}
