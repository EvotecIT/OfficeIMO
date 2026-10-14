using OfficeIMO.Visio;

namespace OfficeIMO.Reader.Visio;

internal static partial class VisioReaderAdapter {
    private static VisioDocument LoadForReader(string path, VisioLoadOptions? options, SourceMetadata source, CancellationToken cancellationToken) {
        using var stream = File.OpenRead(path);
        return LoadForReader(stream, path, options, source, cancellationToken);
    }
    private static VisioDocument LoadForReader(Stream stream, string sourceName, VisioLoadOptions? options, SourceMetadata source, CancellationToken cancellationToken) {
        string extension = Path.GetExtension(sourceName).ToLowerInvariant();
        bool binary = extension is ".vsd" or ".vss" or ".vst";
        if (!binary && stream.CanSeek) {
            long original = stream.Position;
            try {
                stream.Position = 0;
                byte[] signature = new byte[8];
                int read = 0;
                while (read < signature.Length) { int next = stream.Read(signature, read, signature.Length - read); if (next == 0) break; read += next; }
                binary = read == 8 && signature.SequenceEqual(new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 });
            } finally { stream.Position = original; }
        }
        if (binary) {
            var binaryOptions = new VisioLegacyBinaryImportOptions();
            if (options?.MaxInputBytes is long budget) binaryOptions.Limits.MaxInputBytes = (int)Math.Min(binaryOptions.Limits.MaxInputBytes, budget);
            VisioPackageType family = extension == ".vss" ? VisioPackageType.Stencil : extension == ".vst" ? VisioPackageType.Template : VisioPackageType.Drawing;
            var result = VisioDocument.LoadLegacyBinary(stream, family, binaryOptions, cancellationToken);
            source.ImportDiagnostics = result.Report.FidelityDiagnostics;
            return result.Value;
        }
        if (extension != ".vdx" && extension != ".vsx" && extension != ".vtx") return VisioDocument.Load(stream, options);
        VisioPackageType type = extension == ".vsx" ? VisioPackageType.Stencil : extension == ".vtx" ? VisioPackageType.Template : VisioPackageType.Drawing;
        var imported = VisioDocument.LoadLegacyXml(stream, type, options, cancellationToken);
        source.ImportDiagnostics = imported.Report.FidelityDiagnostics;
        return imported.Value;
    }
}
