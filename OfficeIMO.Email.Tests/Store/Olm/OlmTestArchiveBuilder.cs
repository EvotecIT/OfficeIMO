using System.IO.Compression;

namespace OfficeIMO.Email.Store.Tests.Olm;

internal sealed class OlmTestArchiveBuilder : IDisposable {
    private readonly MemoryStream _stream = new MemoryStream();
    private readonly ZipArchive _archive;

    internal OlmTestArchiveBuilder() {
        _archive = new ZipArchive(_stream, ZipArchiveMode.Create, leaveOpen: true);
    }

    internal OlmTestArchiveBuilder AddText(string path, string content) {
        return Add(path, Encoding.UTF8.GetBytes(content));
    }

    internal OlmTestArchiveBuilder AddDirectory(string path) {
        _archive.CreateEntry(path.TrimEnd('/') + "/");
        return this;
    }

    internal OlmTestArchiveBuilder Add(string path, byte[] content) {
        ZipArchiveEntry entry = _archive.CreateEntry(path, CompressionLevel.Optimal);
        using (Stream output = entry.Open()) output.Write(content, 0, content.Length);
        return this;
    }

    internal byte[] Build() {
        _archive.Dispose();
        return _stream.ToArray();
    }

    public void Dispose() {
        _archive.Dispose();
        _stream.Dispose();
    }
}
