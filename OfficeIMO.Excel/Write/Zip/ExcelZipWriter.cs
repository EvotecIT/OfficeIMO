using System.IO.Compression;
using OfficeIMO.SharedSource.Packaging;

namespace OfficeIMO.Excel;

/// <summary>The package surface over the shared <see cref="OfficeZipWriter"/>.</summary>
internal sealed class ExcelZipWriter : ExcelPackageArchive {
    internal ExcelZipWriter(Stream destination, DateTime modified) {
        Writer = new OfficeZipWriter(destination, modified);
    }

    internal OfficeZipWriter Writer { get; }

    internal override Stream OpenEntry(string name, CompressionLevel level, bool background = false) =>
        Writer.OpenEntry(name, level, background && ExcelExecutionPolicy.SupportsBackgroundProcessing);

    internal override bool TryWritePatchablePrefix(Stream entry, byte[] bytes, out object? prefix) {
        bool written = Writer.TryWritePatchablePrefix(entry, bytes, out OfficeZipWriter.PatchablePrefix? patchable);
        prefix = patchable;
        return written;
    }

    internal override void PatchPrefix(object prefix, byte[] bytes) =>
        Writer.PatchPrefix((OfficeZipWriter.PatchablePrefix)prefix, bytes);

    public override void Dispose() => Writer.Dispose();
}
