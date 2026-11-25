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

    public override void Dispose() => Writer.Dispose();
}
