#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Compression format used when reading or writing CSV files.
/// </summary>
public enum CsvCompressionType
{
    /// <summary>No compression is used.</summary>
    None,

    /// <summary>Infer compression from the file extension; unknown extensions are treated as uncompressed CSV.</summary>
    Auto,

    /// <summary>GZip compression, commonly used by <c>.csv.gz</c> and <c>.gzip</c> files.</summary>
    GZip,

    /// <summary>Raw Deflate compression, commonly used by <c>.deflate</c> files.</summary>
    Deflate,

    /// <summary>Brotli compression, commonly used by <c>.br</c> files.</summary>
    Brotli,

    /// <summary>ZLib compression, commonly used by <c>.zlib</c> files.</summary>
    ZLib
}
