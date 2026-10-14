using System.Threading;
using OfficeIMO.Archives;

namespace OfficeIMO.Zip;

public static partial class ZipTraversal {
    /// <summary>Validates a seekable compressed source and physical entry count, restoring its position.
    /// Keep source bytes immutable until an archive opened over them has been read.</summary>
    public static void ValidateSource(Stream zipStream, ZipTraversalOptions? options = null, CancellationToken cancellationToken = default) =>
        ArchiveReader.ValidateZipSource(zipStream, Options(options), cancellationToken);
}
