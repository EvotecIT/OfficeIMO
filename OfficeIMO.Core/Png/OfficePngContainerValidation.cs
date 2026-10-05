using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>One request's validated PNG container, bound to the inspected byte array.</summary>
internal readonly struct OfficePngContainerValidation {
    private readonly byte[]? _bytes;

    private OfficePngContainerValidation(byte[] bytes, int frameCount) {
        _bytes = bytes;
        FrameCount = frameCount;
    }

    internal int FrameCount { get; }

    /// <summary>Requires the same input owned by the current synchronous decode request.</summary>
    internal bool IsFor(byte[] bytes) => ReferenceEquals(_bytes, bytes);

    internal static bool TryCreate(
        byte[] bytes,
        CancellationToken cancellationToken,
        out OfficePngContainerValidation validation) {
        validation = default;
        if (!OfficePngReader.TryGetFrameCount(bytes, cancellationToken, out int frameCount)) return false;
        validation = new OfficePngContainerValidation(bytes, frameCount);
        return true;
    }
}
