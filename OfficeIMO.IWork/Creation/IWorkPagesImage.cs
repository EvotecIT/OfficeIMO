using System.Threading;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>An immutable PNG or JPEG occupying its own inline Pages body paragraph.</summary>
public sealed class IWorkPagesImage : IWorkPagesBodyBlock {
    internal IWorkPagesImage(byte[] data, float width, float height, string? description, CancellationToken cancellationToken) {
        Data = new IWorkNativeImageData(data, description, cancellationToken);
        Geometry = new IWorkGeometry(0, 0, width, height, 0);
    }

    /// <summary>Gets the zero-offset frame in points. The complete image fills this frame without a crop or mask.</summary>
    public IWorkGeometry Geometry { get; }
    /// <summary>Gets the validated pixel width.</summary>
    public int PixelWidth => Data.PixelWidth;
    /// <summary>Gets the validated pixel height.</summary>
    public int PixelHeight => Data.PixelHeight;
    /// <summary>Gets image/png or image/jpeg.</summary>
    public string MediaType => Data.MediaType;
    /// <summary>Gets the encoded payload size in bytes.</summary>
    public int Length => Data.EncodedBytes.Length;
    /// <summary>Gets the optional native accessibility description.</summary>
    public string? AccessibilityDescription => Data.AccessibilityDescription;
    /// <summary>Returns a copy of the validated encoded payload.</summary>
    public byte[] GetBytes() => (byte[])Data.EncodedBytes.Clone();
    internal IWorkNativeImageData Data { get; }
}
