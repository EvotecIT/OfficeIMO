using System.Threading;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>An immutable positioned PNG or JPEG in a newly created Keynote presentation.</summary>
public sealed class IWorkKeynoteImage : IWorkKeynoteCreationDrawable {
    internal IWorkKeynoteImage(byte[] data, float left, float top, float width, float height,
        string? accessibilityDescription, CancellationToken cancellationToken) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        IWorkCreationValidation.ValidateFinite(left, nameof(left));
        IWorkCreationValidation.ValidateFinite(top, nameof(top));
        IWorkCreationValidation.ValidatePositive(width, nameof(width));
        IWorkCreationValidation.ValidatePositive(height, nameof(height));
        Data = new IWorkNativeImageData(data, accessibilityDescription, cancellationToken);
        Geometry = new IWorkGeometry(left, top, width, height, 0);
    }

    /// <summary>Gets the frame geometry in points. The full image fills this frame without a crop or mask.</summary>
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
    internal byte[] ContentHash => Data.ContentHash;
    internal long DecodedBytes => Data.DecodedBytes;
}
