using System.Globalization;
using System.Security.Cryptography;

namespace OfficeIMO.IWork.Internal;

internal abstract partial class IWorkNativeCreationGraph {
    private readonly List<(ulong DataId, ulong ImageId, string FileName, byte[] Bytes, int PixelWidth, int PixelHeight)> _imageResources = new();
    internal IEnumerable<(string FileName, byte[] Bytes)> ImageResources =>
        _imageResources.Select(resource => (resource.FileName, resource.Bytes));

    protected ulong AddImage(ulong parentId, IWorkNativeImageData image, IWorkGeometry frame, ulong styleId) {
        ulong dataId = Allocate(), imageId = Allocate();
        string fileName = "image-" + dataId.ToString(CultureInfo.InvariantCulture) + image.Extension;
        _imageResources.Add((dataId, imageId, fileName, image.EncodedBytes, image.PixelWidth, image.PixelHeight));
        var drawable = P().Message(1, P().Message(1, Point(frame.LeftPoints, frame.TopPoints))
            .Message(2, Size(frame.WidthPoints, frame.HeightPoints)).UInt(3, 3).Float(4, 0))
            .Reference(2, parentId).Bool(5, false).Bool(7, false);
        AddBodyAttachmentWrap(drawable);
        if (image.AccessibilityDescription != null) drawable.String(8, image.AccessibilityDescription);
        Add(imageId, 3005, P().Message(1, drawable).Reference(3, styleId)
            .Message(4, Size(frame.WidthPoints, frame.HeightPoints)).UInt(7, 0)
            .Message(9, Size(image.PixelWidth, image.PixelHeight)).DataReference(11, dataId).Bool(18, false));
        return imageId;
    }

    protected void AddImageComponentReferences(IWorkProtoWriter component) {
        foreach (var resource in _imageResources) {
            _cancellationToken.ThrowIfCancellationRequested();
            component.Message(7, P().UInt(1, resource.DataId)
                .Message(2, P().UInt(1, resource.ImageId).UInt(2, 1)));
        }
    }

    protected void AddImageMetadata(IWorkProtoWriter metadata) {
        foreach (var resource in _imageResources) {
            _cancellationToken.ThrowIfCancellationRequested();
            // Apple's DataInfo digest is a SHA-1 payload checksum, not a security decision.
            using var digest = SHA1.Create();
            // Extension 100 identifies TSD.ImageDataAttributes to Apple's class registry.
            var attributes = P().Message(100, P().Message(1, Size(resource.PixelWidth, resource.PixelHeight))
                .Bool(3, false));
            metadata.Message(4, P().UInt(1, resource.DataId).Bytes(2, digest.ComputeHash(resource.Bytes))
                .String(3, resource.FileName).String(4, resource.FileName).Message(10, attributes));
            _cancellationToken.ThrowIfCancellationRequested();
        }
    }
}
