using System.Globalization;
using System.Security.Cryptography;

namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkKeynoteArchiveBuilder {
    private const ulong ImageStyleId = 16;
    private readonly List<(ulong DataId, ulong ImageId, string FileName, byte[] Bytes, int PixelWidth, int PixelHeight)> _imageResources = new();
    private bool _hasImageStyle;
    internal IEnumerable<(string FileName, byte[] Bytes)> ImageResources =>
        _imageResources.Select(resource => (resource.FileName, resource.Bytes));

    private void AddImageStyle() {
        _hasImageStyle = true;
        var stroke = P().Float(2, 0).Message(6, P().UInt(1, 2));
        var properties = P().Message(1, stroke).Float(2, 1)
            .Message(3, P().Bool(6, false)).Message(4, P());
        Add(ImageStyleId, 3016, P().Message(1, Style(ImageStyleId, "OfficeIMO Image"))
            .UInt(10, 4).Message(11, properties));
    }

    private ulong AddImage(ulong slideId, IWorkKeynoteImage image) {
        ulong dataId = Allocate(), imageId = Allocate();
        string fileName = "image-" + dataId.ToString(CultureInfo.InvariantCulture) + image.Extension;
        _imageResources.Add((dataId, imageId, fileName, image.EncodedBytes, image.PixelWidth, image.PixelHeight));
        IWorkGeometry frame = image.Geometry;
        var drawable = P().Message(1, P().Message(1, Point(frame.LeftPoints, frame.TopPoints))
            .Message(2, Size(frame.WidthPoints, frame.HeightPoints)).UInt(3, 3).Float(4, 0))
            .Reference(2, slideId).Bool(5, false).Bool(7, false);
        if (image.AccessibilityDescription != null) drawable.String(8, image.AccessibilityDescription);
        Add(imageId, 3005, P().Message(1, drawable).Reference(3, ImageStyleId)
            .Message(4, Size(frame.WidthPoints, frame.HeightPoints)).UInt(7, 0)
            .Message(9, Size(image.PixelWidth, image.PixelHeight)).DataReference(11, dataId).Bool(18, false));
        return imageId;
    }

    private void AddImageComponentReferences(IWorkProtoWriter component) {
        foreach (var resource in _imageResources) {
            _cancellationToken.ThrowIfCancellationRequested();
            component.Message(7, P().UInt(1, resource.DataId)
                .Message(2, P().UInt(1, resource.ImageId).UInt(2, 1)));
        }
    }

    private void AddImageMetadata(IWorkProtoWriter metadata) {
        foreach (var resource in _imageResources) {
            _cancellationToken.ThrowIfCancellationRequested();
            // Apple's DataInfo digest is a SHA-1 payload checksum, not a security decision.
            using var digest = SHA1.Create();
            // Extension 100 identifies TSD.ImageDataAttributes. An empty attributes
            // message cannot be instantiated by Keynote's native class registry.
            var attributes = P().Message(100, P().Message(1, Size(resource.PixelWidth, resource.PixelHeight))
                .Bool(3, false));
            metadata.Message(4, P().UInt(1, resource.DataId).Bytes(2, digest.ComputeHash(resource.Bytes))
                .String(3, resource.FileName).String(4, resource.FileName).Message(10, attributes));
            _cancellationToken.ThrowIfCancellationRequested();
        }
    }
}
