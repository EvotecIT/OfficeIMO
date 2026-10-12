namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkKeynoteArchiveBuilder {
    private const ulong ImageStyleId = 16;
    private bool _hasImageStyle;

    private void AddImageStyle() {
        _hasImageStyle = true;
        var stroke = P().Float(2, 0).Message(6, P().UInt(1, 2));
        var properties = P().Message(1, stroke).Float(2, 1)
            .Message(3, P().Bool(6, false)).Message(4, P());
        Add(ImageStyleId, 3016, P().Message(1, Style(ImageStyleId, "OfficeIMO Image"))
            .UInt(10, 4).Message(11, properties));
    }

    private ulong AddImage(ulong slideId, IWorkKeynoteImage image) =>
        base.AddImage(slideId, image.Data, image.Geometry, ImageStyleId);
}
