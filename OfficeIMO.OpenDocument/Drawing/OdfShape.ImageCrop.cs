namespace OfficeIMO.OpenDocument;

public abstract partial class OdfShape {
    internal OdfInsets? ReadImageCrop() => OdfImageCropCodec.Parse(ReadGraphicProperty(OdfNamespaces.Fo + "clip"));
    internal void WriteImageCrop(OdfInsets? crop) {
        string? lexical = OdfImageCropCodec.Serialize(crop); // Validate before copy-on-write style mutation.
        EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "clip", lexical);
    }
}
