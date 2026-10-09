namespace OfficeIMO.OpenDocument;

public abstract partial class OdfShape {
    internal OdfImageMirror? ReadImageMirror() => OdfImageMirrorCodec.Parse(ReadGraphicProperty(OdfNamespaces.Style + "mirror"));
    internal void WriteImageMirror(OdfImageMirror? mirror) {
        string? lexical = OdfImageMirrorCodec.Serialize(mirror);
        EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Style + "mirror", lexical);
    }
}
