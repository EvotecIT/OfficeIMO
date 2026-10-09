namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    /// <summary>Mirrors image pixels through the native graphic style without changing the resource.</summary>
    /// <remarks>Null removes the local override; <see cref="OdfImageMirror.None"/> disables inherited mirroring.
    /// Combine vertical with at most one horizontal mode. Draw projection supports unconditional horizontal
    /// mirroring; other legal native modes are preserved with unsupported mappings. Access on non-images throws.</remarks>
    public OdfImageMirror? Mirror {
        get { EnsureImageMirror(); return ReadImageMirror(); }
        set { EnsureImageMirror(); WriteImageMirror(value); }
    }
    private void EnsureImageMirror() {
        if (!IsImage) throw new InvalidOperationException("Only image frames have editable image mirroring.");
    }
}
