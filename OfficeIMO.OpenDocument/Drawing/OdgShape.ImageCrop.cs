namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    /// <summary>Image crop offsets in intrinsic image lengths, ordered top, right, bottom, left.</summary>
    /// <remarks>Negative offsets add empty space. Null removes the local override and exposes inherited cropping.
    /// Explicit zero insets disable inherited cropping. Access on a shape without an image throws.</remarks>
    public OdfInsets? Crop {
        get { EnsureImageCrop(); return ReadImageCrop(); }
        set { EnsureImageCrop(); WriteImageCrop(value); }
    }
    private void EnsureImageCrop() {
        if (!IsImage) throw new InvalidOperationException("Only image frames have editable image cropping.");
    }
}
