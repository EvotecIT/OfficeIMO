namespace OfficeIMO.Pdf;

internal sealed class SpacerBlock : IPdfBlock {
    public double Height { get; }

    public SpacerBlock(double height) {
        Guard.NonNegative(height, nameof(height));
        Height = height;
    }
}
