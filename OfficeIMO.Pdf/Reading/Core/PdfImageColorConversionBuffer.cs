namespace OfficeIMO.Pdf;

/// <summary>Owns reusable component buffers for one image color-conversion chain.</summary>
internal sealed class PdfImageColorConversionBuffer {
    internal PdfImageColorConversionBuffer(
        int componentCount,
        PdfImageColorConversionBuffer? alternate) {
        Components = new double[componentCount];
        Alternate = alternate;
    }

    internal double[] Components { get; }

    internal PdfImageColorConversionBuffer? Alternate { get; }
}
