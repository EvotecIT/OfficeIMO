namespace OfficeIMO.Pdf;

internal sealed class ImageBlock : IPdfBlock {
    public byte[] Data { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfAlign Align { get; }
    public ImageBlock(byte[] data, double width, double height, PdfAlign align) { Data = data; Width = width; Height = height; Align = align; }
}

