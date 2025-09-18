namespace OfficeIMO.Pdf;

public sealed class PdfOptions {
    public double PageWidth { get; set; } = 612; // Letter 8.5in * 72
    public double PageHeight { get; set; } = 792; // Letter 11in * 72
    public double MarginLeft { get; set; } = 72; // 1 in
    public double MarginRight { get; set; } = 72;
    public double MarginTop { get; set; } = 72;
    public double MarginBottom { get; set; } = 72;
    public PdfStandardFont DefaultFont { get; set; } = PdfStandardFont.Courier;
    public double DefaultFontSize { get; set; } = 11;
}

