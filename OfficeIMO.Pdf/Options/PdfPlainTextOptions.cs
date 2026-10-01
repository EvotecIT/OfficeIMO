namespace OfficeIMO.Pdf;

/// <summary>Layout and resource bounds for literal text. Text is never interpreted as markup.</summary>
public sealed class PdfPlainTextOptions {
    /// <summary>PDF settings, including fonts and page geometry. The default is 10-point Courier.</summary>
    public PdfOptions PdfOptions { get; set; } = new() { DefaultFont = PdfStandardFont.Courier, DefaultFontSize = 10 };
    /// <summary>Explicit input encoding, or null for BOM detection followed by strict UTF-8.</summary>
    public string? EncodingName { get; set; }
    /// <summary>Tab columns counted from the start of each source line, from 1 through 32.</summary>
    public int TabSize { get; set; } = 8;
    /// <summary>Maximum decoded and tab-expanded characters.</summary>
    public int MaximumCharacters { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum generated pages, including explicit form-feed page breaks.</summary>
    public int MaximumPages { get; set; } = 10_000;

    /// <summary>Creates independent validated settings.</summary>
    public PdfPlainTextOptions Clone() {
        if (PdfOptions == null) throw new ArgumentException("PDF settings are required.", nameof(PdfOptions));
        if (TabSize < 1 || TabSize > 32) throw new ArgumentOutOfRangeException(nameof(TabSize));
        if (MaximumCharacters < 1) throw new ArgumentOutOfRangeException(nameof(MaximumCharacters));
        if (MaximumPages < 1) throw new ArgumentOutOfRangeException(nameof(MaximumPages));
        return new PdfPlainTextOptions {
            PdfOptions = PdfOptions.Clone(), EncodingName = EncodingName, TabSize = TabSize,
            MaximumCharacters = MaximumCharacters, MaximumPages = MaximumPages
        };
    }
}
