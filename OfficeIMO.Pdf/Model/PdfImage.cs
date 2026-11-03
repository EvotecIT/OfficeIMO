namespace OfficeIMO.Pdf;

/// <summary>Describes an image added to the document flow.</summary>
public sealed class PdfImage {
    internal PdfImage(global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? altText) {
        Width = width; Height = height; AltText = altText;
    }

    /// <summary>Width of the image frame.</summary>
    public global::OfficeIMO.OfficeLength Width { get; }

    /// <summary>Height of the image frame.</summary>
    public global::OfficeIMO.OfficeLength Height { get; }

    /// <summary>Alternative text for a meaningful image; null for a decorative one.</summary>
    public string? AltText { get; }
}
