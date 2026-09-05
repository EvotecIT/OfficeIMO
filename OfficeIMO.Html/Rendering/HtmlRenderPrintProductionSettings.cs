namespace OfficeIMO.Html;

/// <summary>Print marks requested by CSS Paged Media.</summary>
[Flags]
public enum HtmlRenderPrintMarks {
    /// <summary>No printer marks.</summary>
    None = 0,
    /// <summary>Crop marks aligned with the resolved trim boundary.</summary>
    Crop = 1,
    /// <summary>Registration crosses centered on each trim edge.</summary>
    Cross = 2
}

/// <summary>Resolved per-page print-production geometry in CSS pixels.</summary>
public sealed class HtmlRenderPrintProductionSettings {
    internal HtmlRenderPrintProductionSettings(
        double bleed,
        double markArea,
        HtmlRenderPrintMarks marks) {
        Bleed = bleed;
        MarkArea = markArea;
        Marks = marks;
    }

    /// <summary>Bleed distance outside each trim edge.</summary>
    public double Bleed { get; }

    /// <summary>Reserved sheet area outside the bleed boundary for printer marks.</summary>
    public double MarkArea { get; }

    /// <summary>Requested crop and registration marks.</summary>
    public HtmlRenderPrintMarks Marks { get; }

    /// <summary>Inset from the sheet MediaBox to the trim boundary.</summary>
    public double TrimInset => Bleed + MarkArea;

    /// <summary>Inset from the sheet MediaBox to the bleed boundary.</summary>
    public double BleedInset => MarkArea;
}
