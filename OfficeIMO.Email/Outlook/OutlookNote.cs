namespace OfficeIMO.Email;

/// <summary>Typed Outlook sticky-note fields.</summary>
public sealed class OutlookNote {
    /// <summary>Outlook note color numeric value.</summary>
    public int? Color { get; set; }
    /// <summary>Saved note window width.</summary>
    public int? Width { get; set; }
    /// <summary>Saved note window height.</summary>
    public int? Height { get; set; }
    /// <summary>Saved note window horizontal position.</summary>
    public int? X { get; set; }
    /// <summary>Saved note window vertical position.</summary>
    public int? Y { get; set; }
}
