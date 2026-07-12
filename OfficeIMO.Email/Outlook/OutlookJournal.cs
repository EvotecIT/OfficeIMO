namespace OfficeIMO.Email;

/// <summary>Typed Outlook journal fields.</summary>
public sealed class OutlookJournal {
    /// <summary>Journal activity start.</summary>
    public DateTimeOffset? Start { get; set; }
    /// <summary>Journal activity end.</summary>
    public DateTimeOffset? End { get; set; }
    /// <summary>Journal activity duration in minutes.</summary>
    public int? DurationMinutes { get; set; }
    /// <summary>Journal activity type.</summary>
    public string? Type { get; set; }
    /// <summary>Journal activity type description.</summary>
    public string? TypeDescription { get; set; }
    /// <summary>Journal flags.</summary>
    public int? Flags { get; set; }
    /// <summary>Whether the tracked document was printed.</summary>
    public bool? DocumentPrinted { get; set; }
    /// <summary>Whether the tracked document was saved.</summary>
    public bool? DocumentSaved { get; set; }
    /// <summary>Whether the tracked document was routed.</summary>
    public bool? DocumentRouted { get; set; }
    /// <summary>Whether the tracked document was posted.</summary>
    public bool? DocumentPosted { get; set; }
}
