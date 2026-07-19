namespace OfficeIMO.Email;

/// <summary>Format-neutral projection of Outlook reminder properties.</summary>
public sealed class OutlookReminder {
    /// <summary>Whether the reminder is enabled.</summary>
    public bool? IsSet { get; set; }
    /// <summary>Reminder lead time in minutes.</summary>
    public int? DeltaMinutes { get; set; }
    /// <summary>Reference time used to calculate the reminder.</summary>
    public DateTimeOffset? Time { get; set; }
    /// <summary>Next signal time.</summary>
    public DateTimeOffset? SignalTime { get; set; }
    /// <summary>Whether an item-specific reminder overrides the series reminder.</summary>
    public bool? Override { get; set; }
    /// <summary>Whether Outlook should play a sound.</summary>
    public bool? PlaySound { get; set; }
    /// <summary>Optional sound-file parameter retained from Outlook.</summary>
    public string? SoundFile { get; set; }
}
