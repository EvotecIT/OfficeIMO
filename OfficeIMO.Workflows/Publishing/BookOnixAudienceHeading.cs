namespace OfficeIMO.Workflows;

/// <summary>A plain-text audience heading or code equivalent, with optional ONIX list 74 language.</summary>
public sealed record BookOnixAudienceHeading(string Text, string? LanguageCode = null);
