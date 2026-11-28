using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public readonly partial struct WordLineSpacing {
    /// <summary>Reads an SDK spacing element; the rule is read as written, since EnumValue.Value throws on a token the SDK doesn't know.</summary>
    internal static WordLineSpacing? Read(SpacingBetweenLines? spacing) => Read(spacing?.Line?.Value, spacing?.LineRule?.InnerText);
}
