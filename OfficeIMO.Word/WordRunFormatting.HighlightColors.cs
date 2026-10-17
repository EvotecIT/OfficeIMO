using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal static partial class WordRunFormatting {
    /// <summary>Returns the RGB highlight colour from ECMA-376 Part 1, 17.18.40; none and unknown values have no paint.</summary>
    internal static string? GetHighlightColorHex(HighlightColorValues? highlight) {
        if (!highlight.HasValue) return null;
        return ((IEnumValue)highlight.Value).Value?.Trim().ToLowerInvariant() switch {
            "black" => "000000", "blue" => "0000ff", "cyan" => "00ffff",
            "darkblue" => "00008b", "darkcyan" => "008b8b", "darkgray" => "a9a9a9",
            "darkgreen" => "006400", "darkmagenta" => "800080", "darkred" => "8b0000",
            "darkyellow" => "808000", "green" => "00ff00", "lightgray" => "d3d3d3",
            "magenta" => "ff00ff", "red" => "ff0000", "white" => "ffffff", "yellow" => "ffff00",
            _ => null
        };
    }
}
