using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Workflows;

internal static partial class CupsPdfPrinter {
    // CUPS printer-type bits (cups.h): black-and-white, color, duplex, collated copies.
    private const long TypeBlackAndWhite = 0x4, TypeColor = 0x8, TypeDuplex = 0x10, TypeCollate = 0x80;

    /// <summary>Builds capabilities from <c>lpoptions -p queue -l</c> choices and the <c>lpoptions -p queue</c> destination attributes.</summary>
    internal static PdfPrinterCapabilities ParseCapabilities(string printerName, string choices, string destination) {
        Dictionary<string, (IReadOnlyList<string> Values, string? Default)> options = ParseChoices(choices);
        long? type = null;
        Match typeMatch = Regex.Match(destination, @"(?:^|\s)printer-type=(\d+)", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
        if (typeMatch.Success && long.TryParse(typeMatch.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out long parsedType)) type = parsedType;

        var media = new List<PdfPrinterMedia>();
        PdfPrinterMedia? defaultMedia = null;
        if (TryGetOption(options, out var sizes, "PageSize", "media")) {
            foreach (string value in sizes.Values) {
                if (value.StartsWith("Custom.", StringComparison.OrdinalIgnoreCase)) continue;
                var item = new PdfPrinterMedia(value, value, PdfPrinterMediaNames.TryGetSize(value));
                media.Add(item);
                if (value == sizes.Default) defaultMedia = item;
            }
        }

        IReadOnlyList<PdfPaperSourceInfo> sources = [];
        foreach (string key in new[] { "InputSlot", "media-source" }) {
            if (!options.TryGetValue(key, out var slots)) continue;
            sources = slots.Values.Select(value => new PdfPaperSourceInfo(key + "=" + value, value)).ToArray();
            break;
        }

        bool? duplex = null;
        if (TryGetOption(options, out var duplexChoices, "Duplex", "sides"))
            duplex = duplexChoices.Values.Any(value => value.IndexOf("Duplex", StringComparison.OrdinalIgnoreCase) >= 0 || value.StartsWith("two-sided", StringComparison.OrdinalIgnoreCase));
        if (type is long duplexType && (duplexType & TypeDuplex) != 0) duplex = true;

        bool? color = null;
        if (type is long colorType && (colorType & (TypeColor | TypeBlackAndWhite)) != 0) color = (colorType & TypeColor) != 0;
        else if (TryGetOption(options, out var colorChoices, "ColorModel", "print-color-mode"))
            color = colorChoices.Values.Any(value => value.IndexOf("RGB", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("CMY", StringComparison.OrdinalIgnoreCase) >= 0 || value.Equals("color", StringComparison.OrdinalIgnoreCase));

        var resolutions = new List<PdfPrinterResolution>();
        if (TryGetOption(options, out var resolutionChoices, "Resolution", "printer-resolution")) {
            foreach (string value in resolutionChoices.Values) {
                Match match = Regex.Match(value, @"^(\d{2,5})(?:x(\d{2,5}))?dpi$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromSeconds(1));
                if (!match.Success) continue;
                int x = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                int y = match.Groups[2].Success ? int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) : x;
                resolutions.Add(new(x, y));
            }
        }

        return new PdfPrinterCapabilities {
            PrinterName = printerName,
            Media = media,
            DefaultMedia = defaultMedia,
            PaperSources = sources,
            SupportsDuplex = duplex,
            SupportsColor = color,
            // The CUPS filter chain collates copies whether or not the device reports the capability.
            SupportsCollation = true,
            MaximumCopies = null,
            Resolutions = resolutions.Distinct().ToArray()
        };
    }

    private static bool TryGetOption(Dictionary<string, (IReadOnlyList<string> Values, string? Default)> options,
        out (IReadOnlyList<string> Values, string? Default) option, params string[] keys) {
        foreach (string key in keys) if (options.TryGetValue(key, out option)) return true;
        option = default;
        return false;
    }

    /// <summary>Parses <c>Key/Display text: choice *default choice</c> lines, keeping only safe option keywords.</summary>
    internal static Dictionary<string, (IReadOnlyList<string> Values, string? Default)> ParseChoices(string output) {
        var options = new Dictionary<string, (IReadOnlyList<string>, string?)>(StringComparer.Ordinal);
        foreach (string line in output.Split('\n')) {
            int separator = line.IndexOf(':');
            if (separator <= 0) continue;
            string key = line.Substring(0, separator).Split('/')[0].Trim();
            if (key.Length == 0 || options.ContainsKey(key)) continue;
            var values = new List<string>();
            string? selected = null;
            foreach (string item in line.Substring(separator + 1).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
                bool isDefault = item.StartsWith("*", StringComparison.Ordinal);
                string choice = item.TrimStart('*');
                if (!IsSafeKeyword(choice) || values.Contains(choice)) continue;
                values.Add(choice);
                if (isDefault) selected = choice;
            }
            options[key] = (values, selected);
        }
        return options;
    }

    private static bool IsSafeKeyword(string value) => value.Length is > 0 and <= 128 &&
        value.All(character => OfficeRuntimeCompatibility.IsAsciiLetterOrDigit(character) || character is '-' or '_' or '.');
}
