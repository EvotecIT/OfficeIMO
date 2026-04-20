namespace OfficeIMO.Markup;

internal sealed class OfficeMarkupResolvedTransition {
    private readonly Dictionary<string, string> _attributes;

    public OfficeMarkupResolvedTransition(string rawText, string? effect, string? resolvedIdentifier, bool hasArguments, Dictionary<string, string> attributes) {
        RawText = rawText ?? string.Empty;
        Effect = effect;
        ResolvedIdentifier = resolvedIdentifier;
        HasArguments = hasArguments;
        _attributes = attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    public string RawText { get; }
    public string? Effect { get; }
    public string? ResolvedIdentifier { get; }
    public bool HasArguments { get; }
    public IReadOnlyDictionary<string, string> Attributes => _attributes;
}

internal static class OfficeMarkupTransitionResolver {
    public static OfficeMarkupResolvedTransition Parse(string? transition) {
        var rawText = transition?.Trim() ?? string.Empty;
        if (rawText.Length == 0) {
            return new OfficeMarkupResolvedTransition(string.Empty, null, null, false, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
        }

        var tokens = rawText.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length == 0) {
            return new OfficeMarkupResolvedTransition(rawText, null, null, false, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
        }

        var effect = tokens[0].Trim();
        var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 1; index < tokens.Length; index++) {
            var token = tokens[index].Trim();
            if (token.Length == 0) {
                continue;
            }

            var separatorIndex = token.IndexOf('=');
            if (separatorIndex > 0 && separatorIndex < token.Length - 1) {
                var key = token.Substring(0, separatorIndex).Trim();
                var value = token.Substring(separatorIndex + 1).Trim().Trim('"', '\'');
                if (key.Length > 0) {
                    attributes[key] = value;
                }
            } else {
                attributes[token] = string.Empty;
            }
        }

        return new OfficeMarkupResolvedTransition(
            rawText,
            effect,
            ResolveIdentifier(effect, attributes),
            tokens.Length > 1,
            attributes);
    }

    private static string? ResolveIdentifier(string? effect, IReadOnlyDictionary<string, string> attributes) {
        var normalized = Normalize(effect);
        return normalized switch {
            "none" => "None",
            "fade" => "Fade",
            "wipe" => "Wipe",
            "cut" => "Cut",
            "flash" => "Flash",
            "prism" => "Prism",
            "morph" => "Morph",
            "blinds" => ResolveVerticalHorizontal(attributes, "Blinds"),
            "blindsvertical" => "BlindsVertical",
            "blindshorizontal" => "BlindsHorizontal",
            "comb" => ResolveVerticalHorizontal(attributes, "Comb"),
            "combvertical" => "CombVertical",
            "combhorizontal" => "CombHorizontal",
            "push" => ResolveDirectional(attributes, "Push"),
            "pushup" => "PushUp",
            "pushdown" => "PushDown",
            "pushleft" => "PushLeft",
            "pushright" => "PushRight",
            "warp" => ResolveInOut(attributes, "Warp"),
            "warpin" => "WarpIn",
            "warpout" => "WarpOut",
            "ferris" => ResolveLeftRight(attributes, "Ferris"),
            "ferrisleft" => "FerrisLeft",
            "ferrisright" => "FerrisRight",
            _ => null
        };
    }

    private static string? ResolveVerticalHorizontal(IReadOnlyDictionary<string, string> attributes, string prefix) {
        var value = GetAttribute(attributes, "direction", "dir", "orientation", "axis");
        return Normalize(value) switch {
            "vertical" or "vert" => prefix + "Vertical",
            "horizontal" or "horiz" => prefix + "Horizontal",
            _ => null
        };
    }

    private static string? ResolveDirectional(IReadOnlyDictionary<string, string> attributes, string prefix) {
        var value = GetAttribute(attributes, "direction", "dir");
        return Normalize(value) switch {
            "up" or "top" => prefix + "Up",
            "down" or "bottom" => prefix + "Down",
            "left" => prefix + "Left",
            "right" => prefix + "Right",
            _ => null
        };
    }

    private static string? ResolveInOut(IReadOnlyDictionary<string, string> attributes, string prefix) {
        var value = GetAttribute(attributes, "direction", "dir", "mode");
        return Normalize(value) switch {
            "in" => prefix + "In",
            "out" => prefix + "Out",
            _ => null
        };
    }

    private static string? ResolveLeftRight(IReadOnlyDictionary<string, string> attributes, string prefix) {
        var value = GetAttribute(attributes, "direction", "dir");
        return Normalize(value) switch {
            "left" => prefix + "Left",
            "right" => prefix + "Right",
            _ => null
        };
    }

    private static string? GetAttribute(IReadOnlyDictionary<string, string> attributes, params string[] names) {
        foreach (var name in names) {
            if (attributes.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)) {
                return value.Trim();
            }
        }

        return null;
    }

    private static string Normalize(string? value) =>
        new string((value ?? string.Empty).Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
}
