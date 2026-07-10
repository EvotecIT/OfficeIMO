namespace OfficeIMO.AsciiDoc;

/// <summary>Substitution types in their mandated evaluation order.</summary>
public enum AsciiDocSubstitutionType {
    /// <summary>Special character handling.</summary>
    SpecialCharacters = 0,
    /// <summary>Inline quote and formatting recognition.</summary>
    Quotes,
    /// <summary>Document attribute reference expansion.</summary>
    Attributes,
    /// <summary>Typographic character replacements.</summary>
    Replacements,
    /// <summary>Inline macro recognition.</summary>
    Macros,
    /// <summary>Post replacements such as line breaks.</summary>
    PostReplacements
}

/// <summary>Immutable ordered substitution plan for a block.</summary>
public sealed class AsciiDocSubstitutionPlan {
    internal AsciiDocSubstitutionPlan(string group, IReadOnlyList<AsciiDocSubstitutionType> substitutions) {
        Group = group;
        Substitutions = substitutions;
    }

    /// <summary>Named default group or <c>custom</c>.</summary>
    public string Group { get; }

    /// <summary>Enabled substitutions in evaluation order.</summary>
    public IReadOnlyList<AsciiDocSubstitutionType> Substitutions { get; }

    /// <summary>Tests whether a substitution is enabled.</summary>
    public bool Contains(AsciiDocSubstitutionType substitution) => Substitutions.Contains(substitution);
}

/// <summary>Resolves block defaults and explicit <c>subs</c> overrides.</summary>
public static class AsciiDocSubstitutionResolver {
    private static readonly AsciiDocSubstitutionType[] Normal = {
        AsciiDocSubstitutionType.SpecialCharacters,
        AsciiDocSubstitutionType.Quotes,
        AsciiDocSubstitutionType.Attributes,
        AsciiDocSubstitutionType.Replacements,
        AsciiDocSubstitutionType.Macros,
        AsciiDocSubstitutionType.PostReplacements
    };

    private static readonly AsciiDocSubstitutionType[] Header = {
        AsciiDocSubstitutionType.SpecialCharacters,
        AsciiDocSubstitutionType.Attributes,
        AsciiDocSubstitutionType.Macros
    };

    private static readonly AsciiDocSubstitutionType[] Verbatim = { AsciiDocSubstitutionType.SpecialCharacters };

    /// <summary>Gets the effective ordered plan for a block.</summary>
    public static AsciiDocSubstitutionPlan GetPlan(AsciiDocBlock block) {
        if (block == null) throw new ArgumentNullException(nameof(block));
        string? custom = GetSubs(block);
        if (custom != null) return ParseOverride(custom);
        if (block is AsciiDocHeading || block is AsciiDocBlockTitle || block is AsciiDocAttributeEntry) {
            return new AsciiDocSubstitutionPlan("header", Header);
        }
        if (block is AsciiDocDelimitedBlock delimited) {
            if (delimited.Kind == AsciiDocDelimitedBlockKind.Listing || delimited.Kind == AsciiDocDelimitedBlockKind.Literal) {
                return new AsciiDocSubstitutionPlan("verbatim", Verbatim);
            }
            if (delimited.Kind == AsciiDocDelimitedBlockKind.Passthrough || delimited.Kind == AsciiDocDelimitedBlockKind.Comment) {
                return new AsciiDocSubstitutionPlan("none", Array.Empty<AsciiDocSubstitutionType>());
            }
        }
        if (block is AsciiDocLineComment) return new AsciiDocSubstitutionPlan("none", Array.Empty<AsciiDocSubstitutionType>());
        return new AsciiDocSubstitutionPlan("normal", Normal);
    }

    private static string? GetSubs(AsciiDocBlock block) {
        for (int index = block.AttributeLists.Count - 1; index >= 0; index--) {
            string? value = block.AttributeLists[index].Attributes.GetNamedValue("subs");
            if (value != null) return value;
        }
        return null;
    }

    private static AsciiDocSubstitutionPlan ParseOverride(string value) {
        string normalized = value.Trim();
        if (string.Equals(normalized, "normal", StringComparison.OrdinalIgnoreCase)) return new AsciiDocSubstitutionPlan("normal", Normal);
        if (string.Equals(normalized, "header", StringComparison.OrdinalIgnoreCase)) return new AsciiDocSubstitutionPlan("header", Header);
        if (string.Equals(normalized, "verbatim", StringComparison.OrdinalIgnoreCase)) return new AsciiDocSubstitutionPlan("verbatim", Verbatim);
        if (string.Equals(normalized, "none", StringComparison.OrdinalIgnoreCase)) return new AsciiDocSubstitutionPlan("none", Array.Empty<AsciiDocSubstitutionType>());

        var enabled = new HashSet<AsciiDocSubstitutionType>();
        string[] entries = normalized.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        for (int index = 0; index < entries.Length; index++) {
            string entry = entries[index].Trim();
            bool remove = entry.StartsWith("-", StringComparison.Ordinal);
            bool add = entry.StartsWith("+", StringComparison.Ordinal);
            if (remove || add) entry = entry.Substring(1).Trim();
            if (!TryParse(entry, out AsciiDocSubstitutionType type)) continue;
            if (remove) enabled.Remove(type);
            else enabled.Add(type);
        }
        AsciiDocSubstitutionType[] ordered = Normal.Where(enabled.Contains).ToArray();
        return new AsciiDocSubstitutionPlan("custom", ordered);
    }

    private static bool TryParse(string value, out AsciiDocSubstitutionType type) {
        switch (value.ToLowerInvariant()) {
            case "specialchars": type = AsciiDocSubstitutionType.SpecialCharacters; return true;
            case "quotes": type = AsciiDocSubstitutionType.Quotes; return true;
            case "attributes": type = AsciiDocSubstitutionType.Attributes; return true;
            case "replacements": type = AsciiDocSubstitutionType.Replacements; return true;
            case "macros": type = AsciiDocSubstitutionType.Macros; return true;
            case "post_replacements":
            case "post-replacements": type = AsciiDocSubstitutionType.PostReplacements; return true;
            default: type = default; return false;
        }
    }
}
