namespace OfficeIMO.Email;

/// <summary>One ordered, forward-compatible RRULE part.</summary>
public sealed class IcsRecurrencePart {
    /// <summary>Creates a recurrence part.</summary>
    public IcsRecurrencePart(string name, string value) {
        Name = ContentLineSyntax.RequireToken(name, nameof(name));
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Part name such as FREQ, BYDAY, or a future registered extension.</summary>
    public string Name { get; set; }
    /// <summary>Raw part value.</summary>
    public string Value { get; set; }
}

/// <summary>Parsed recurrence rule that retains unknown rule parts in source order.</summary>
public sealed class IcsRecurrenceRule {
    private readonly List<IcsRecurrencePart> _parts = new List<IcsRecurrencePart>();

    /// <summary>Ordered recurrence rule parts.</summary>
    public IList<IcsRecurrencePart> Parts => _parts;

    /// <summary>FREQ value.</summary>
    public string? Frequency {
        get => GetValue("FREQ");
        set => SetValue("FREQ", value);
    }

    /// <summary>COUNT value when present and valid.</summary>
    public int? Count {
        get => TryGetPositiveInteger("COUNT");
        set => SetValue("COUNT", value?.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>INTERVAL value when present and valid.</summary>
    public int? Interval {
        get => TryGetPositiveInteger("INTERVAL");
        set => SetValue("INTERVAL", value?.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>Parses an RRULE value without discarding unknown parts.</summary>
    public static IcsRecurrenceRule Parse(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        var rule = new IcsRecurrenceRule();
        foreach (string segment in value.Split(';')) {
            int equals = segment.IndexOf('=');
            if (equals <= 0 || equals == segment.Length - 1)
                throw new FormatException("The recurrence rule contains a malformed part.");
            try {
                rule._parts.Add(new IcsRecurrencePart(
                    segment.Substring(0, equals), segment.Substring(equals + 1)));
            } catch (ArgumentException exception) {
                throw new FormatException("The recurrence rule contains an invalid part name.", exception);
            }
        }
        if (string.IsNullOrWhiteSpace(rule.Frequency)) throw new FormatException("The recurrence rule does not declare FREQ.");
        return rule;
    }

    /// <summary>Returns the first matching rule part value.</summary>
    public string? GetValue(string name) => _parts.FirstOrDefault(part =>
        string.Equals(part.Name, name, StringComparison.OrdinalIgnoreCase))?.Value;

    /// <summary>Replaces a rule part, or removes it when <paramref name="value"/> is null.</summary>
    public void SetValue(string name, string? value) {
        IcsRecurrencePart? first = null;
        for (int index = _parts.Count - 1; index >= 0; index--) {
            if (!string.Equals(_parts[index].Name, name, StringComparison.OrdinalIgnoreCase)) continue;
            if (first == null) first = _parts[index]; else _parts.RemoveAt(index);
        }
        if (value == null) {
            if (first != null) _parts.Remove(first);
        } else if (first != null) first.Value = value;
        else _parts.Add(new IcsRecurrencePart(name, value));
    }

    /// <summary>Serializes the recurrence rule in its current part order.</summary>
    public override string ToString() => string.Join(";", _parts.Select(part =>
        string.Concat(ContentLineSyntax.RequireToken(part.Name, nameof(part.Name)), "=", part.Value)));

    private int? TryGetPositiveInteger(string name) {
        string? value = GetValue(name);
        return int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) && parsed > 0
            ? parsed
            : (int?)null;
    }
}
