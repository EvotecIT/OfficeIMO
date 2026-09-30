namespace OfficeIMO.Rtf;

/// <summary>Tracks numbering for one story in document order. Instances that share a definition have independent counters.</summary>
public sealed class RtfListNumbering {
    private readonly RtfDocument _document;
    private readonly Dictionary<string, Dictionary<int, long>> _counters = new Dictionary<string, Dictionary<int, long>>(StringComparer.Ordinal);

    /// <summary>Creates numbering state for a document. Use a separate state for each independent header, footer or note story.</summary>
    public RtfListNumbering(RtfDocument document) => _document = document ?? throw new ArgumentNullException(nameof(document));

    /// <summary>Advances the paragraph's list and returns its marker, or null for a paragraph outside a list.</summary>
    public RtfListMarker? Next(RtfParagraph paragraph) {
        RtfListFormatting? formatting = _document.ResolveListFormatting(paragraph);
        if (formatting == null) {
            foreach (string key in _counters.Keys.Where(key => key.StartsWith("anonymous:", StringComparison.Ordinal)).ToArray()) _counters.Remove(key);
            return null;
        }
        if (!_counters.TryGetValue(formatting.Identity, out Dictionary<int, long>? counters)) {
            counters = new Dictionary<int, long>();
            _counters.Add(formatting.Identity, counters);
        }
        int index = formatting.LevelIndex;
        long value = counters.TryGetValue(index, out long previous) ? checked(previous + 1) : formatting.Level.StartAt ?? 1;
        if (!formatting.IsDefined && paragraph.ListText != null) {
            string authored = paragraph.ListText.ToPlainText().TrimStart();
            string leading = new string(authored.TakeWhile(char.IsDigit).ToArray());
            if (long.TryParse(leading, NumberStyles.Integer, CultureInfo.InvariantCulture, out long authoredValue)) value = authoredValue;
        }
        counters[index] = value;
        foreach (int deeper in counters.Keys.Where(level => level > index).ToArray()) {
            RtfListFormatting? child = ResolveLevel(paragraph, deeper);
            if (child?.Level.NoRestart != true) counters.Remove(deeper);
        }
        string template = formatting.Level.Text ?? (formatting.Level.Kind == RtfListKind.Bullet ? "\u2022" : "%" + (index + 1).ToString(CultureInfo.InvariantCulture) + ".");
        bool supported = true;
        var text = new StringBuilder();
        for (int position = 0; position < template.Length; position++) {
            if (template[position] == '%' && position + 1 < template.Length && template[position + 1] is >= '1' and <= '9') {
                int referenced = template[++position] - '1';
                RtfListFormatting? level = referenced == index ? formatting : ResolveLevel(paragraph, referenced);
                long number = counters.TryGetValue(referenced, out long current) ? current : level?.Level.StartAt ?? 1;
                int format = referenced < index && formatting.Level.LegalNumbering == true ? 0 : level?.Level.NumberFormatN ?? level?.Level.NumberFormat ?? 0;
                text.Append(FormatNumber(number, format, ref supported));
            } else text.Append(template[position]);
        }
        if (!formatting.IsDefined && paragraph.ListText != null) text = new StringBuilder(paragraph.ListText.ToPlainText().Trim());
        string separator = formatting.Level.FollowCharacter == RtfListLevelFollowCharacter.Nothing ? string.Empty
            : formatting.Level.FollowCharacter == RtfListLevelFollowCharacter.Space ? " " : "\t";
        return new RtfListMarker(formatting, value, text.ToString(), separator, supported);
    }

    private RtfListFormatting? ResolveLevel(RtfParagraph paragraph, int index) {
        RtfParagraph view = paragraph.CopyFormattingView();
        view.ListLevel = index;
        return _document.ResolveListFormatting(view);
    }

    private static string FormatNumber(long number, int format, ref bool supported) {
        if (format == 0) return number.ToString(CultureInfo.InvariantCulture);
        if (format == 22) return number.ToString("D2", CultureInfo.InvariantCulture);
        if (format == 5) {
            long lastTwo = number % 100;
            string suffix = lastTwo is >= 11 and <= 13 ? "th" : (number % 10) switch { 1 => "st", 2 => "nd", 3 => "rd", _ => "th" };
            return number.ToString(CultureInfo.InvariantCulture) + suffix;
        }
        if ((format == 3 || format == 4) && number > 0) {
            var letters = new StringBuilder();
            while (number > 0) {
                number--;
                letters.Insert(0, (char)('A' + number % 26));
                number /= 26;
            }
            string result = letters.ToString();
            return format == 4 ? result.ToLowerInvariant() : result;
        }
        if ((format == 1 || format == 2) && number is > 0 and <= 3999) {
            int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] symbols = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            var roman = new StringBuilder();
            for (int i = 0; i < values.Length; i++) while (number >= values[i]) { roman.Append(symbols[i]); number -= values[i]; }
            string result = roman.ToString();
            return format == 2 ? result.ToLowerInvariant() : result;
        }
        supported = false;
        return number.ToString(CultureInfo.InvariantCulture);
    }
}

/// <summary>One effective list marker, including its number, visible text and following separator.</summary>
public sealed class RtfListMarker {
    internal RtfListMarker(RtfListFormatting formatting, long value, string text, string separator, bool supported) {
        Formatting = formatting; Value = value; Text = text; Separator = separator; IsNumberFormatSupported = supported;
    }
    /// <summary>Effective list instance and level.</summary>
    public RtfListFormatting Formatting { get; }
    /// <summary>Current number at this level.</summary>
    public long Value { get; }
    /// <summary>Visible marker text without the following separator.</summary>
    public string Text { get; }
    /// <summary>Following tab, space, or empty string.</summary>
    public string Separator { get; }
    /// <summary>Whether each requested numeric format is supported. Unsupported formats use decimal text.</summary>
    public bool IsNumberFormatSupported { get; }
}
