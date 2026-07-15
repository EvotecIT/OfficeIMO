using System.Net;

namespace OfficeIMO.Reader.Subtitles;

internal static class SubtitleParser {
    internal static SubtitleParseResult Parse(
        string content,
        ReaderSubtitleOptions options,
        CancellationToken cancellationToken) {
        string normalized = (content ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        bool webVtt = lines.Length > 0 && IsWebVttSignature(lines[0]);
        var cues = new List<SubtitleCue>(Math.Min(lines.Length / 3, options.MaxCues));
        var warnings = new List<string>();
        int index = webVtt ? SkipWebVttHeader(lines) : 0;
        while (index < lines.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            while (index < lines.Length && string.IsNullOrWhiteSpace(lines[index])) index++;
            if (index >= lines.Length) break;
            if (webVtt && IsWebVttMetadataBlock(lines[index])) {
                index = SkipBlock(lines, index);
                continue;
            }
            if (cues.Count >= options.MaxCues) {
                AddWarning(warnings, "Subtitle cue parsing stopped at MaxCues.");
                break;
            }

            int blockStart = index;
            string? identifier = null;
            string timingLine = lines[index].Trim();
            if (!timingLine.Contains("-->", StringComparison.Ordinal)) {
                identifier = timingLine;
                index++;
                if (index >= lines.Length) break;
                timingLine = lines[index].Trim();
            }
            if (!TryParseTiming(timingLine, out TimeSpan start, out TimeSpan end)) {
                AddWarning(warnings, "Ignored a subtitle block with an invalid timing line at line " +
                    (blockStart + 1).ToString(CultureInfo.InvariantCulture) + ".");
                index = SkipBlock(lines, blockStart);
                continue;
            }

            index++;
            var text = new StringBuilder();
            while (index < lines.Length && !string.IsNullOrWhiteSpace(lines[index])) {
                if (text.Length > 0) text.AppendLine();
                text.Append(lines[index]);
                index++;
            }
            string cueText = text.ToString().Trim();
            if (options.StripCueMarkup) cueText = StripMarkup(cueText);
            if (string.IsNullOrWhiteSpace(cueText)) continue;
            bool truncated = cueText.Length > options.MaxCueCharacters;
            if (truncated) cueText = cueText.Substring(0, options.MaxCueCharacters);
            cues.Add(new SubtitleCue(
                identifier,
                start,
                end,
                cueText,
                blockStart + 1,
                Math.Max(blockStart + 1, index),
                truncated));
        }

        return new SubtitleParseResult(cues, warnings, webVtt ? "webvtt" : "srt");
    }

    internal static string FormatTimestamp(TimeSpan value) {
        long totalHours = value.Ticks / TimeSpan.TicksPerHour;
        return totalHours.ToString("D2", CultureInfo.InvariantCulture) + ":" +
            value.Minutes.ToString("D2", CultureInfo.InvariantCulture) + ":" +
            value.Seconds.ToString("D2", CultureInfo.InvariantCulture) + "." +
            value.Milliseconds.ToString("D3", CultureInfo.InvariantCulture);
    }

    private static int SkipWebVttHeader(string[] lines) {
        int index = 1;
        while (index < lines.Length && !string.IsNullOrWhiteSpace(lines[index])) index++;
        return index;
    }

    private static bool IsWebVttSignature(string value) {
        string signature = value.TrimStart('\uFEFF');
        return string.Equals(signature, "WEBVTT", StringComparison.OrdinalIgnoreCase) ||
            (signature.Length > "WEBVTT".Length &&
             signature.StartsWith("WEBVTT", StringComparison.OrdinalIgnoreCase) &&
             char.IsWhiteSpace(signature["WEBVTT".Length]));
    }

    private static bool IsWebVttMetadataBlock(string value) {
        string trimmed = value.Trim();
        return IsWebVttNote(trimmed) ||
            string.Equals(trimmed, "STYLE", StringComparison.Ordinal) ||
            string.Equals(trimmed, "REGION", StringComparison.Ordinal);
    }

    private static bool IsWebVttNote(string value) {
        return string.Equals(value, "NOTE", StringComparison.Ordinal) ||
            (value.Length > "NOTE".Length &&
             value.StartsWith("NOTE", StringComparison.Ordinal) &&
             char.IsWhiteSpace(value["NOTE".Length]));
    }

    private static int SkipBlock(string[] lines, int index) {
        while (index < lines.Length && !string.IsNullOrWhiteSpace(lines[index])) index++;
        return index;
    }

    private static bool TryParseTiming(string value, out TimeSpan start, out TimeSpan end) {
        start = default;
        end = default;
        int arrow = value.IndexOf("-->", StringComparison.Ordinal);
        if (arrow < 0) return false;
        string startText = value.Substring(0, arrow).Trim().Replace(',', '.');
        string endAndSettings = value.Substring(arrow + 3).Trim();
        int separator = endAndSettings.IndexOfAny(new[] { ' ', '\t' });
        string endText = (separator < 0 ? endAndSettings : endAndSettings.Substring(0, separator)).Replace(',', '.');
        return TryParseTimestamp(startText, out start) &&
            TryParseTimestamp(endText, out end) &&
            start >= TimeSpan.Zero && end >= start;
    }

    private static bool TryParseTimestamp(string value, out TimeSpan timestamp) {
        timestamp = default;
        string[] timeParts = value.Split(':');
        if (timeParts.Length is < 2 or > 3) return false;

        long hours = 0;
        string minutesText;
        string secondsAndMilliseconds;
        if (timeParts.Length == 3) {
            if (timeParts[0].Length < 2 ||
                !long.TryParse(timeParts[0], NumberStyles.None, CultureInfo.InvariantCulture, out hours)) {
                return false;
            }
            minutesText = timeParts[1];
            secondsAndMilliseconds = timeParts[2];
        } else {
            minutesText = timeParts[0];
            secondsAndMilliseconds = timeParts[1];
        }

        int decimalPoint = secondsAndMilliseconds.IndexOf('.');
        if (minutesText.Length != 2 ||
            decimalPoint != 2 ||
            secondsAndMilliseconds.Length != 6 ||
            !int.TryParse(minutesText, NumberStyles.None, CultureInfo.InvariantCulture, out int minutes) ||
            !int.TryParse(secondsAndMilliseconds.Substring(0, 2), NumberStyles.None, CultureInfo.InvariantCulture, out int seconds) ||
            !int.TryParse(secondsAndMilliseconds.Substring(3, 3), NumberStyles.None, CultureInfo.InvariantCulture, out int milliseconds) ||
            minutes > 59 || seconds > 59) {
            return false;
        }

        try {
            long ticks = checked(
                checked(hours * TimeSpan.TicksPerHour) +
                checked(minutes * TimeSpan.TicksPerMinute) +
                checked(seconds * TimeSpan.TicksPerSecond) +
                checked(milliseconds * TimeSpan.TicksPerMillisecond));
            timestamp = TimeSpan.FromTicks(ticks);
            return true;
        } catch (OverflowException) {
            return false;
        }
    }

    private static string StripMarkup(string value) {
        var text = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '<' && TryFindCueTagEnd(value, index, out int tagEnd)) {
                index = tagEnd;
                continue;
            }
            text.Append(value[index]);
        }
        return WebUtility.HtmlDecode(text.ToString()).Trim();
    }

    private static bool TryFindCueTagEnd(string value, int tagStart, out int tagEnd) {
        tagEnd = value.IndexOf('>', tagStart + 1);
        if (tagEnd < 0) return false;

        int nameStart = tagStart + 1;
        if (nameStart < tagEnd && value[nameStart] == '/') nameStart++;
        if (nameStart >= tagEnd) return false;

        if (!char.IsLetter(value[nameStart])) {
            string timestamp = value.Substring(nameStart, tagEnd - nameStart);
            return TryParseTimestamp(timestamp, out _);
        }

        int nameEnd = nameStart + 1;
        while (nameEnd < tagEnd && (char.IsLetterOrDigit(value[nameEnd]) || value[nameEnd] == '-')) nameEnd++;
        string name = value.Substring(nameStart, nameEnd - nameStart);
        return name.Equals("b", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("br", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("c", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("font", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("i", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("lang", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("rt", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("ruby", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("span", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("u", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("v", StringComparison.OrdinalIgnoreCase);
    }

    private static void AddWarning(List<string> warnings, string value) {
        if (warnings.Count < 100) warnings.Add(value);
    }
}

internal sealed class SubtitleParseResult {
    internal SubtitleParseResult(IReadOnlyList<SubtitleCue> cues, IReadOnlyList<string> warnings, string format) {
        Cues = cues;
        Warnings = warnings;
        Format = format;
    }

    internal IReadOnlyList<SubtitleCue> Cues { get; }
    internal IReadOnlyList<string> Warnings { get; }
    internal string Format { get; }
}

internal sealed class SubtitleCue {
    internal SubtitleCue(
        string? identifier,
        TimeSpan start,
        TimeSpan end,
        string text,
        int startLine,
        int endLine,
        bool truncated) {
        Identifier = identifier;
        Start = start;
        End = end;
        Text = text;
        StartLine = startLine;
        EndLine = endLine;
        Truncated = truncated;
    }

    internal string? Identifier { get; }
    internal TimeSpan Start { get; }
    internal TimeSpan End { get; }
    internal string Text { get; }
    internal int StartLine { get; }
    internal int EndLine { get; }
    internal bool Truncated { get; }
}
