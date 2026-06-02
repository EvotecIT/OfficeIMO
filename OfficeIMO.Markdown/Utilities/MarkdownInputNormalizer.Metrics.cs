using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownInputNormalizer {
    private static string ExpandCollapsedMetricLines(string text) {
        if (string.IsNullOrEmpty(text)) {
            return text ?? string.Empty;
        }

        var newline = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0 ? "\r\n" : "\n";
        var current = text;

        while (true) {
            var afterStatus = StatusCollapsedLineRegex.Replace(
                current,
                match => match.Groups["lead"].Value + newline + "- " + match.Groups["rest"].Value);

            var afterBullets = BulletCollapsedLineRegex.Replace(
                afterStatus,
                match => match.Groups["lead"].Value + newline + "- " + match.Groups["rest"].Value.TrimStart());

            if (afterBullets == current) {
                return afterBullets;
            }

            current = afterBullets;
        }
    }

    private static string NormalizeLegacyMetricBulletLeads(string text) {
        if (string.IsNullOrEmpty(text)) {
            return text ?? string.Empty;
        }

        var spaced = LineStartMissingSpaceBeforeBoldBulletRegex.Replace(text, "${indent}- ");
        return SingleStarMetricBulletRegex.Replace(spaced, "${indent}- **");
    }

    private static string ConvertLegacyMetricMarkdown(string text) {
        if (string.IsNullOrEmpty(text)) {
            return text ?? string.Empty;
        }

        var statusNormalized = LegacyStatusSummaryRegex.Replace(
            text,
            match => {
                var indent = match.Groups["indent"].Value;
                var value = match.Groups["value"].Value.Trim();
                return value.Length == 0 ? indent + "Status" : indent + "Status **" + value + "**";
            });

        return LegacyBoldMetricBulletRegex.Replace(
            statusNormalized,
            match => {
                var indent = match.Groups["indent"].Value;
                var label = match.Groups["label"].Value.Trim();
                var value = match.Groups["value"].Value.Trim();
                if (value.Length == 0) {
                    return indent + label;
                }

                if (value.IndexOf("**", StringComparison.Ordinal) >= 0
                    || value.IndexOf('`') >= 0
                    || value.IndexOf("~~", StringComparison.Ordinal) >= 0
                    || value.IndexOf("==", StringComparison.Ordinal) >= 0) {
                    return indent + label + " " + value;
                }

                return indent + label + " **" + value + "**";
            });
    }

    private static string RepairMalformedMetricValueStrongRuns(string text) {
        if (string.IsNullOrEmpty(text) || text.IndexOf("**", StringComparison.Ordinal) < 0) {
            return text ?? string.Empty;
        }

        var hasCrLf = text.IndexOf("\r\n", StringComparison.Ordinal) >= 0;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var changed = false;

        for (var i = 0; i < lines.Length; i++) {
            var line = lines[i] ?? string.Empty;
            if (!TryRepairMalformedMetricValueStrongRunLine(line, out var repaired)
                || repaired.Equals(line, StringComparison.Ordinal)) {
                continue;
            }

            lines[i] = repaired;
            changed = true;
        }

        if (!changed) {
            return text;
        }

        var rebuilt = string.Join("\n", lines);
        return hasCrLf ? rebuilt.Replace("\n", "\r\n") : rebuilt;
    }

    private static bool TryRepairMalformedMetricValueStrongRunLine(string line, out string repaired) {
        repaired = line;
        var trimmedStart = line.TrimStart();
        if (!trimmedStart.StartsWith("- ", StringComparison.Ordinal)
            && !OrderedListLeadRegex.IsMatch(trimmedStart)) {
            return false;
        }

        repaired = OveropenedMetricValueStrongRegex.Replace(line, static match => {
            var value = match.Groups["value"].Value.Trim();
            return value.Length == 0
                ? match.Value
                : match.Groups["prefix"].Value + "**" + value + "**" + match.Groups["tail"].Value;
        });

        repaired = AdjacentMetricStrongValueRegex.Replace(repaired, static match => {
            var first = match.Groups["first"].Value.Trim();
            var second = match.Groups["second"].Value.Trim();
            if (first.Length == 0 || second.Length == 0) {
                return match.Value;
            }

            if (IsSymbolOnlyMetricValue(first)) {
                return match.Groups["prefix"].Value + first + " **" + second + "**" + match.Groups["tail"].Value;
            }

            return match.Groups["prefix"].Value
                   + "**"
                   + first
                   + "** **"
                   + second
                   + "**"
                   + match.Groups["tail"].Value;
        });

        repaired = MissingTrailingStrongMetricCloseRegex.Replace(repaired, static match => {
            var value = match.Groups["value"].Value.Trim();
            return value.Length == 0
                ? match.Value
                : match.Groups["prefix"].Value + "**" + value + "**" + match.Groups["tail"].Value;
        });

        return true;
    }

    private static bool IsSymbolOnlyMetricValue(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        foreach (var ch in value) {
            if (char.IsWhiteSpace(ch)) {
                continue;
            }

            if (char.IsLetterOrDigit(ch)) {
                return false;
            }
        }

        return true;
    }
}
