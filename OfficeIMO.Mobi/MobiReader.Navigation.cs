using System.Globalization;
using System.Text;
using OfficeIMO.Epub;

namespace OfficeIMO.Mobi;

public static partial class MobiReader {
    private sealed partial class Projection {
        private List<Navigation> ReadNavigation(Kf8Content? kindle) {
            var result = new List<Navigation>();
            if (selected.NavigationIndex is not { } index) return result;
            try {
                var table = new MobiIndex(database, selected.Record + index, selected.Encoding, token);
                foreach (MobiIndex.Row row in table.Rows) {
                    token.ThrowIfCancellationRequested();
                    int level = row.Get(4) ?? 0;
                    if (level >= 32) {
                        Warn("MOBI_TOC_ENTRY_OMITTED", "The Mobipocket navigation depth exceeds 32 levels.", row.Name);
                        continue;
                    }
                    string? label = null;
                    if (row.Get(3) is { } labelOffset) {
                        try { label = XmlText(table.ReadLabel(labelOffset)); }
                        catch (Exception error) when (error is InvalidDataException or DecoderFallbackException) {
                            Warn("MOBI_TOC_ENTRY_OMITTED", error.Message, row.Name);
                        }
                    }
                    (int Part, int Offset)? target;
                    if (kindle == null && row.Get(1) is { } position) target = (0, position);
                    else if (kindle != null && row.Get(6) is { } file && row.Get(6, 1) is { } fragmentOffset)
                        target = kindle.Resolve(file, fragmentOffset);
                    else target = null;
                    if (string.IsNullOrWhiteSpace(label) || target == null) {
                        Warn("MOBI_TOC_ENTRY_OMITTED", "A Mobipocket navigation label or target is outside the supported profile.", row.Name);
                        result.Add(new Navigation(null, level, null));
                    } else {
                        result.Add(new Navigation(label, level, target));
                    }
                }
            } catch (Exception error) when (error is InvalidDataException or DecoderFallbackException) {
                Warn("MOBI_TOC_UNAVAILABLE", error.Message);
            }
            return result;
        }

        private IReadOnlyList<EpubNavigationEntry> BuildNavigation(List<Navigation> entries, MarkupPart[] parts) {
            var roots = new List<NavigationNode>();
            var ancestors = new List<(int Level, NavigationNode? Node)>();
            foreach (Navigation entry in entries) {
                token.ThrowIfCancellationRequested();
                while (ancestors.Count > 0 && ancestors[ancestors.Count - 1].Level >= entry.Level)
                    ancestors.RemoveAt(ancestors.Count - 1);
                NavigationNode? node = null;
                if (entry.Label is { } label && entry.Target is { } address) {
                    if (TryResolvePosition(address, parts, out string? target)) {
                        node = new NavigationNode(label, "EPUB/" + target);
                        NavigationNode? parent = ancestors.LastOrDefault(static item => item.Node != null).Node;
                        if (parent == null) roots.Add(node); else parent.Children.Add(node);
                    } else {
                        Warn("MOBI_TOC_ENTRY_OMITTED", "A Mobipocket navigation entry has no readable content target.", label);
                    }
                }
                // An omitted source entry still owns its depth, so its child cannot inherit a prior sibling.
                ancestors.Add((entry.Level, node));
            }
            return roots.Select(item => item.ToEntry()).ToArray();
        }

        private bool TryResolvePosition((int Part, int Offset) target, MarkupPart[] parts, out string? href) {
            href = null;
            if (target.Part < 0 || target.Part >= parts.Length) return false;
            MarkupPart part = parts[target.Part];
            if (target.Offset < 0 || target.Offset > part.Length) return false;
            href = part.File;
            if (part.AnchorAt(target.Offset) is { } anchor) href += "#" + anchor;
            return true;
        }

        private static (int Part, int Offset)? LinkPosition(string? href, string? filePosition, Kf8Content? kindle) {
            if (kindle == null)
                return int.TryParse(filePosition, NumberStyles.None, CultureInfo.InvariantCulture, out int position) ? (0, position) : null;
            const string prefix = "kindle:pos:fid:";
            if (href == null || !href.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return null;
            int separator = href.IndexOf(":off:", prefix.Length, StringComparison.OrdinalIgnoreCase);
            if (separator < 0 || !TryBase32(href.AsSpan(prefix.Length, separator - prefix.Length), out int file) ||
                !TryBase32(href.AsSpan(separator + 5), out int offset)) return null;
            return kindle.Resolve(file, offset);
        }

        private static bool TryBase32(ReadOnlySpan<char> value, out int result) {
            result = 0;
            if (value.Length == 0 || value.Length > 10) return false;
            foreach (char next in value) {
                int digit = next is >= '0' and <= '9' ? next - '0' : char.ToUpperInvariant(next) - 'A' + 10;
                if ((next < '0' || next > '9') && (char.ToUpperInvariant(next) < 'A' || char.ToUpperInvariant(next) > 'V')) return false;
                if (result > (int.MaxValue - digit) / 32) return false;
                result = result * 32 + digit;
            }
            return true;
        }
    }

    private sealed record Navigation(string? Label, int Level, (int Part, int Offset)? Target);
    private sealed record NavigationNode(string Label, string Target) {
        internal List<NavigationNode> Children { get; } = new();
        internal EpubNavigationEntry ToEntry() => new(Label, Target, Children.Select(child => child.ToEntry()));
    }
}
