using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
    internal static class VisioMasterIdentity {
        internal static bool MatchesAny(VisioAssets.MasterInfo master, ISet<string>? filters) {
            if (filters == null || filters.Count == 0) {
                return true;
            }

            foreach (string key in GetKeys(master)) {
                if (filters.Contains(key)) {
                    return true;
                }
            }

            return false;
        }

        internal static IReadOnlyList<string> GetKeys(VisioAssets.MasterInfo master) {
            if (master == null) throw new ArgumentNullException(nameof(master));

            return new string?[] {
                    master.NameU,
                    master.Name,
                    master.Id,
                    master.RelationshipId,
                    ToSlug(master.NameU),
                    ToSlug(master.Name ?? string.Empty),
                    (master.Name ?? string.Empty).Replace(" ", "-")
                }
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        internal static string ToSlug(string value, string fallback = "") {
            if (string.IsNullOrWhiteSpace(value)) {
                return fallback;
            }

            StringBuilder builder = new(value.Length);
            bool previousDash = false;
            foreach (char character in value.Trim()) {
                if (char.IsLetterOrDigit(character)) {
                    builder.Append(char.ToLowerInvariant(character));
                    previousDash = false;
                } else if (!previousDash) {
                    builder.Append('-');
                    previousDash = true;
                }
            }

            string slug = builder.ToString().Trim('-');
            return string.IsNullOrWhiteSpace(slug) ? fallback : slug;
        }
    }
}
